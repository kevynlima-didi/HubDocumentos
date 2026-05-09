from utils import limpar_valor
import pandas as pd

class BusinessValidators:
    """
    O Tribunal do Sistema.
    Centraliza todas as regras de negocio e Poka-yokes (Anti-Erro).
    """

    @staticmethod
    def validar_campos_obrigatorios(dados_pdf):
        """
        Verifica quais campos cruciais falharam na extracao.
        Retorna: Lista de nomes de campos para o UI destacar em vermelho.
        """
        campos_criticos = [
            'NOME_INSTITUICAO',
            'CNPJ_INSTITUICAO',
            'TIPO_CONTRATO',
            'NUMERO_CCB',
            'NOME_CLIENTE',
            'CPF_CLIENTE',
            'ENDERECO_CLIENTE',
            'CIDADE',
            'ESTADO',
            'TEL_CLIENTE',
            'EMAIL_CLIENTE',
            'VALOR_PRINCIPAL',
            'VALOR_LIBERADO',
            'VALOR_IOF',
            'DATA_EMISSAO_CCB',
            'VENCIMENTO_FINAL',
            'TAXA_JUROS_MENSAL',
            'TAXA_JUROS_ANUAL',
            'CET_MENSAL',
            'CET_ANUAL',
            'JUROS_MORA',
            'MULTA_ATRASO',
            'PRAZO_CONTRATO'
        ]
        
        campos_vazios = []
        for campo in campos_criticos:
            valor = dados_pdf.get(campo, "")
            if not valor or str(valor).strip() == "":
                campos_vazios.append(campo)
        
        return campos_vazios

    @staticmethod
    def validar_regras_financeiras(dados_pdf):
        """
        Validacoes matematicas internas do contrato (Sem o MIS).
        Retorna: (is_valid, lista_mensagens_erro)
        """
        erros = []
        
        # Regra 1: IOF vs Principal
        iof = limpar_valor(dados_pdf.get('VALOR_IOF', '0'))
        principal = limpar_valor(dados_pdf.get('VALOR_PRINCIPAL', '0'))
        
        if iof >= principal and principal > 0:
            erros.append(f"Erro Critico: IOF (R$ {iof}) e maior ou igual ao Principal (R$ {principal}).")

        return (len(erros) == 0), erros

    @staticmethod
    def validar_cruzamento_mis(df_mis, dados_pdf):
        """
        O GRANDE POKA-YOKE.
        Cruza a tabela do MIS (colada pelo operador) com os dados do PDF.
        """
        erros = []
        avisos = []

        # 1. Validacao Estrutural da Tabela
        if df_mis.empty:
            return False, ["A tabela do MIS esta vazia."], []

        try:
            # Identifica a linha de Total (ultima linha)
            linha_total = df_mis.iloc[-1]
            label_total = str(linha_total.iloc[0]).lower()
        except:
             return False, ["Erro ao ler a estrutura da tabela. Verifique a colagem."], []
        
        # Verifica se a ultima linha e realmente o Total
        # Aceita "Total", "total", "Soma", "soma" para flexibilidade
        if 'total' not in label_total and 'soma' not in label_total:
            erros.append("A ultima linha da tabela nao e 'Total'. Verifique a copia do MIS.")
            return False, erros, avisos

        # ==========================================================
        # NOVO POKA-YOKE (V3.0.5): VALIDAÇÃO DE IDENTIDADE (PRINCIPAL)
        # Garante que a tabela colada pertence ao mesmo contrato do PDF
        # ==========================================================
        val_mis = limpar_valor(linha_total.get('Principal', '0'))
        val_pdf = limpar_valor(dados_pdf.get('VALOR_PRINCIPAL', '0'))
        
        # Se o MIS vier zerado no principal (erro de colagem), tenta o Total Devido para nao travar a toa
        # Mas para Principal Contratado, a coluna 'Principal' é a ideal.
        
        diferenca = abs(val_mis - val_pdf)
        
        # Tolerancia de R$ 1.00 para arredondamentos
        if diferenca > 1.0:
            erros.append(f"Essa tabela não corresponde a CCB selecionada.\nPrincipal MIS: R$ {val_mis} | Principal PDF: R$ {val_pdf}")
            return False, erros, avisos

        # ==========================================================
        # POKA-YOKE 1: CONTAGEM DE PARCELAS (N_PDF = N_MIS - 1)
        # ==========================================================
        qtd_parcelas_mis = len(df_mis) - 1
        
        # Se o engine retornou 0, e um contrato legado, entao ignoramos essa validacao
        qtd_parcelas_pdf = int(dados_pdf.get('QTD_PARCELAS_ENCONTRADAS', 0))

        if qtd_parcelas_pdf > 0:
            if qtd_parcelas_mis != qtd_parcelas_pdf:
                # Alterado para Aviso se a diferenca for pequena, ou Erro se for grande?
                # Mantendo logica original de Erro
                erros.append(f"Divergencia de Prazos: O Contrato preve {qtd_parcelas_pdf} parcelas, mas o MIS tem {qtd_parcelas_mis}.")

        # ==========================================================
        # POKA-YOKE 2: PAGAMENTO FANTASMA
        # ==========================================================
        total_pago_mis = limpar_valor(linha_total.get('Total Pago', '0'))
        
        # Pega apenas as linhas de parcelas para olhar o status
        df_apenas_parcelas = df_mis.iloc[:-1]
        
        tem_parcela_paga = df_apenas_parcelas['Status'].str.contains('Pago', case=False, na=False).any()
        qtd_parcelas_pagas = int(df_apenas_parcelas['Status'].str.contains('Pago', case=False, na=False).sum())

        if total_pago_mis > 0 and not tem_parcela_paga:
            erros.append(f"Inconsistencia Grave: Valor Pago e R$ {total_pago_mis}, mas nenhuma parcela consta como 'Pago'.")

        if total_pago_mis > 0 and qtd_parcelas_pagas == 0:
            erros.append(f"Inconsistencia: Total Pago R$ {total_pago_mis}, mas zero parcelas com status 'Pago'. Verifique o MIS.")

        # ==========================================================
        # POKA-YOKE 3: COERENCIA DE SALDOS
        # ==========================================================
        total_devido_mis = limpar_valor(linha_total.get('Total Devido', '0'))
        
        if total_pago_mis > (total_devido_mis + 1.0):
            erros.append(f"Valor Pago (R$ {total_pago_mis}) e maior que o Total da Divida (R$ {total_devido_mis}).")

        is_valid = len(erros) == 0
        return is_valid, erros, avisos

    @staticmethod
    def validar_quitacao(df_mis, dados_pdf):
        """
        Poka-yoke para Carta de Quitacao.
        Verifica que todas as parcelas estao pagas e saldo devedor = 0.
        Retorna: (is_valid, lista_mensagens_erro)
        """
        erros = []

        if df_mis is None or df_mis.empty:
            return False, ["Cole a tabela do MIS para validar a quitação."]

        try:
            linha_total = df_mis.iloc[-1]
            label_total = str(linha_total.iloc[0]).lower()
        except:
            return False, ["Erro ao ler a tabela do MIS."]

        if 'total' not in label_total and 'soma' not in label_total:
            return False, ["A última linha da tabela não é 'Total'. Verifique a colagem."]

        df_parcelas = df_mis.iloc[:-1]

        # Verifica se todas as parcelas estao pagas
        qtd_total = len(df_parcelas)
        qtd_pagas = int(df_parcelas['Status'].str.contains('Pago', case=False, na=False).sum())
        qtd_aberto = qtd_total - qtd_pagas

        if qtd_aberto > 0:
            erros.append(f"Existem {qtd_aberto} parcela(s) em aberto de {qtd_total} total. A quitação só pode ser emitida com todas as parcelas pagas.")

        # Verifica saldo devedor
        total_devido = limpar_valor(linha_total.get('Total Devido', '0'))
        total_pago = limpar_valor(linha_total.get('Total Pago', '0'))
        saldo = total_devido - total_pago

        if saldo > 1.0:
            erros.append(f"Saldo devedor de R$ {saldo:,.2f}. A quitação só pode ser emitida com saldo zero.")

        # Verifica identidade da tabela
        val_mis = limpar_valor(linha_total.get('Principal', '0'))
        val_pdf = limpar_valor(dados_pdf.get('VALOR_PRINCIPAL', '0'))
        if abs(val_mis - val_pdf) > 1.0 and val_pdf > 0:
            erros.append(f"Essa tabela não corresponde à CCB selecionada.\nPrincipal MIS: R$ {val_mis} | Principal PDF: R$ {val_pdf}")

        return (len(erros) == 0), erros