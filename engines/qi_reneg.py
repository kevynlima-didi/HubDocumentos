from config import Config
from utils import buscar_multiplo, converter_extenso_para_numero_seguro

def extrair(texto):
    """
    Motor legado para contratos de RENEGOCIAÇÃO da QI Sociedade de Crédito Direto.
    Baseado na versão v2.2 de produção.
    """
    
    # Lógica de CET por extenso (Preservada da v2.2)
    cet_raw = buscar_multiplo([r'Custo Efetivo Total \(CET\) Anual:\s*([^(\n]+)'], texto)
    val_cet = converter_extenso_para_numero_seguro(cet_raw)
    str_cet = f"{val_cet:.4f}".replace('.', ',') if val_cet > 0 else "0,00"

    # Lógica de Dívida Originária (Preservada da v2.2)
    contratos_origem = buscar_multiplo([r'1\.1\. Dívida\(s\) Originária\(s\).*?:\s*(.*?)(?=\s1\.2\.|$)'], texto)

    return {
        'TIPO_CONTRATO': 'RENEGOCIAÇÃO',
        'NOME_INSTITUICAO': 'QI Sociedade de Crédito Direto S.A.',
        'CNPJ_INSTITUICAO': '32.402.502/0001-35',

        # Identificadores
        'NOME_CLIENTE': str(buscar_multiplo(Config.PATTERNS_LEGACY["nome"], texto)).title(),
        'CPF_CLIENTE': buscar_multiplo(Config.PATTERNS_LEGACY["cpf"], texto),
        'NUMERO_CCB': contratos_origem if contratos_origem else "RENEG_GENERICA",
        
        # Valores
        'VALOR_PRINCIPAL': buscar_multiplo(Config.PATTERNS_LEGACY["valor_principal"], texto),
        'VALOR_LIBERADO': buscar_multiplo([r'Valor Liberado:\s*R\$\s*([\d\.,]+)'], texto),
        'VALOR_IOF': buscar_multiplo(Config.PATTERNS_LEGACY["iof"], texto),
        
        # Datas
        'DATA_EMISSAO_CCB': buscar_multiplo(Config.PATTERNS_LEGACY["data_emissao"], texto),
        'VENCIMENTO_FINAL': buscar_multiplo([r'Vencimento Final:\s*(\d{2}/\d{2}/\d{4})'], texto),
        
        # Taxas (Com conversão de extenso aplicada)
        'CET_ANUAL': str_cet,
        'CET_MENSAL': buscar_multiplo([r'CET.*Mensal:\s*([\d,]+)%'], texto),
        'TAXA_JUROS_MENSAL': buscar_multiplo([r'Juros.*?de\s*([\d,]+)%\s*a\.m\.'], texto),
        'TAXA_JUROS_ANUAL': buscar_multiplo([r'equivalente.*?de\s*([\d,]+)%\s*a\.a\.'], texto),
        'JUROS_MORA': buscar_multiplo([r'Juros Moratórios:\s*([\d,]+)%\s*a\.m\.'], texto),
        'MULTA_ATRASO': buscar_multiplo([r'Multa Moratória.*?:.*?([\d,]+)%'], texto),
        'PRAZO_CONTRATO': buscar_multiplo([r'Prazo.*?:.*?(\d+\s+dias)'], texto),
        
        # Dados Cadastrais
        'ENDERECO_CLIENTE': buscar_multiplo([r'Endereço residencial:\s*(.*?)(?=\sCidade:)', r'Endereço:\s*(.*?)(?=\sCidade:|\sCEP:)'], texto),
        'CIDADE': buscar_multiplo([r'Cidade:\s*(.*?)(?=Estado:|CEP:|UF)'], texto),
        'ESTADO': buscar_multiplo([r'Estado:\s*([A-Z]{2}|[A-Za-z\s]+)'], texto),
        'EMAIL_CLIENTE': buscar_multiplo([r'E-mail:\s*([^\s,]+)', r'([\w\.-]+@[\w\.-]+\.\w+)'], texto),
        'TEL_CLIENTE': buscar_multiplo([r'Telefone:\s*([\+\d\s\(\)-]+)', r'Tel\.:\s*([\+\d\s\(\)-]+)'], texto),
        
        # Flag para pular a validação de contagem de parcelas
        'QTD_PARCELAS_ENCONTRADAS': 0
    }