import re
from utils import buscar_multiplo, limpar_moeda

def extrair(texto):
    # Regex da V2.2
    padroes = {
        # ... padrões existentes ...
    }
    
    dados = {
        'TIPO_CONTRATO': 'PADRÃO (QI)',
        # DADOS FIXOS DA QI (Pois nos contratos antigos isso varia pouco, mas vamos fixar para segurança)
        'NOME_INSTITUICAO': 'QI Sociedade de Crédito Direto S.A.',
        'CNPJ_INSTITUICAO': '32.402.502/0001-35',
        
        'NOME_CLIENTE': str(buscar_multiplo([r'Nome:\s*(.*?),', r'EMITENTE\s*Nome:\s*(.*?)(?=\sEndereço)'], texto)).title(),
        # ... resto da extração igual ...
        'CPF_CLIENTE': buscar_multiplo([r'CPF:\s*([\d\.-]+)'], texto),
        'NUMERO_CCB': buscar_multiplo([r'Cédula de Crédito Bancário nº\s*(DiDi\d+)\b', r'(DiDi\s*\d{4,10})\b'], texto),
        'VALOR_PRINCIPAL': buscar_multiplo([r'Valor Principal:\s*R\$\s*([\d\.,]+)'], texto), 
        'VALOR_LIBERADO': buscar_multiplo([r'Valor Liberado:\s*R\$\s*([\d\.,]+)'], texto),
        'DATA_EMISSAO_CCB': buscar_multiplo([r'Data de Emissão.*?: \s*(\d{2}/\d{2}/\d{4})'], texto), 
        'VENCIMENTO_FINAL': buscar_multiplo([r'Vencimento Final:\s*(\d{2}/\d{2}/\d{4})'], texto),
        'CET_ANUAL': buscar_multiplo([r'CET.*Anual:.*?([\d,]+)'], texto), 
        'VALOR_IOF': buscar_multiplo([r'IOF:\s*R\$\s*([\d\.,]+)'], texto),
        'TAXA_JUROS_MENSAL': buscar_multiplo([r'Juros.*?de\s*([\d,]+)%\s*a\.m\.'], texto), 
        'TAXA_JUROS_ANUAL': buscar_multiplo([r'equivalente.*?de\s*([\d,]+)%\s*a\.a\.'], texto),
        'CET_MENSAL': buscar_multiplo([r'CET.*Mensal:\s*([\d,]+)%'], texto), 
        'JUROS_MORA': buscar_multiplo([r'Juros Moratórios:\s*([\d,]+)%\s*a\.m\.'], texto),
        'MULTA_ATRASO': buscar_multiplo([r'Multa Moratória.*?:.*?([\d,]+)%'], texto), 
        'PRAZO_CONTRATO': buscar_multiplo([r'Prazo.*?:.*?(\d+\s+dias)'], texto),
        'ENDERECO_CLIENTE': buscar_multiplo([r'Endereço:\s*(.*?)(?=\sCidade:|\sCEP:)'], texto), 
        'CIDADE': buscar_multiplo([r'Cidade:\s*(.*?)(?=Estado:|CEP:|UF)'], texto),
        'ESTADO': buscar_multiplo([r'Estado:\s*([A-Z]{2}|[A-Za-z\s]+)'], texto), 
        'EMAIL_CLIENTE': buscar_multiplo([r'E-mail:\s*([^\s,]+)', r'([\w\.-]+@[\w\.-]+\.\w+)'], texto),
        'TEL_CLIENTE': buscar_multiplo([r'Telefone:\s*([\+\d\s\(\)-]+)', r'Tel\.:\s*([\+\d\s\(\)-]+)'], texto),
    }
    return dados