from config import Config
from utils import buscar_multiplo

def extrair(texto):
    """
    Motor legado para contratos PADRÃO da QI Sociedade de Crédito Direto.
    Baseado na versão v2.2 de produção.
    """
    
    # Mapeamento direto da função _extrair_padrao do main.py v2.2
    return {
        'TIPO_CONTRATO': 'PADRÃO',
        'NOME_INSTITUICAO': 'QI Sociedade de Crédito Direto S.A.',
        'CNPJ_INSTITUICAO': '32.402.502/0001-35',
        
        # Identificadores (Usando PATTERNS_LEGACY do config.py)
        'NOME_CLIENTE': str(buscar_multiplo(Config.PATTERNS_LEGACY["nome"], texto)).title(),
        'CPF_CLIENTE': buscar_multiplo(Config.PATTERNS_LEGACY["cpf"], texto),
        'NUMERO_CCB': buscar_multiplo(Config.PATTERNS_LEGACY["contrato_ccb"], texto),
        
        # Valores
        'VALOR_PRINCIPAL': buscar_multiplo(Config.PATTERNS_LEGACY["valor_principal"], texto),
        'VALOR_LIBERADO': buscar_multiplo([r'Valor Liberado:\s*R\$\s*([\d\.,]+)'], texto),
        'VALOR_IOF': buscar_multiplo(Config.PATTERNS_LEGACY["iof"], texto),
        
        # Datas
        'DATA_EMISSAO_CCB': buscar_multiplo(Config.PATTERNS_LEGACY["data_emissao"], texto),
        'VENCIMENTO_FINAL': buscar_multiplo([r'Vencimento Final:\s*(\d{2}/\d{2}/\d{4})'], texto),
        
        # Taxas e Encargos
        'CET_ANUAL': buscar_multiplo([r'CET.*Anual:.*?([\d,]+)'], texto),
        'CET_MENSAL': buscar_multiplo([r'CET.*Mensal:\s*([\d,]+)%'], texto),
        'TAXA_JUROS_MENSAL': buscar_multiplo([r'Juros.*?de\s*([\d,]+)%\s*a\.m\.'], texto),
        'TAXA_JUROS_ANUAL': buscar_multiplo([r'equivalente.*?de\s*([\d,]+)%\s*a\.a\.'], texto),
        'JUROS_MORA': buscar_multiplo([r'Juros Moratórios:\s*([\d,]+)%\s*a\.m\.'], texto),
        'MULTA_ATRASO': buscar_multiplo([r'Multa Moratória.*?:.*?([\d,]+)%'], texto),
        'PRAZO_CONTRATO': buscar_multiplo([r'Prazo.*?:.*?(\d+\s+dias)'], texto),
        
        # Dados Cadastrais
        'ENDERECO_CLIENTE': buscar_multiplo([r'Endereço:\s*(.*?)(?=\sCidade:|\sCEP:)'], texto),
        'CIDADE': buscar_multiplo([r'Cidade:\s*(.*?)(?=Estado:|CEP:|UF)'], texto),
        'ESTADO': buscar_multiplo([r'Estado:\s*([A-Z]{2}|[A-Za-z\s]+)'], texto),
        'EMAIL_CLIENTE': buscar_multiplo([r'E-mail:\s*([^\s,]+)'], texto),
        'TEL_CLIENTE': buscar_multiplo([r'Telefone:\s*([\+\d\s\(\)-]+)', r'Tel\.:\s*([\+\d\s\(\)-]+)'], texto),
        
        # Flag para pular a validação de contagem de parcelas (recurso exclusivo da V3)
        'QTD_PARCELAS_ENCONTRADAS': 0
    }