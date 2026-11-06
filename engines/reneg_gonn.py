import re

from utils import buscar_padrao


def extrair(texto):
    """
    Motor para o layout de Instrumento Particular de Confissao e Renegociacao
    de Divida com uma credora representada por mandataria.
    """
    credora_match = re.search(
        r'E,\s*de outro lado,\s*(.*?)\s*,?\s*inscrit[oa]\s+no\s+CNPJ(?:/MF|/ME)?\s+sob\s+n[º°�?]?\s*([\d./-]+)',
        texto,
        re.IGNORECASE
    )
    nome_instituicao = credora_match.group(1).strip(' ,.;') if credora_match else ''
    cnpj_instituicao = credora_match.group(2).strip() if credora_match else ''

    contrato_origem = buscar_padrao([
        r'D[ií�?]vida\(s\) Origin[aá�?]ria\(s\).*?D[ií�?]vida Origin[aá�?]ria\s+([A-Za-z0-9_-]+)\s+R\$?\s*\d{1,3}(?:\.\d{3})*,\d{2}',
        r'\(N[º°�?]\s*do Contrato\).*?D[ií�?]vida Origin[aá�?]ria\s+([A-Za-z0-9_-]+)\s+R\$?\s*\d{1,3}(?:\.\d{3})*,\d{2}',
        r'D[ií�?]vida Origin[aá�?]ria\s+([A-Za-z0-9_-]+)\s+R\$?\s*\d{1,3}(?:\.\d{3})*,\d{2}',
        r'D[ií�?]vida Origin[aá�?]ria\s+([A-Za-z0-9_-]*\d[A-Za-z0-9_-]*)'
    ], texto)
    valor_repack = buscar_padrao([r'2\.1\..*?importa\s+em\s*R\s*\$?\s*([\d.,]+)'], texto)
    valor_originario = buscar_padrao([r'importa\s+em\s*R\s*\$?\s*([\d.,]+)', r'Valor Total a Pagar\s*R\s*\$?\s*([\d.,]+)', r'import[aâ�?]ncia\s*total\s*de\s*R\s*\$?\s*([\d.,]+)', r'Valor Principal\s*:?\s*R\s*\$?\s*([\d.,]+)'], texto)
    qtd_parcelas = buscar_padrao([r'em\s+(\d+)\s+parcelas\s+mensais', r'em\s+(\d+)\s+parcelas', r'parcelas\s+mensais\s*(\d+)', r'parcelas.*?mensais\s*(\d+)'], texto)

    return {
        'TIPO_CONTRATO': 'CONFISSÃO E RENEGOCIAÇÃO',
        'NOME_INSTITUICAO': nome_instituicao,
        'CNPJ_INSTITUICAO': cnpj_instituicao,

        'NOME_CLIENTE': str(buscar_padrao([r'Nome:\s*(.*?)(?:,\s*doravante|\s+doravante)'], texto)).title(),
        'CPF_CLIENTE': buscar_padrao([r'CPF:\s*([\d\.-]+)'], texto),
        'NUMERO_CCB': '',  # Confissão e Renegociação não tem próprio ID — operador preenche manualmente

        'VALOR_PRINCIPAL': valor_originario,
        'VALOR_LIBERADO': valor_originario,
        'VALOR_TOTAL_REPACTUADO': valor_repack,
        'VALOR_IOF': '0,00',

        'DATA_EMISSAO_CCB': buscar_padrao([r'Local e data:\s*[^,]+,\s*(\d{2}/\d{2}/\d{4})'], texto),
        'VENCIMENTO_FINAL': buscar_padrao([
            r'Data de vencimento.*?\s(\d{2}/\d{2}/\d{4})\s+R\$?\s*\d{1,3}(?:\.\d{3})*,\d{2}',
            r'D[ií�?]vida Origin[aá�?]ria\s+[A-Za-z0-9_-]*\d[A-Za-z0-9_-]*\s+R?\$?\s*\d{1,3}(?:\.\d{3})*,\d{2}\s+(\d{2}/\d{2}/\d{4})'
        ], texto),

        'CET_ANUAL': '0,00',
        'CET_MENSAL': '0,00',
        'TAXA_JUROS_MENSAL': '0,00',
        'TAXA_JUROS_ANUAL': '0,00',
        'JUROS_MORA': buscar_padrao([r'Juros de Mora\s*:?\s*([\d.,]+)\s*%', r'([\d.,]+)\s*%\s*\(um por cento\)\s*ao', r'([\d.,]+)\s*%.*?ao.*?sobre.*?atraso'], texto),
        'MULTA_ATRASO': buscar_padrao([r'([\d.,]+)\s*%\s*\(dois por cento\)', r'Multa Contratual\s*:?\s*([\d.,]+)\s*%', r'encargos\s*:?\s*([\d.,]+)\s*%'], texto),
        'PRAZO_CONTRATO': f'{qtd_parcelas} parcelas' if qtd_parcelas else buscar_padrao([r'em\s+(\d+)\s+parcelas', r'Prazo\s*:?\s*(\d+)\s+dias'], texto),

        'ENDERECO_CLIENTE': buscar_padrao([r'Endere[cç�?]o:\s*(.*?)(?=\s*Cidade:)'], texto).title(),
        'CIDADE': buscar_padrao([r'Cidade:\s*(.*?)(?=\s*Estado:)'], texto).title(),
        'ESTADO': buscar_padrao([r'Estado:\s*(.*?)(?=\s*CEP:)'], texto).title(),
        'EMAIL_CLIENTE': buscar_padrao([r'E-mail:\s*([^\s,]+)'], texto).lower(),
        'TEL_CLIENTE': buscar_padrao([r'Tel\.:\s*([\(\)\d\s-]+)'], texto),

        'QTD_PARCELAS_ENCONTRADAS': int(qtd_parcelas) if qtd_parcelas else 0
    }
