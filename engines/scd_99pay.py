import re
from utils import buscar_padrao

def extrair(texto):
    """
    Motor de Extração v3.0.3 - SCD 99Pay (Estratégia Cirúrgica)
    - Partitioning: Corta o texto antes do QUADRO II para dados do cliente.
    - Segurança: Elimina risco de captura de dados da 99Pay (Quadro I).
    - Foco: Regex simplificados atuando apenas na área de interesse.
    """
    dados = {}

    # =========================================================================
    # 1. ESCOPO GLOBAL (Busca no texto inteiro)
    # Motivo: O cabeçalho (CCB) e o Rodapé (Credor) podem estar em extremos.
    # =========================================================================

    # --- DADOS DA INSTITUIÇÃO (CREDOR - QUADRO V) ---
    dados['NOME_INSTITUICAO'] = buscar_padrao([
        r'QUADRO V.*?CREDOR\s*(99\s*SCD\s*SOCIEDADE.*?S\.?A\.?)',
        r'QUADRO V.*?CREDOR\s*(99PAY\s*SOCIEDADE.*?S\.?A\.?)',
        r'(99PAY\s*SOCIEDADE\s*DE\s*CRÉDITO\s*DIRETO\s*S\.?A\.?)'
    ], texto).strip()

    cnpj_raw = buscar_padrao([
        r'QUADRO V.*?CNPJ.*?(\d{2}\.\d{3}\.\d{3}\s*/\s*\d{4}-\d{2})',
        r'CREDOR.*?CNPJ.*?(\d{2}\.\d{3}\.\d{3}\s*/\s*\d{4}-\d{2})'
    ], texto)
    dados['CNPJ_INSTITUICAO'] = cnpj_raw.replace(' ', '')

    # Fallback seguro
    if not dados['NOME_INSTITUICAO']: dados['NOME_INSTITUICAO'] = "99 SCD SOCIEDADE DE CRÉDITO DIRETO S.A."
    if not dados['CNPJ_INSTITUICAO']: dados['CNPJ_INSTITUICAO'] = "59.379.565/0001-74"

    # --- DADOS DO CONTRATO (CCB) ---
    dados['NUMERO_CCB'] = buscar_padrao([
        r'CÉDULA DE CRÉDITO BANCÁRIO N[º°].?\s*([a-zA-Z0-9]{5,})(?=\s)',
        r'\(CCB\):\s*([a-zA-Z0-9]+)(?=\s)'
    ], texto)
    dados['TIPO_CONTRATO'] = "99PAY SCD v1.0"

    # =========================================================================
    # 2. ESCOPO LOCAL (Partitioning: Apenas do QUADRO II em diante)
    # Motivo: Blindar contra o endereço da 99Pay no QUADRO I.
    # =========================================================================
    
    if "QUADRO II" in texto:
        _, texto_cliente = texto.split("QUADRO II", 1)
    else:
        texto_cliente = texto # Fallback se OCR falhar

    # --- DADOS DO CLIENTE ---
    dados['NOME_CLIENTE'] = buscar_padrao([
        r'Nome:\s*(.*?)(?:,|doravante)', 
        r'Nome completo:\s*(.*?)(?=\sCPF)'
    ], texto_cliente).title()
    
    dados['CPF_CLIENTE'] = buscar_padrao([r'CPF:\s*([\d\.-]+)'], texto_cliente)
    
    # --- ENDEREÇO (Regex Simples e Cirúrgico) ---
    # Como cortamos o texto, o primeiro "Endereço:" agora É O DO CLIENTE.
    regex_endereco = r'Endereço:\s*(.*?)(?=Cidade:|Estado:|CEP:|Tel|QUADRO)'
    
    end_match = re.search(regex_endereco, texto_cliente, re.IGNORECASE | re.DOTALL)
    if end_match:
        # Remove quebras de linha e espaços extras
        dados['ENDERECO_CLIENTE'] = end_match.group(1).replace('\n', ' ').strip().title()
    else:
        dados['ENDERECO_CLIENTE'] = ""

    dados['CIDADE'] = buscar_padrao([r'Cidade:\s*(.*?)(?=Tel|Estado|UF)'], texto_cliente).title()
    dados['ESTADO'] = buscar_padrao([r'Estado:\s*(.*?)(?=Fax|CEP|RG|CPF)'], texto_cliente)
    dados['TEL_CLIENTE'] = buscar_padrao([r'Tel\.:\s*([\(\)\d\s-]+)', r'Telefone:\s*([\(\)\d\s-]+)'], texto_cliente)

    # Tratamento especial de E-mail (evita lixo de OCR)
    raw_email_match = re.search(r'E-mail:(.*?)(?:\(\*\)|CPF:)', texto_cliente, re.DOTALL | re.IGNORECASE)
    email_final = ""
    if raw_email_match:
        fragmento = raw_email_match.group(1)
        fragmento = re.sub(r'Estado\s*Civil:.*', '', fragmento, flags=re.IGNORECASE | re.DOTALL)
        email_limpo = fragmento.replace('\n', '').replace(' ', '').strip()
        match_valid = re.search(r'([\w\.-]+@[\w\.-]+\.(?:com|br|net|org)(?:\.br)?)', email_limpo, re.IGNORECASE)
        if match_valid: email_final = match_valid.group(1).lower()
    
    if not email_final: email_final = buscar_padrao([r'E-mail:\s*([^\s]+)', r'([\w\.-]+@[\w\.-]+\.\w+)'], texto_cliente).lower()
    dados['EMAIL_CLIENTE'] = email_final

    # =========================================================================
    # 3. VALORES E TAXAS (Busca Global ou Local, conforme conveniência)
    # =========================================================================
    
    dados['VALOR_PRINCIPAL'] = buscar_padrao([r'Valor Principal\*?:\s*R\$\s*([\d\.,]+)'], texto)
    dados['VALOR_LIBERADO'] = buscar_padrao([r'Valor Liberado:\s*R\$\s*([\d\.,]+)', r'Valor Total Liberado\s*R\$\s*([\d\.,]+)'], texto)
    dados['VALOR_IOF'] = buscar_padrao([r'IOF:\s*R\$\s*([\d\.,]+)'], texto)

    dados['DATA_EMISSAO_CCB'] = buscar_padrao([r'Data de Emissão.*?:\s*(\d{2}/\d{2}/\d{4})'], texto)
    dados['VENCIMENTO_FINAL'] = buscar_padrao([r'Vencimento Final:\s*(\d{2}/\d{2}/\d{4})', r'vencimento da última parcela:\s*(\d{2}/\d{2}/\d{4})'], texto)
    dados['PRAZO_CONTRATO'] = buscar_padrao([r'Prazo:\s*(\d+\s+dias)'], texto)

    dados['TAXA_JUROS_MENSAL'] = buscar_padrao([r'Juros pré-fixados de\s*([\d,]+)\s*%'], texto)
    dados['TAXA_JUROS_ANUAL'] = buscar_padrao([r'equivalente à taxa de\s*([\d,]+)\s*%'], texto)
    dados['CET_MENSAL'] = buscar_padrao([r'\(CET\) Mensal:\s*([\d,]+)\s*%'], texto)
    dados['CET_ANUAL'] = buscar_padrao([r'\(CET\) Anual:\s*([\d,]+)\s*%'], texto)
    dados['JUROS_MORA'] = buscar_padrao([r'Juros Moratórios:\s*([\d,]+)\s*%'], texto)
    dados['MULTA_ATRASO'] = buscar_padrao([r'Multa.*?:\s*([\d,]+)\s*%'], texto)

    return dados