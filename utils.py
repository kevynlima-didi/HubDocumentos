import re

def limpar_valor(v):
    """
    Converte strings financeiras (R$ 1.000,00) para float (1000.00).
    Trata erros e valores vazios.
    """
    if not v: return 0.0
    v_str = str(v).strip().replace("R$", "").strip()
    if v_str in ["-", "", "N/A"]: return 0.0
    try:
        return float(v_str.replace('.', '').replace(',', '.'))
    except ValueError:
        return 0.0

def buscar_padrao(padroes, texto):
    """
    Itera sobre uma lista de regex e retorna o primeiro match encontrado.
    Ideal para documentos com layouts variados.
    """
    for padrao in padroes:
        match = re.search(padrao, texto, re.IGNORECASE)
        if match: 
            return match.group(1).strip()
    return ""

# ==============================================================================
#  LÓGICA MATEMÁTICA V2.2 (RESTAURADA PARA RENEGOCIAÇÃO)
# ==============================================================================

def processar_trecho_numerico(texto_fragmento, vocab, mults):
    """Auxiliar: Soma os valores compostos (Ex: cento + vinte = 120)"""
    total = 0
    atual = 0
    palavras = texto_fragmento.split()
    
    for palavra in palavras:
        if palavra in vocab:
            atual += vocab[palavra]
        elif palavra in mults:
            atual *= mults[palavra]
            total += atual
            atual = 0
            
    return total + atual

def converter_extenso_para_numero_seguro(texto):
    """
    Converte prazos/taxas escritos por extenso para números.
    Lógica robusta da v2.2 restaurada para suportar compostos e decimais.
    """
    if not texto: return 0.0
    
    # Vocabulário completo da v2.2
    vocabulario = {
        'zero': 0, 'um': 1, 'uma': 1, 'dois': 2, 'duas': 2, 'três': 3, 'tres': 3,
        'quatro': 4, 'cinco': 5, 'seis': 6, 'sete': 7, 'oito': 8, 'nove': 9,
        'dez': 10, 'onze': 11, 'doze': 12, 'treze': 13, 'quatorze': 14,
        'catorze': 14, 'quinze': 15, 'dezesseis': 16, 'dezessete': 17,
        'dezoito': 18, 'dezenove': 19, 'vinte': 20, 'trinta': 30,
        'quarenta': 40, 'cinquenta': 50, 'sessenta': 60, 'setenta': 70,
        'oitenta': 80, 'noventa': 90, 'cem': 100, 'cento': 100,
        'duzentos': 200, 'trezentos': 300, 'quatrocentos': 400,
        'quinhentos': 500, 'seiscentos': 600, 'setecentos': 700,
        'oitocentos': 800, 'novecentos': 900
    }
    
    multiplicadores = {
        'mil': 1000, 'milhão': 1000000, 'milhões': 1000000
    }
    
    decimais = {
        "décimos de milésimo": 10000, "milésimos": 1000, 
        "centésimos": 100, "décimos": 10, "inteiros": 1
    }

    # Limpeza
    texto_clean = str(texto).lower().replace(" e ", " ").replace(",", "").replace("por cento", "").replace("/", "").strip()
    
    # Lógica de Decimais (Ex: "dois inteiros e trinta centésimos")
    if "inteiros" in texto_clean:
        try:
            partes = texto_clean.split("inteiros")
            val_int = processar_trecho_numerico(partes[0], vocabulario, multiplicadores)
            
            divisor = 1
            txt_dec = partes[1]
            
            for nome_dec, v_div in decimais.items():
                if nome_dec in partes[1]:
                    divisor = v_div
                    txt_dec = partes[1].replace(nome_dec, "")
                    break
            
            val_dec = processar_trecho_numerico(txt_dec, vocabulario, multiplicadores)
            return float(val_int + (val_dec / divisor))
        except:
            return 0.0

    # Lógica Simples/Inteira
    val = processar_trecho_numerico(texto_clean, vocabulario, multiplicadores)
    
    # Fallback: Se deu zero mas tem dígitos no texto original (ex: "30 dias")
    if val == 0:
        digitos = re.sub(r'\D', '', str(texto))
        if digitos: return float(digitos)
        
    return float(val)

# ==============================================================================
#  RETROCOMPATIBILIDADE (ALIASES)
# ==============================================================================
buscar_multiplo = buscar_padrao
limpar_moeda = limpar_valor