import os
import sys

class Config:
    # Identidade da Aplicacao
    APP_TITLE = "Hub de Documentos"
    VERSION = "v3.0"
    
    # Configuracoes de Atualizacao (Git)
    REPO_URL = "kevynlimai-bit/HubDocumentos"
    # SEU TOKEN NOVO (Read-Only) VAI AQUI:
    GITHUB_TOKEN = "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" 
    EXE_NAME = "Hub.exe"

    # Definicao de Caminhos Locais
    if getattr(sys, 'frozen', False):
        BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
    else:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    # Caminhos de Pastas
    DIR_ASSETS = os.path.join(BASE_DIR, "assets")
    DIR_TEMPLATES = os.path.join(BASE_DIR, "templates")
    DIR_OUTPUT = os.path.join(BASE_DIR, "output")
    DIR_TEMP = os.path.join(BASE_DIR, "temp")
    LOG_FILE = os.path.join(BASE_DIR, "debug_log.txt")
    
    # CAMINHO DE REDE (BPO)
    DRIVE_PATH = r"G:\Shared drives\DiDi   99\CX IBG\06. CX Fintech BR\02. Cashloan\Hub - Documentos Gerados"

    # Identidade Visual
    LOGO_PATH = os.path.join(DIR_ASSETS, "logo.png")
    FONT_FAMILY = "Arial" 

    # --- CORES DINÂMICAS (Light Mode, Dark Mode) ---
    COLORS = {
        # Amarelo 99
        "primary": "#F3C623",      
        "primary_hover": "#D4AC0D",
        
        # Verde Sucesso
        "success": ("#27AE60", "#2ECC71"),      
        "success_hover": ("#219150", "#27AE60"),
        
        # Vermelho Erro
        "danger": ("#E74C3C", "#FF5252"),       
        "danger_hover": "#C0392B",
        
        # Textos e Fundos
        "dark": ("#000000", "#FFFFFF"),       
        "text": ("#333333", "#E0E0E0"),       
        "text_light": ("#666666", "#AAAAAA"), 
        
        "bg_main": ("#F5F7F9", "#121212"),    
        "bg_card": ("#FFFFFF", "#212121"),    
        "border": ("#E0E0E0", "#404040"),     
        
        # --- A COR QUE ESTAVA FALTANDO ---
        "input_bg": ("#FAFAFA", "#2A2A2A"),   
        
        "disabled": ("#BDC3C7", "#4A4A4A")
    }

    # Regex Legado (Para contratos antigos)
    PATTERNS_LEGACY = {
        "nome": [r'Nome:\s*(.*?),', r'EMITENTE\s*Nome:\s*(.*?)(?=\sEndereço)'],
        "cpf": [r'CPF:\s*([\d\.-]+)'],
        "contrato_ccb": [r'parte integrante.*?Cédula de Crédito Bancário nº\s*(DiDi\d+)', r'(DiDi\s*\d+)'],
        "valor_principal": [r'Valor Principal:\s*R\$\s*([\d\.,]+)'],
        "data_emissao": [r'Data de Emissão.*?: \s*(\d{2}/\d{2}/\d{4})'],
        "iof": [r'IOF:\s*R\$\s*([\d\.,]+)']
    }