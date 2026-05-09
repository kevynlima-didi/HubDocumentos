import os
import sys

# ==============================================================================
#  BLOCO DE INICIALIZAÇÃO DE AMBIENTE (GLOBAL)
#  Executa antes da classe existir para evitar o NameError
# ==============================================================================
if getattr(sys, 'frozen', False):
    # Caminho Interno (Dentro do .exe - Leitura do PyInstaller)
    SYS_INTERNAL = sys._MEIPASS 
    # Caminho Externo (Pasta onde o usuário clicou no .exe - Escrita)
    SYS_EXTERNAL = os.path.dirname(os.path.abspath(sys.executable))
else:
    # Modo Desenvolvimento
    SYS_INTERNAL = os.path.dirname(os.path.abspath(__file__))
    SYS_EXTERNAL = SYS_INTERNAL

def _get_safe_path(folder_name):
    """
    Helper interno para resolver caminhos.
    Procura dentro do .exe; se falhar, procura fora.
    """
    path_in = os.path.join(SYS_INTERNAL, folder_name)
    path_out = os.path.join(SYS_EXTERNAL, folder_name)
    
    # Prioridade: Usa o interno se existir (Build), senão usa o externo (Dev/Fallback)
    if os.path.exists(path_in):
        return path_in
    return path_out

class Config:
    # Identidade da Aplicacao
    APP_TITLE = "Hub de Documentos"
    VERSION = "v3.2" # Bug fixes & GitHub migration

    # Configuracoes de Atualizacao (Git)
    REPO_URL = "kevynlima-didi/HubDocumentos"
    GITHUB_TOKEN = ""
    EXE_NAME = "Hub.exe"

    # --- CAMINHOS (Usando as variáveis globais) ---
    DIR_ASSETS = _get_safe_path("assets")
    DIR_TEMPLATES = _get_safe_path("templates")

    # Caminhos de Escrita (SEMPRE Externos - Logs e Saída)
    DIR_OUTPUT = os.path.join(SYS_EXTERNAL, "output")
    DIR_TEMP = os.path.join(SYS_EXTERNAL, "temp")
    LOG_FILE = os.path.join(SYS_EXTERNAL, "debug_log.txt")
    
    # Caminho de Rede (Referência Base)
    DRIVE_PATH = r"G:\Shared drives\DiDi   99\CX IBG\06. CX Fintech BR\02. Cashloan\Hub - Documentos Gerados"

    # Identidade Visual
    LOGO_PATH = os.path.join(DIR_ASSETS, "logo.png")
    FONT_FAMILY = "Arial" 

    # --- CORES DINÂMICAS ---
    COLORS = {
        "primary": "#F3C623",      
        "primary_hover": "#D4AC0D",
        "success": ("#27AE60", "#2ECC71"),      
        "success_hover": ("#219150", "#27AE60"),
        "danger": ("#E74C3C", "#FF5252"),       
        "danger_hover": "#C0392B",
        "dark": ("#000000", "#FFFFFF"),       
        "text": ("#333333", "#E0E0E0"),       
        "text_light": ("#666666", "#AAAAAA"), 
        "bg_main": ("#F5F7F9", "#121212"),    
        "bg_card": ("#FFFFFF", "#212121"),    
        "border": ("#E0E0E0", "#404040"),     
        "input_bg": ("#FAFAFA", "#2A2A2A"),   
        "disabled": ("#BDC3C7", "#4A4A4A")
    }

    PATTERNS_LEGACY = {
        "nome": [r'Nome:\s*(.*?),', r'EMITENTE\s*Nome:\s*(.*?)(?=\sEndereço)'],
        "cpf": [r'CPF:\s*([\d\.-]+)'],
        "contrato_ccb": [r'parte integrante.*?Cédula de Crédito Bancário nº\s*(DiDi\d+)\b', r'(DiDi\s*\d{4,10})\b'],
        "valor_principal": [r'Valor Principal:\s*R\$\s*([\d\.,]+)'],
        "data_emissao": [r'Data de Emissão.*?: \s*(\d{2}/\d{2}/\d{4})'],
        "iof": [r'IOF:\s*R\$\s*([\d\.,]+)']
    }

    @staticmethod
    def encontrar_caminho_rede():
        """Scanner de Drive de Rede (G: até Z:)"""
        caminho_relativo = r"Shared drives\DiDi   99\CX IBG\06. CX Fintech BR\02. Cashloan\Hub - Documentos Gerados"
        
        # 1. Tenta a configuração padrão (G:)
        default_path = r"G:\Shared drives\DiDi   99\CX IBG\06. CX Fintech BR\02. Cashloan\Hub - Documentos Gerados"
        if os.path.exists(default_path): return default_path
        
        # 2. Se falhar, procura em outras letras
        for letra in "HIJKLMNOPQRSTUVWXYZ":
            candidato = f"{letra}:\\{caminho_relativo}"
            if os.path.exists(candidato):
                return candidato
        return None