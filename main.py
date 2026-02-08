import customtkinter as ctk
import pandas as pd
# --- MATPLOTLIB ---
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt
# ------------------
import os
import re
import sys
import logging
import threading
import pythoncom
import shutil
import win32com.client as win32
import requests 
import subprocess 
import csv 
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from pypdf import PdfReader
from tkinter import messagebox
from docx.shared import Inches
from PIL import Image

# IMPORTS MODULARES
from config import Config
from utils import limpar_valor, buscar_padrao
from pokayokes import BusinessValidators 
from engines import qi_standard, qi_reneg, scd_99pay

# ==============================================================================
#  1. ATUALIZADOR (TURBINADO COM CHANGELOG)
# ==============================================================================
class SelfUpdater:
    def __init__(self, current_version, repo_url, token, exe_name="Hub.exe"):
        self.current_version = current_version
        self.repo_url = repo_url
        self.token = token 
        self.exe_name = exe_name
        self.api_url = f"https://api.github.com/repos/{repo_url}/releases/latest"
        self.download_url = None
        self.new_version = None
        self.release_notes = "" # Armazena o texto da novidade

    def check_for_updates(self):
        try:
            headers = {"Authorization": f"token {self.token}", "Accept": "application/vnd.github.v3+json"}
            response = requests.get(self.api_url, headers=headers, timeout=5)
            if response.status_code == 200:
                data = response.json()
                self.new_version = data['tag_name']
                # Captura o texto do release (O Changelog)
                self.release_notes = data.get('body', 'Melhorias gerais e correções de segurança.')
                
                if self.new_version > self.current_version:
                    if 'assets' in data and len(data['assets']) > 0:
                        self.download_url = data['assets'][0]['url'] 
                        # Retorna True, a Versão E AS NOTAS
                        return True, self.new_version, self.release_notes
            return False, None, None
        except Exception: return False, None, None

    def start_update_process(self, parent_window, notes=""):
        if not getattr(sys, 'frozen', False): return
        
        self.release_notes = notes # Garante que temos as notas
        
        self.progress_window = ctk.CTkToplevel(parent_window)
        self.progress_window.title("Atualização Crítica"); self.progress_window.geometry("400x180")
        self.progress_window.attributes("-topmost", True); self.progress_window.overrideredirect(True) 
        x = parent_window.winfo_x() + (parent_window.winfo_width() // 2) - 200
        y = parent_window.winfo_y() + (parent_window.winfo_height() // 2) - 90
        self.progress_window.geometry(f"+{x}+{y}")
        ctk.CTkFrame(self.progress_window, height=10, fg_color=Config.COLORS["primary"]).pack(fill="x")
        ctk.CTkLabel(self.progress_window, text=f"NOVA VERSÃO: {self.new_version}\nBaixando atualização...", font=(Config.FONT_FAMILY, 14, "bold"), text_color="#333").pack(pady=25)
        self.bar = ctk.CTkProgressBar(self.progress_window, width=320, progress_color=Config.COLORS["success"]); self.bar.pack(pady=10); self.bar.set(0)
        threading.Thread(target=self._download_and_swap).start()

    def _download_and_swap(self):
        try:
            # --- SALVA O BILHETE DO CHANGELOG ---
            try:
                with open("changelog.update", "w", encoding="utf-8") as f:
                    f.write(f"VERSÃO {self.new_version}\n\n{self.release_notes}")
            except: pass
            # ------------------------------------

            headers = {"Authorization": f"token {self.token}", "Accept": "application/octet-stream"}
            r = requests.get(self.download_url, headers=headers, stream=True)
            total_length = int(r.headers.get('content-length', 0))
            dl = 0; new_file_name = "update.new"
            with open(new_file_name, 'wb') as f:
                for chunk in r.iter_content(chunk_size=4096):
                    dl += len(chunk); f.write(chunk)
                    if total_length > 0: self.bar.set(dl / total_length)
            bat_script = f"""@echo off\ntimeout /t 2 /nobreak > NUL\n:loop\ndel "{self.exe_name}"\nif exist "{self.exe_name}" goto loop\nrename "{new_file_name}" "{self.exe_name}"\nstart "" "{self.exe_name}"\ndel "%~f0" """
            with open("updater.bat", "w") as bat: bat.write(bat_script)
            subprocess.Popen("updater.bat", shell=True); os._exit(0)
        except Exception: os._exit(1)

# ==============================================================================
#  2. APLICAÇÃO PRINCIPAL (V3.3.0 - INTEGRAL)
# ==============================================================================

logging.basicConfig(filename=Config.LOG_FILE, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', encoding='utf-8')

class AutoDED_App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{Config.APP_TITLE} | {Config.VERSION}")
        
        # Geometria Responsiva
        screen_width = self.winfo_screenwidth(); screen_height = self.winfo_screenheight()
        win_width = int(screen_width * 0.85); win_height = int(screen_height * 0.85)
        x = (screen_width - win_width) // 2; y = (screen_height - win_height) // 2
        self.geometry(f"{win_width}x{win_height}+{x}+{y}")
        self.minsize(1024, 600)
        
        ctk.set_appearance_mode("Light") 
        
        # Garante criação de pastas de ESCRITA
        for f in [Config.DIR_OUTPUT, Config.DIR_TEMP]: 
            os.makedirs(f, exist_ok=True)
        
        # Limpa pastas antigas do usuário (pois agora usamos internas)
        self.limpar_lixo_legado()

        self.dados_pdf = None 
        self.df_mis_atual = None 
        self.modo_atual = None
        self.placeholder_mis_txt = "2. COLE A TABELA DO MIS ABAIXO..." 
        
        self.setup_ui_structure()
        self.mostrar_menu_inicial()
        
        # Inicialização Async (Sync de Logs, Changelog e Updates)
        self.carregar_historico_sidebar() 
        self.after(1000, self.sincronizar_logs_rede)
        self.after(1500, self.verificar_changelog_pendente) # <--- NOVO CHECK
        self.after(3000, self.verificar_atualizacao_silenciosa)

    # --- NOVO: LÓGICA DO CHANGELOG ---
    def verificar_changelog_pendente(self):
        """Verifica se existe um arquivo de novidades deixado pelo atualizador"""
        changelog_file = "changelog.update"
        if os.path.exists(changelog_file):
            try:
                with open(changelog_file, "r", encoding="utf-8") as f:
                    texto = f.read()
                self.mostrar_modal_changelog(texto)
                os.remove(changelog_file) # Limpa para não mostrar de novo
            except: pass

    def mostrar_modal_changelog(self, texto):
        """Exibe o Changelog Bonitão"""
        top = ctk.CTkToplevel(self)
        top.title("Novidades da Atualização")
        top.geometry("500x520")
        top.attributes("-topmost", True)
        
        # Centraliza
        x = self.winfo_x() + (self.winfo_width() // 2) - 250
        y = self.winfo_y() + (self.winfo_height() // 2) - 260
        top.geometry(f"+{x}+{y}")
        
        # Design
        ctk.CTkFrame(top, height=10, fg_color=Config.COLORS["success"]).pack(fill="x")
        ctk.CTkLabel(top, text="SISTEMA ATUALIZADO! 🚀", font=(Config.FONT_FAMILY, 22, "bold"), text_color=Config.COLORS["success"]).pack(pady=(20, 5))
        ctk.CTkLabel(top, text="Confira o que mudou nesta versão:", font=(Config.FONT_FAMILY, 12)).pack(pady=(0, 10))
        
        # Caixa de Texto
        txt = ctk.CTkTextbox(top, width=440, height=320, corner_radius=8, border_color=Config.COLORS["border"], border_width=1, font=("Consolas", 12))
        txt.pack(pady=10, padx=20)
        txt.insert("0.0", texto)
        txt.configure(state="disabled") # Somente leitura
        
        # Botão Fechar
        ctk.CTkButton(top, text="ENTENDI, VAMOS TRABALHAR!", fg_color=Config.COLORS["primary"], text_color="black", hover_color=Config.COLORS["primary_hover"], 
                      font=(Config.FONT_FAMILY, 13, "bold"), height=45, corner_radius=8, command=top.destroy).pack(pady=20)

    # --- ATUALIZADOR ---
    def verificar_atualizacao_silenciosa(self):
        self.updater = SelfUpdater(Config.VERSION, Config.REPO_URL, Config.GITHUB_TOKEN, exe_name=Config.EXE_NAME)
        has_update, new_ver, notes = self.updater.check_for_updates() # Recebe notes também
        if has_update: 
            self.updater.start_update_process(self, notes)

    def toggle_theme(self):
        if ctk.get_appearance_mode() == "Light": ctk.set_appearance_mode("Dark"); self.btn_theme.configure(text="☀") 
        else: ctk.set_appearance_mode("Light"); self.btn_theme.configure(text="🌙") 
        if self.df_mis_atual is not None: self.validar_e_tabular_mis(redraw=True)
        self.carregar_historico_sidebar()

    def limpar_lixo_legado(self):
        """Apaga pastas 'templates' e 'assets' externas que não são mais usadas."""
        if getattr(sys, 'frozen', False): # Só faz isso se for .exe
            pastas_para_apagar = ["templates", "assets"]
            for pasta in pastas_para_apagar:
                caminho_velho = os.path.join(Config.BASE_DIR, pasta)
                if os.path.exists(caminho_velho) and os.path.isdir(caminho_velho):
                    try:
                        shutil.rmtree(caminho_velho)
                    except: pass

    # --- LOGGER ROBUSTO (DUPLA ESCRITA + SYNC) ---
    def get_local_log(self): return os.path.join(Config.BASE_DIR, "historico_local.csv")
    def get_network_log(self): 
        if os.path.exists(Config.DRIVE_PATH): return os.path.join(Config.DRIVE_PATH, "historico_central.csv")
        return None

    def registrar_log_auditoria(self, tipo_doc, ticket, cliente, status, arquivo_path=""):
        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        try: usuario = os.getlogin()
        except: usuario = "DESCONHECIDO"
        row_data = [data_hora, usuario, tipo_doc, ticket, cliente, status, arquivo_path]
        header = ["DATA_HORA", "USUARIO", "TIPO_DOC", "TICKET", "CLIENTE", "STATUS", "CAMINHO_ARQUIVO"]

        # 1. Escrita Local
        try:
            f_local = self.get_local_log()
            novo = not os.path.exists(f_local)
            with open(f_local, "a", encoding="utf-8", newline='') as f:
                w = csv.writer(f, delimiter=';')
                if novo: w.writerow(header)
                w.writerow(row_data)
        except Exception as e: logging.error(f"Erro Log Local: {e}")

        # 2. Escrita Rede
        f_rede = self.get_network_log()
        if f_rede:
            try:
                novo = not os.path.exists(f_rede)
                with open(f_rede, "a", encoding="utf-8", newline='') as f:
                    w = csv.writer(f, delimiter=';')
                    if novo: w.writerow(header)
                    w.writerow(row_data)
            except Exception as e: logging.error(f"Erro Log Rede: {e}")

        self.carregar_historico_sidebar()

    def sincronizar_logs_rede(self):
        f_rede = self.get_network_log()
        f_local = self.get_local_log()
        if not f_rede or not os.path.exists(f_local): return

        try:
            logs_rede = set()
            if os.path.exists(f_rede):
                with open(f_rede, "r", encoding="utf-8") as f:
                    reader = csv.reader(f, delimiter=';')
                    for row in reader:
                        if len(row) > 3: logs_rede.add(f"{row[0]}-{row[3]}") 

            novos_registros = []
            with open(f_local, "r", encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=';')
                next(reader, None)
                for row in reader:
                    if len(row) > 3:
                        chave = f"{row[0]}-{row[3]}"
                        if chave not in logs_rede:
                            novos_registros.append(row)
            
            if novos_registros:
                with open(f_rede, "a", encoding="utf-8", newline='') as f:
                    w = csv.writer(f, delimiter=';')
                    w.writerows(novos_registros)
        except Exception as e: logging.error(f"Erro Sync: {e}")

    def carregar_historico_sidebar(self):
        try:
            for widget in self.history_scroll.winfo_children(): widget.destroy()
            f_log = self.get_local_log()
            if not os.path.exists(f_log):
                ctk.CTkLabel(self.history_scroll, text="Sem histórico.", text_color="gray", font=(Config.FONT_FAMILY, 11)).pack(pady=10)
                return

            historico = []
            try:
                with open(f_log, "r", encoding="utf-8") as f:
                    reader = csv.DictReader(f, delimiter=';')
                    for row in reader: historico.append(row)
            except: return 

            for item in reversed(historico[-15:]):
                self.criar_item_historico(item)
        except Exception: pass

    def criar_item_historico(self, item):
        try:
            tipo = item.get("TIPO_DOC", "DOC")
            cliente = item.get("CLIENTE", "Desconhecido")
            if len(cliente) > 16: cliente = cliente[:14] + ".."
            
            path_local = item.get("CAMINHO_ARQUIVO", "")
            fname = os.path.basename(path_local)
            path_rede = os.path.join(Config.DRIVE_PATH, fname) if Config.DRIVE_PATH else ""
            
            row = ctk.CTkFrame(self.history_scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)
            
            lbl = ctk.CTkLabel(row, text=f"{tipo} | {cliente}", font=(Config.FONT_FAMILY, 11), text_color=Config.COLORS["text"], anchor="w")
            lbl.pack(side="left", padx=(5,0), expand=True, fill="x")
            
            if Config.DRIVE_PATH:
                btn_cloud = ctk.CTkButton(row, text="☁️", width=30, height=25, fg_color=Config.COLORS["input_bg"], 
                                          text_color=Config.COLORS["primary"], hover_color=Config.COLORS["border"],
                                          command=lambda: self.abrir_arquivo(path_rede, "REDE"))
                btn_cloud.pack(side="right", padx=2)
            
            btn_local = ctk.CTkButton(row, text="📂", width=30, height=25, fg_color=Config.COLORS["input_bg"], 
                                      text_color=Config.COLORS["text_light"], hover_color=Config.COLORS["border"],
                                      command=lambda: self.abrir_arquivo(path_local, "LOCAL"))
            btn_local.pack(side="right", padx=2)
        except: pass

    def abrir_arquivo(self, path, tipo):
        if path and os.path.exists(path):
            try: os.startfile(path)
            except: messagebox.showinfo("Info", f"Arquivo em:\n{path}")
        else:
            msg = "Arquivo não encontrado na Rede." if tipo == "REDE" else "Arquivo local movido ou deletado."
            if tipo == "REDE" and not os.path.exists(Config.DRIVE_PATH): msg += "\n(Sem conexão com o Drive G:)"
            messagebox.showwarning("Erro", msg)

    # --- UI STRUCTURE ---
    def setup_ui_structure(self):
        self.grid_columnconfigure(0, weight=0); self.grid_columnconfigure(1, weight=1); self.grid_rowconfigure(0, weight=1)
        
        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=Config.COLORS["bg_card"], border_width=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        ctk.CTkFrame(self.sidebar, width=1, fg_color=Config.COLORS["border"]).place(relx=1, rely=0, relheight=1, anchor="ne")
        ctk.CTkLabel(self.sidebar, text="HISTÓRICO", font=(Config.FONT_FAMILY, 12, "bold"), text_color=Config.COLORS["text_light"]).pack(pady=(35, 15), padx=15, anchor="w")
        self.history_scroll = ctk.CTkScrollableFrame(self.sidebar, fg_color="transparent")
        self.history_scroll.pack(fill="both", expand=True, padx=5, pady=5)

        self.main_area = ctk.CTkFrame(self, fg_color=Config.COLORS["bg_main"], corner_radius=0)
        self.main_area.grid(row=0, column=1, sticky="nsew")
        self.main_area.grid_rowconfigure(0, weight=0) 
        self.main_area.grid_rowconfigure(1, weight=1) 
        self.main_area.grid_columnconfigure(0, weight=1)

        self.header_container = ctk.CTkFrame(self.main_area, height=90, fg_color="transparent", corner_radius=0)
        self.header_container.grid(row=0, column=0, sticky="ew")
        ctk.CTkFrame(self.header_container, height=6, fg_color=Config.COLORS["primary"], corner_radius=0).pack(fill="x")
        
        header_bar = ctk.CTkFrame(self.header_container, height=84, fg_color=Config.COLORS["bg_card"], corner_radius=0)
        header_bar.pack(fill="both", expand=True)
        header_content = ctk.CTkFrame(header_bar, fg_color="transparent"); header_content.pack(side="left", padx=40, pady=10)
        
        logo_carregado = False
        if os.path.exists(Config.LOGO_PATH):
            try:
                pil_img = Image.open(Config.LOGO_PATH)
                ratio = pil_img.width / pil_img.height; h_size = 48; w_size = int(h_size * ratio)
                self.logo_img = ctk.CTkImage(light_image=pil_img, size=(w_size, h_size))
                ctk.CTkLabel(header_content, image=self.logo_img, text="").pack(side="left", padx=(0, 20))
                logo_carregado = True
            except: pass
        if not logo_carregado: ctk.CTkLabel(header_content, text="99Pay", font=(Config.FONT_FAMILY, 30, "bold"), text_color=Config.COLORS["text"]).pack(side="left")

        ctk.CTkLabel(header_content, text="|", font=("Arial", 32), text_color=Config.COLORS["border"]).pack(side="left", padx=15)
        self.lbl_header_title = ctk.CTkLabel(header_content, text=Config.APP_TITLE, font=(Config.FONT_FAMILY, 20, "bold"), text_color=Config.COLORS["text_light"])
        self.lbl_header_title.pack(side="left")

        self.btn_theme = ctk.CTkButton(header_bar, text="🌙", width=45, height=45, fg_color="transparent", hover_color=Config.COLORS["border"], text_color=Config.COLORS["text"], font=("Segoe UI Emoji", 22), command=self.toggle_theme)
        self.btn_theme.pack(side="right", padx=40)

        self.content_area = ctk.CTkFrame(self.main_area, fg_color="transparent")
        self.content_area.grid(row=1, column=0, sticky="nsew", padx=40, pady=30)

    # --- NAVEGAÇÃO ---
    def mostrar_menu_inicial(self):
        for w in self.content_area.winfo_children(): w.destroy()
        self.modo_atual = None
        self.content_area.grid_rowconfigure(0, weight=1); self.content_area.grid_columnconfigure(0, weight=1)
        self.lbl_header_title.configure(text=Config.APP_TITLE); self.dados_pdf = None

        menu = ctk.CTkFrame(self.content_area, fg_color="transparent"); menu.grid(row=0, column=0)
        try: user = os.getlogin().split('.')[0].title()
        except: user = "Operador"
        
        ctk.CTkLabel(menu, text=f"Olá, {user}. Selecione:", font=(Config.FONT_FAMILY, 28, "bold"), text_color=Config.COLORS["text"]).pack(pady=(0, 60))
        
        grid = ctk.CTkFrame(menu, fg_color="transparent"); grid.pack()
        self.criar_btn_menu(grid, "GERAR DED", "Demonstrativo de Dívida", "📊", Config.COLORS["primary"], self.iniciar_ded, 0)
        self.criar_btn_menu(grid, "QUITAÇÃO", "Carta de Pagamento", "✅", Config.COLORS["success"], self.iniciar_quitacao, 1)
        
        ctk.CTkLabel(menu, text=f"{Config.APP_TITLE} {Config.VERSION} • Build Segura", text_color=Config.COLORS["text_light"], font=(Config.FONT_FAMILY, 11)).pack(side="bottom", pady=60)

    def criar_btn_menu(self, parent, t, d, i, c, cmd, col):
        btn = ctk.CTkButton(parent, text=f"{i}\n\n{t}\n{d}", 
                            font=(Config.FONT_FAMILY, 16, "bold"), 
                            width=300, height=180, 
                            corner_radius=16, 
                            fg_color=Config.COLORS["bg_card"], 
                            text_color=Config.COLORS["text"], 
                            hover_color=c, 
                            border_width=1, 
                            border_color=Config.COLORS["border"], 
                            command=cmd)
        btn.grid(row=0, column=col, padx=25, pady=10)

    def iniciar_ded(self): 
        self.limpar(); self.modo_atual = "DED"
        self.lbl_header_title.configure(text="Gerar DED"); self.dashboard("DED")
    
    def iniciar_quitacao(self): 
        self.limpar(); self.modo_atual = "QUIT"
        self.lbl_header_title.configure(text="Carta de Quitação"); self.dashboard("QUIT")
    
    def limpar(self): 
        for w in self.content_area.winfo_children(): w.destroy()

    def dashboard(self, modo):
        self.content_area.grid_columnconfigure(0, weight=1)
        self.content_area.grid_rowconfigure(0, weight=0) 
        self.content_area.grid_rowconfigure(1, weight=0) 
        self.content_area.grid_rowconfigure(2, weight=1) 
        
        top = ctk.CTkFrame(self.content_area, fg_color="transparent"); top.grid(row=0, column=0, sticky="w", pady=(0, 15))
        ctk.CTkButton(top, text="← Voltar ao Menu", fg_color="transparent", text_color=Config.COLORS["text_light"], hover_color=Config.COLORS["border"], width=100, anchor="w", font=(Config.FONT_FAMILY, 13), command=self.mostrar_menu_inicial).pack(side="left")

        cards = ctk.CTkFrame(self.content_area, fg_color="transparent"); cards.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        self.card_nome = self.info_card(cards, "CLIENTE", "---", "Aguardando PDF")
        self.card_contrato = self.info_card(cards, "CCB / CONTRATO", "---", "---")
        self.card_valor = self.info_card(cards, "PRINCIPAL", "---", "---")

        self.work_box = ctk.CTkFrame(self.content_area, fg_color=Config.COLORS["bg_card"], corner_radius=12, border_width=0)
        self.work_box.grid(row=2, column=0, sticky="nsew")
        self.work_box.grid_columnconfigure(0, weight=1)
        self.work_box.grid_rowconfigure(1, weight=1) 

        if modo == "DED": self.form_ded(self.work_box)
        else: self.form_quit(self.work_box)

    def info_card(self, p, t, v, e):
        c = ctk.CTkFrame(p, fg_color=Config.COLORS["bg_card"], corner_radius=10, border_width=0)
        c.pack(side="left", fill="both", expand=True, padx=5)
        ctk.CTkFrame(c, height=4, fg_color=Config.COLORS["primary"], corner_radius=0).pack(fill="x")
        cnt = ctk.CTkFrame(c, fg_color="transparent"); cnt.pack(padx=20, pady=10, fill="both")
        ctk.CTkLabel(cnt, text=t, font=(Config.FONT_FAMILY, 10, "bold"), text_color=Config.COLORS["text_light"]).pack(anchor="w")
        l2 = ctk.CTkLabel(cnt, text=v, font=(Config.FONT_FAMILY, 16, "bold"), text_color=Config.COLORS["text"]); l2.pack(anchor="w", pady=2)
        l3 = ctk.CTkLabel(cnt, text=e, font=(Config.FONT_FAMILY, 10), text_color=Config.COLORS["text_light"]); l3.pack(anchor="w")
        return {'val': l2, 'extra': l3}

    # --- ENGINES ---
    def carregar_pdf(self, origem):
        path = ctk.filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not path: return
        try:
            self.dados_pdf = self.extrair_dados_completo_pdf(path)
            self.atualizar_resumo()
            if origem == 'DED': self.validar_e_tabular_mis() 
            else: self.check_ready_quit()
        except Exception as e: messagebox.showerror("Erro Leitura", f"Falha: {e}")

    def extrair_dados_completo_pdf(self, path):
        reader = PdfReader(path)
        texto = re.sub(r'\s+', ' ', "".join([p.extract_text() for p in reader.pages]))
        if "99Pay SCD_v.1.0_2025" in texto or "99PAY SOCIEDADE DE" in texto.upper(): return scd_99pay.extrair(texto)
        elif "Dívida Originária" in texto: return qi_reneg.extrair(texto)
        else: return qi_standard.extrair(texto)

    def atualizar_resumo(self):
        if not self.dados_pdf:
            self.card_nome['val'].configure(text="---")
            self.card_nome['extra'].configure(text="Aguardando PDF")
            self.card_contrato['val'].configure(text="---")
            self.card_contrato['extra'].configure(text="---")
            self.card_valor['val'].configure(text="---")
            return

        self.card_nome['val'].configure(text=str(self.dados_pdf.get('NOME_CLIENTE', ''))[:22]+"..")
        self.card_nome['extra'].configure(text=f"CPF: {self.dados_pdf.get('CPF_CLIENTE', '')}")
        self.card_contrato['val'].configure(text=self.dados_pdf.get('NUMERO_CCB', '-'))
        self.card_contrato['extra'].configure(text=self.dados_pdf.get('TIPO_CONTRATO', ''), text_color=Config.COLORS["success"])
        self.card_valor['val'].configure(text=f"R$ {self.dados_pdf.get('VALOR_PRINCIPAL', '-')}")

    # --- FORMS ---
    def form_ded(self, p):
        p.grid_rowconfigure(0, weight=0)
        p.grid_rowconfigure(1, weight=1)
        p.grid_rowconfigure(2, weight=0)
        p.grid_columnconfigure(0, weight=1)

        h = ctk.CTkFrame(p, fg_color="transparent"); h.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        ctk.CTkButton(h, text="1. CLIQUE AQUI PARA CARREGAR O PDF (CCB)", command=lambda: self.carregar_pdf('DED'), fg_color=Config.COLORS["primary"], text_color="black", hover_color=Config.COLORS["primary_hover"], height=50, font=(Config.FONT_FAMILY, 12, "bold"), corner_radius=8).pack(fill="x")

        c = ctk.CTkFrame(p, fg_color="transparent"); c.grid(row=1, column=0, sticky="nsew", padx=20, pady=5)
        self.txt_mis = ctk.CTkTextbox(c, border_color=Config.COLORS["border"], border_width=1, font=("Consolas", 11), fg_color=Config.COLORS["input_bg"], text_color=Config.COLORS["text"], corner_radius=6)
        self.txt_mis.pack(fill="both", expand=True) 
        self.txt_mis.insert("0.0", self.placeholder_mis_txt); self.txt_mis.configure(text_color="gray")
        self.txt_mis.bind("<FocusIn>", self.on_mis_focus_in); self.txt_mis.bind("<FocusOut>", self.on_mis_focus_out); self.txt_mis.bind("<KeyRelease>", lambda e: self.validar_e_tabular_mis())
        self.frame_tabela = ctk.CTkScrollableFrame(c, fg_color=Config.COLORS["bg_card"], border_width=1, border_color=Config.COLORS["border"])
        
        ctrl = ctk.CTkFrame(c, fg_color="transparent"); ctrl.pack(fill="x", pady=5)
        self.lbl_status = ctk.CTkLabel(ctrl, text="Aguardando dados...", text_color=Config.COLORS["text_light"], font=(Config.FONT_FAMILY, 11)); self.lbl_status.pack(side="left")
        
        f = ctk.CTkFrame(p, fg_color="transparent"); f.grid(row=2, column=0, sticky="ew", padx=20, pady=(5, 20))
        f.grid_columnconfigure(0, weight=1); f.grid_columnconfigure(1, weight=0)

        self.ent_ticket = ctk.CTkEntry(f, height=45, border_color=Config.COLORS["border"], corner_radius=6, border_width=1, fg_color=Config.COLORS["input_bg"], text_color=Config.COLORS["text"], placeholder_text="3. NÚMERO DO TICKET")
        self.ent_ticket.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.ent_ticket.bind("<KeyRelease>", lambda e: self.check_ded())
        self.btn_ded = ctk.CTkButton(f, text="CONFERIR E GERAR", state="disabled", width=220, height=45, fg_color=Config.COLORS["disabled"], font=(Config.FONT_FAMILY, 13, "bold"), corner_radius=8, command=lambda: self.pre_validacao("DED"))
        self.btn_ded.grid(row=0, column=1)

    def on_mis_focus_in(self, event):
        if self.txt_mis.get("0.0", "end").strip() == self.placeholder_mis_txt:
            self.txt_mis.delete("0.0", "end"); self.txt_mis.configure(text_color=Config.COLORS["text"][0] if ctk.get_appearance_mode()=="Light" else Config.COLORS["text"][1])

    def on_mis_focus_out(self, event):
        if not self.txt_mis.get("0.0", "end").strip():
            self.txt_mis.insert("0.0", self.placeholder_mis_txt); self.txt_mis.configure(text_color="gray")

    def form_quit(self, p):
        container = ctk.CTkFrame(p, fg_color="transparent"); container.pack(fill="both", expand=True, padx=50, pady=30) 
        ctk.CTkButton(container, text="1. CARREGAR PDF (CCB)", command=lambda: self.carregar_pdf('QUIT'), fg_color=Config.COLORS["success"], height=55, font=(Config.FONT_FAMILY, 14, "bold"), corner_radius=8).pack(fill="x", pady=(0, 20))
        ctk.CTkLabel(container, text="DATA QUITAÇÃO (DD/MM/AAAA)", font=(Config.FONT_FAMILY, 11, "bold"), text_color=Config.COLORS["text_light"]).pack(anchor="w", padx=5)
        self.ent_dt_quit = ctk.CTkEntry(container, height=45, border_color=Config.COLORS["border"], border_width=1, fg_color=Config.COLORS["input_bg"], text_color=Config.COLORS["text"], corner_radius=6); self.ent_dt_quit.pack(fill="x", pady=(5, 20))
        self.ent_tk_quit = ctk.CTkEntry(container, height=45, border_color=Config.COLORS["border"], border_width=1, fg_color=Config.COLORS["input_bg"], text_color=Config.COLORS["text"], corner_radius=6, placeholder_text="NÚMERO DO TICKET")
        self.ent_tk_quit.pack(fill="x", pady=(5, 30))
        self.ent_tk_quit.bind("<KeyRelease>", lambda e: self.check_ready_quit())
        self.btn_quit = ctk.CTkButton(container, text="GERAR CARTA DE QUITAÇÃO", state="disabled", height=55, fg_color=Config.COLORS["disabled"], font=(Config.FONT_FAMILY, 14, "bold"), corner_radius=8, command=lambda: self.pre_validacao("QUIT"))
        self.btn_quit.pack(fill="x")

    # --- LÓGICA ---
    def validar_e_tabular_mis(self, redraw=False):
        if not hasattr(self, 'txt_mis') or not self.txt_mis.winfo_exists(): return
        if redraw and self.df_mis_atual is not None: df = self.df_mis_atual
        else:
            txt = self.txt_mis.get("0.0", "end").strip()
            if not txt or txt == self.placeholder_mis_txt: return
            try:
                lines = [s.strip() for s in txt.split('\n') if s.strip()]
                data = []
                for i in range(0, len(lines), 12):
                    grp = [re.sub(r'integralment(?!e)', 'integralmente', g, flags=re.IGNORECASE) for g in lines[i:i+12]]
                    if len(grp) == 12: data.append(grp)
                    elif "Total" in str(grp[0]): data.append(grp + ["-"] * (12 - len(grp)))
                headers = ["Parcelas", "Principal", "Juros", "Imposto", "Juros de Mora", "Vencimento", "Status", "Desconto de Juros", "Valor do Desconto", "Valor Antes do Desconto", "Total Devido", "Total Pago"]
                df = pd.DataFrame(data, columns=headers)
                self.df_mis_atual = df
            except Exception as e: self.lbl_status.configure(text=f"Erro Formato: {e}", text_color=Config.COLORS["danger"]); return

        if not df.empty:
            self.txt_mis.pack_forget(); self.frame_tabela.pack(fill="both", expand=True)
            for w in self.frame_tabela.winfo_children(): w.destroy()
            
            cols = list(df.columns)
            is_dark = ctk.get_appearance_mode() == "Dark"
            head_bg = "#2B2B2B" if is_dark else "#E0E0E0"
            cell_fg = "white" if is_dark else "black"
            
            for i, c in enumerate(cols): 
                ctk.CTkLabel(self.frame_tabela, text=c, font=(Config.FONT_FAMILY, 9, "bold"), width=90, fg_color=head_bg, text_color=cell_fg, corner_radius=4).grid(row=0, column=i, padx=1, pady=1, sticky="ew")
            for idx, row in df.iterrows():
                if idx > 100: break
                is_total = "Total" in str(row['Parcelas'])
                row_bg = ("#444444" if is_total else ("#333333" if idx % 2 == 0 else "#222222")) if is_dark else ("#DDDDDD" if is_total else ("white" if idx % 2 == 0 else "#F5F5F5"))
                for ci, c in enumerate(cols): 
                    ctk.CTkLabel(self.frame_tabela, text=str(row[c]), font=(Config.FONT_FAMILY, 9), width=90, fg_color=row_bg, text_color=cell_fg).grid(row=idx+1, column=ci, padx=1, pady=1, sticky="ew")
            
            if self.dados_pdf:
                is_valid, erros, avisos = BusinessValidators.validar_cruzamento_mis(df, self.dados_pdf)
                if is_valid: 
                    self.lbl_status.configure(text="✓ DADOS CONFEREM", text_color=Config.COLORS["success"])
                    self.btn_ded.configure(state="normal", fg_color=Config.COLORS["success"], text="CONFERIR E GERAR", command=lambda: self.pre_validacao("DED"))
                else: 
                    self.lbl_status.configure(text=f"⚠ {erros[0]}", text_color=Config.COLORS["danger"])
                    self.btn_ded.configure(state="normal", fg_color=Config.COLORS["danger"], text="DADOS INCORRETOS - LIMPAR", command=self.reset_mis)
            else:
                self.lbl_status.configure(text="Tabela OK. Carregue o PDF para validar.", text_color=Config.COLORS["primary"])

    def check_ded(self):
        if not self.dados_pdf: 
            self.btn_ded.configure(state="disabled", text="CARREGUE O PDF", fg_color=Config.COLORS["disabled"])
            return
        if "⚠" in self.lbl_status.cget("text"): return 
        if len(self.ent_ticket.get()) > 0: self.btn_ded.configure(state="normal", fg_color=Config.COLORS["success"], text="CONFERIR E GERAR", command=lambda: self.pre_validacao("DED"))
        else: self.btn_ded.configure(state="disabled", text="INFORME O TICKET")

    def reset_mis(self):
        try:
            self.frame_tabela.pack_forget()
            self.txt_mis.delete("0.0", "end")
            self.txt_mis.pack(fill="both", expand=True)
            self.lbl_status.configure(text="Aguardando dados...", text_color=Config.COLORS["text_light"])
            
            self.txt_mis.insert("0.0", self.placeholder_mis_txt)
            self.txt_mis.configure(text_color="gray")
            
            self.check_ded()
        except Exception: pass

    def check_ready_quit(self):
        if self.dados_pdf and len(self.ent_tk_quit.get()) > 0: self.btn_quit.configure(state="normal", fg_color=Config.COLORS["success"])
        else: self.btn_quit.configure(state="disabled")

    def pre_validacao(self, modo):
        if not self.dados_pdf:
            messagebox.showwarning("Aviso", "Nenhum PDF carregado na memória.")
            if modo == "DED": self.check_ded()
            else: self.check_ready_quit()
            return

        ok, erros = BusinessValidators.validar_regras_financeiras(self.dados_pdf)
        if not ok: 
            if not messagebox.askyesno("Alerta", f"{erros[0]}\nContinuar?"): return
        self.campos_faltantes = BusinessValidators.validar_campos_obrigatorios(self.dados_pdf)
        self.modal_conferencia(modo)

    def modal_conferencia(self, modo):
        m = ctk.CTkToplevel(self); m.title("Auditoria Final"); m.geometry("600x800"); m.grab_set(); 
        m.configure(fg_color=Config.COLORS["bg_main"])
        s = ctk.CTkScrollableFrame(m, fg_color="transparent"); s.pack(fill="both", expand=True, padx=20, pady=20)
        
        self.ents = {}; campos = [
            ("NOME_INSTITUICAO", "Instituição Financeira"), ("CNPJ_INSTITUICAO", "CNPJ Instituição"),
            ("TIPO_CONTRATO", "Tipo Contrato"), ("NUMERO_CCB", "Nº Contrato (CCB)"), 
            ("NOME_CLIENTE", "Nome do Cliente"), ("CPF_CLIENTE", "CPF"), 
            ("ENDERECO_CLIENTE", "Endereço"), ("CIDADE", "Cidade"), ("ESTADO", "Estado (UF)"), 
            ("TEL_CLIENTE", "Telefone"), ("EMAIL_CLIENTE", "E-mail"), 
            ("VALOR_PRINCIPAL", "Valor Principal"), ("VALOR_LIBERADO", "Valor Liberado"), 
            ("VALOR_IOF", "Valor IOF"), ("DATA_EMISSAO_CCB", "Data Emissão"), 
            ("VENCIMENTO_FINAL", "Vencimento Final"), 
            ("TAXA_JUROS_MENSAL", "Tx Juros Mensal"), ("TAXA_JUROS_ANUAL", "Tx Juros Anual"), 
            ("CET_MENSAL", "CET Mensal"), ("CET_ANUAL", "CET Anual"), 
            ("JUROS_MORA", "Juros Mora"), ("MULTA_ATRASO", "Multa Atraso"), 
            ("PRAZO_CONTRATO", "Prazo Dias")
        ]
        
        ctk.CTkLabel(s, text="Confira os dados extraídos:", font=(Config.FONT_FAMILY, 16, "bold"), text_color=Config.COLORS["text"]).pack(pady=(0, 15))
        
        for k, label in campos:
            f = ctk.CTkFrame(s, fg_color=Config.COLORS["bg_card"], corner_radius=6); f.pack(fill="x", pady=4, padx=5)
            ctk.CTkLabel(f, text=label, width=150, anchor="w", font=(Config.FONT_FAMILY, 11, "bold"), text_color=Config.COLORS["text_light"]).pack(side="left", padx=10)
            e = ctk.CTkEntry(f, width=300, border_width=0, fg_color=Config.COLORS["bg_card"], text_color=Config.COLORS["text"])
            valor = str(self.dados_pdf.get(k, ""))
            e.insert(0, valor); e.pack(side="right", fill="x", expand=True, padx=10)
            if k in self.campos_faltantes: e.configure(fg_color=Config.COLORS["danger"], text_color="white"); ctk.CTkLabel(s, text=f"⚠ Preencha o {label}!", text_color=Config.COLORS["danger"]).pack()
            self.ents[k] = e
            
        f = ctk.CTkFrame(s, fg_color=Config.COLORS["bg_card"], corner_radius=6); f.pack(fill="x", pady=10, padx=5)
        ctk.CTkLabel(f, text="Ticket", width=150, anchor="w", font=(Config.FONT_FAMILY, 11, "bold"), text_color=Config.COLORS["text_light"]).pack(side="left", padx=10)
        et = ctk.CTkEntry(f, width=300, border_width=0, fg_color=Config.COLORS["bg_card"], text_color=Config.COLORS["text"])
        et.insert(0, self.ent_ticket.get() if modo == "DED" else self.ent_tk_quit.get()); et.pack(side="right", fill="x", expand=True, padx=10); self.ents['TICKET'] = et
        
        btn = ctk.CTkFrame(m, fg_color="transparent"); btn.pack(fill="x", padx=20, pady=20)
        ctk.CTkButton(btn, text="CANCELAR", fg_color=Config.COLORS["danger"], width=150, height=40, command=m.destroy).pack(side="left")
        ctk.CTkButton(btn, text="CONFIRMAR E GERAR", fg_color=Config.COLORS["success"], width=200, height=40, command=lambda: self.executar_geracao(modo, m)).pack(side="right")

    def executar_geracao(self, modo, modal):
        for k, e in self.ents.items():
            if k != 'TICKET': self.dados_pdf[k] = e.get()
        ticket = self.ents['TICKET'].get()
        modal.destroy()
        if modo == "DED": self.processar_ded(ticket)
        else: self.processar_quitacao(ticket)

    def processar_ded(self, ticket):
        try:
            df = self.df_mis_atual
            last_row_label = str(df.iloc[-1]['Parcelas']).upper()
            if any(x in last_row_label for x in ["TOTAL", "TUTAL", "SOMA"]):
                linha_total = df.iloc[-1]
                df_parcelas = df.iloc[:-1] 
            else:
                linha_total = df.iloc[-1]
                df_parcelas = df

            qtd_total = len(df_parcelas)
            qtd_pagas = len(df_parcelas[df_parcelas['Status'].str.contains('Pago', case=False, na=False)])
            qtd_aberto = qtd_total - qtd_pagas
            
            parcelas_pagas_txt = f"{qtd_pagas} de {qtd_total}" if qtd_pagas > 0 else "Nenhuma"
            if qtd_pagas == qtd_total: parcelas_pagas_txt = "Todas"

            status_geral = str(linha_total['Status'])
            if "Pago" in status_geral and "integralmente" in status_geral:
                saldo_fmt = "0,00"
            else:
                dev = limpar_valor(linha_total['Total Devido'])
                pag = limpar_valor(linha_total['Total Pago'])
                saldo_fmt = f"{(dev - pag):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

            img_path = self.gerar_img_tab(df) 
            
            resumo = {
                'VALOR_TOTAL_DEVIDO': str(linha_total['Total Devido']), 
                'VALOR_TOTAL_PAGO': str(linha_total['Total Pago']), 
                'SALDO_DEVEDOR_ATUAL': saldo_fmt, 
                'STATUS_GERAL': status_geral, 
                'PARCELAS_PAGAS_TXT': parcelas_pagas_txt, 
                'PARCELAS_ABERTO_QTD': str(qtd_aberto),
                'DATA_EMISSAO_DED': datetime.now().strftime("%d/%m/%Y")
            }
            
            path_final = self.gerar_arquivo("template.docx", resumo, ticket, "DED", img_path)
            
            self.registrar_log_auditoria("DED", ticket, self.dados_pdf.get('NOME_CLIENTE'), "SUCESSO", path_final)

        except Exception as e: messagebox.showerror("Erro Geração", str(e))

    def processar_quitacao(self, ticket):
        try: 
            dt = self.ent_dt_quit.get()
            path_final = self.gerar_arquivo("template_quitacao.docx", {'DATA_QUITACAO': dt}, ticket, "QUITACAO")
            self.registrar_log_auditoria("QUITACAO", ticket, self.dados_pdf.get('NOME_CLIENTE'), "SUCESSO", path_final)

        except Exception as e: messagebox.showerror("Erro", str(e))

    def gerar_arquivo(self, tpl_name, extra_ctx, ticket, prefix, img=None):
        tpl = os.path.join(Config.DIR_TEMPLATES, tpl_name)
        if not os.path.exists(tpl): raise Exception(f"Template {tpl_name} não encontrado.")
        doc = DocxTemplate(tpl)
        ctx = {**self.dados_pdf, **extra_ctx}
        if img: ctx['TABELA_FINANCEIRA'] = InlineImage(doc, img, width=Inches(7.1))
        
        safe_ccb = re.sub(r'[^\w\-]', '_', str(self.dados_pdf.get('NUMERO_CCB', 'SEM_CCB')))
        fname = f"{prefix}_{safe_ccb}_Ticket-{ticket}"
        docx_out = os.path.abspath(os.path.join(Config.DIR_OUTPUT, fname + ".docx"))
        pdf_out = os.path.abspath(os.path.join(Config.DIR_OUTPUT, fname + ".pdf"))
        
        doc.render(ctx); doc.save(docx_out)
        
        threading.Thread(target=lambda: self.converter_e_copiar(docx_out, pdf_out)).start()
        
        return pdf_out

    def converter_e_copiar(self, docx, pdf):
        pythoncom.CoInitialize()
        try:
            w = win32.Dispatch('Word.Application'); doc = w.Documents.Open(docx); doc.SaveAs(pdf, FileFormat=17); doc.Close()
            copiou_drive = False
            if Config.DRIVE_PATH and os.path.exists(Config.DRIVE_PATH):
                try: 
                    shutil.copy2(pdf, os.path.join(Config.DRIVE_PATH, os.path.basename(pdf)))
                    copiou_drive = True
                except: pass
            
            msg = "Documento Gerado com Sucesso!\nSalvo na pasta local."
            if copiou_drive: msg += "\n\n✅ Salvo também no Drive Compartilhado!"
            
            messagebox.showinfo("Sucesso", msg)
            os.startfile(Config.DIR_OUTPUT); self.after(0, self.limpar_interface) 
        except Exception as e: messagebox.showerror("Erro Conversão", str(e))
        finally: pythoncom.CoUninitialize()

    def limpar_interface(self):
        self.dados_pdf = None; self.df_mis_atual = None
        self.atualizar_resumo()
        
        if self.modo_atual == "DED":
            if hasattr(self, 'txt_mis'): self.reset_mis()
            if hasattr(self, 'ent_ticket'): self.ent_ticket.delete(0, 'end')
            self.check_ded()
            
        elif self.modo_atual == "QUIT":
            if hasattr(self, 'ent_ticket_quit'): self.ent_ticket_quit.delete(0, 'end')
            self.check_ready_quit()
            
        else:
            self.mostrar_menu_inicial()

    def gerar_img_tab(self, df):
        plt.clf(); plt.close('all')
        max_len = df.astype(str).map(len).max().max()
        font_size = 9 
        if max_len > 12: font_size = 8
        if max_len > 15: font_size = 7

        num_rows = len(df) + 1; total_h = 0.6 + (num_rows * 0.35)
        plt.rcParams['font.sans-serif'] = "Arial"; fig, ax = plt.subplots(figsize=(12, total_h)); ax.axis('off')
        plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
        
        tab = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center')
        tab.auto_set_font_size(False); tab.set_fontsize(font_size)
        
        col_widths = {0: 0.06, 5: 0.09, 6: 0.12, 10: 0.09, 11: 0.09}
        for col_idx, width in col_widths.items(): 
             if col_idx < len(df.columns):
                for r in range(len(df) + 1): tab[r, col_idx].set_width(width)
        
        for (row, col), cell in tab.get_celld().items():
            cell.set_edgecolor('#BDBDBD'); cell.set_linewidth(0.4)
            if row == 0: cell.set_facecolor('#CCCCCC'); cell.set_height(0.6 / total_h)
            elif row == len(df): cell.set_facecolor('#E0E0E0'); cell.set_height(0.35 / total_h)
            else: cell.set_facecolor('#F9F9F9' if row % 2 == 0 else 'white'); cell.set_height(0.35 / total_h)
            
        p = os.path.join(Config.DIR_TEMP, "tab.png"); plt.savefig(p, bbox_inches='tight', pad_inches=0, dpi=300); plt.close()
        return p

if __name__ == "__main__":
    app = AutoDED_App(); app.mainloop()