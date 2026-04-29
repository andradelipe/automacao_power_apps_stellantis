import customtkinter as ctk
from tkinter import messagebox
import threading
import asyncio
import os
import json
from datetime import datetime
import webbrowser
from bot import PowerAppsBot

# Configurações de Aparência do CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AplicativoRobo(ctk.CTk):
    # Metadados do Aplicativo
    APP_VERSION = "1.2"
    DEVELOPER_NAME = "Felipe Andrade"
    CONTACT_EMAIL = "adsalfsa@gmail.com"

    def __init__(self):
        super().__init__()

        self.title(f"Robô de Automação - PowerApps v{self.APP_VERSION} | Dev: {self.DEVELOPER_NAME}")
        self.geometry("700x740")
        self.config_file = "settings.json"
        self.evento_fechar = threading.Event()

        # Configuração de Grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Frame Principal (Card)
        self.main_frame = ctk.CTkFrame(self, corner_radius=15)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        # --- Sistema de Abas ---
        self.tabview = ctk.CTkTabview(self.main_frame, corner_radius=10)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        
        self.tab_automacao = self.tabview.add("Automação")
        self.tab_sobre = self.tabview.add("Sobre o Desenvolvedor")

        # --- CONTEÚDO DA ABA AUTOMAÇÃO ---
        
        # --- Linha Analista + Turno ---
        self.frame_analista_turno = ctk.CTkFrame(self.tab_automacao, fg_color="transparent")
        self.frame_analista_turno.pack(fill="x", padx=20, pady=(10, 15))
        self.frame_analista_turno.grid_columnconfigure(0, weight=3)
        self.frame_analista_turno.grid_columnconfigure(1, weight=1)

        self.label_analista = ctk.CTkLabel(self.frame_analista_turno, text="Analista:", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_analista.grid(row=0, column=0, sticky="w")
        self.entry_analista = ctk.CTkEntry(self.frame_analista_turno, placeholder_text="Ex: Steve Jobs", height=35)
        self.entry_analista.grid(row=1, column=0, sticky="nsew", padx=(0, 10))

        self.label_turno = ctk.CTkLabel(self.frame_analista_turno, text="Turno:", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_turno.grid(row=0, column=1, sticky="w")
        self.option_turno = ctk.CTkOptionMenu(self.frame_analista_turno, values=["1° Turno", "2° Turno", "3° Turno"], height=35)
        self.option_turno.grid(row=1, column=1, sticky="nsew")

        # --- Seção Usuário ---
        self.label_usuario = ctk.CTkLabel(self.tab_automacao, text="Usuário:", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_usuario.pack(anchor="w", padx=20)
        self.entry_usuario = ctk.CTkEntry(self.tab_automacao, placeholder_text="Ex: IA2026", height=35)
        self.entry_usuario.pack(fill="x", padx=20, pady=(5, 5))

        self.var_salvar_usuario = ctk.BooleanVar(value=True)
        self.check_salvar_usuario = ctk.CTkCheckBox(self.tab_automacao, text="Salvar usuário", variable=self.var_salvar_usuario, font=ctk.CTkFont(size=12))
        self.check_salvar_usuario.pack(anchor="w", padx=20, pady=(0, 10))

        # --- Seção Planilha ---
        self.label_planilha = ctk.CTkLabel(self.tab_automacao, text="Link da Planilha Google (Opcional):", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_planilha.pack(anchor="w", padx=20)
        self.entry_planilha = ctk.CTkEntry(self.tab_automacao, placeholder_text="Link da planilha...", height=35)
        self.entry_planilha.pack(fill="x", padx=20, pady=(5, 5))
        self.entry_planilha.insert(0, "https://docs.google.com/spreadsheets/d/1_OAJyc98f6eeWDTwp3C2j9XfMdkwV4L7ZYORk59iZuw/edit?usp=sharing")

        self.var_salvar_planilha = ctk.BooleanVar(value=True)
        self.check_salvar_planilha = ctk.CTkCheckBox(self.tab_automacao, text="Salvar link da planilha", variable=self.var_salvar_planilha, font=ctk.CTkFont(size=12))
        self.check_salvar_planilha.pack(anchor="w", padx=20, pady=(0, 10))

        # --- Opções Checkbox ---
        self.frame_checks = ctk.CTkFrame(self.tab_automacao, fg_color="transparent")
        self.frame_checks.pack(fill="x", padx=20, pady=5)
        
        self.var_manter_aberto = ctk.BooleanVar(value=False)
        self.check_manter_aberto = ctk.CTkCheckBox(self.frame_checks, text="Manter navegador aberto", variable=self.var_manter_aberto)
        self.check_manter_aberto.pack(side="left", padx=(0, 20))

        self.var_adicionar_tasks = ctk.BooleanVar(value=False)
        self.check_adicionar_tasks = ctk.CTkCheckBox(self.frame_checks, text="Adicionar Tasks", variable=self.var_adicionar_tasks)
        self.check_adicionar_tasks.pack(side="left")

        # --- Botão Iniciar ---
        self.btn_iniciar = ctk.CTkButton(self.tab_automacao, text="Iniciar Automação", command=self.iniciar_automacao, height=45, font=ctk.CTkFont(size=15, weight="bold"))
        self.btn_iniciar.pack(fill="x", padx=20, pady=20)

        # --- Progresso ---
        self.label_status_progresso = ctk.CTkLabel(self.tab_automacao, text="Aguardando início...", font=ctk.CTkFont(size=12))
        self.label_status_progresso.pack(anchor="w", padx=20)
        self.progress_bar = ctk.CTkProgressBar(self.tab_automacao)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=20, pady=(5, 15))

        # --- Área de Log ---
        self.log_area = ctk.CTkTextbox(self.tab_automacao, height=180, font=("Consolas", 12), text_color="#10B981", fg_color="#0F172A")
        self.log_area.pack(fill="both", expand=True, padx=20, pady=(5, 10))
        self.log_area.configure(state="disabled")

        # --- CONTEÚDO DA ABA SOBRE ---
        self.label_sobre_titulo = ctk.CTkLabel(self.tab_sobre, text="Robô PowerApps", font=ctk.CTkFont(size=28, weight="bold"), text_color=("#3B82F6", "#60A5FA"))
        self.label_sobre_titulo.pack(pady=(40, 5))
        
        self.label_versao = ctk.CTkLabel(self.tab_sobre, text=f"Versão {self.APP_VERSION}", font=ctk.CTkFont(size=12), text_color="gray")
        self.label_versao.pack(pady=(0, 20))

        self.card_info = ctk.CTkFrame(self.tab_sobre, fg_color="transparent", border_width=2, border_color=("#3B82F6", "#60A5FA"), corner_radius=15)
        self.card_info.pack(padx=60, pady=10, fill="x")
        
        self.label_dev_header = ctk.CTkLabel(self.card_info, text="DESENVOLVEDOR", font=ctk.CTkFont(size=10, weight="bold"), text_color="gray")
        self.label_dev_header.pack(pady=(20, 0))
        
        self.label_dev_nome = ctk.CTkLabel(self.card_info, text=self.DEVELOPER_NAME, font=ctk.CTkFont(size=18, weight="bold"))
        self.label_dev_nome.pack(pady=(0, 15))
        
        self.label_contato_header = ctk.CTkLabel(self.card_info, text="SUPORTE E CONTATO", font=ctk.CTkFont(size=10, weight="bold"), text_color="gray")
        self.label_contato_header.pack(pady=(5, 0))
        
        self.label_email = ctk.CTkLabel(self.card_info, text=self.CONTACT_EMAIL, font=ctk.CTkFont(size=14, underline=True), text_color=("#2563EB", "#93C5FD"), cursor="hand2")
        self.label_email.pack(pady=(0, 20))
        self.label_email.bind("<Button-1>", lambda e: self.abrir_email())
        
        self.label_info = ctk.CTkLabel(self.tab_sobre, text="Automação especializada para passagem de turno no PowerApps.\nSistema Seguro | Stellantis Automotive", font=ctk.CTkFont(size=11), text_color="gray")
        self.label_info.pack(side="bottom", pady=30)

        self.carregar_configuracoes()

    def abrir_email(self):
        webbrowser.open(f"mailto:{self.CONTACT_EMAIL}")

    def log(self, mensagem):
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        self.after(0, self._inserir_log, f"{timestamp} {mensagem}")

    def atualizar_progresso(self, valor, texto):
        self.after(0, lambda: self._atualizar_progresso_ui(valor, texto))

    def _atualizar_progresso_ui(self, valor, texto):
        self.progress_bar.set(valor)
        self.label_status_progresso.configure(text=texto)

    def _inserir_log(self, mensagem):
        self.log_area.configure(state="normal")
        self.log_area.insert("end", mensagem + "\n")
        self.log_area.see("end")
        self.log_area.configure(state="disabled")

    def carregar_configuracoes(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    config = json.load(f)
                
                if config.get("salvar_usuario") and "usuario" in config:
                    self.entry_usuario.delete(0, "end")
                    self.entry_usuario.insert(0, config["usuario"])
                
                if "analista" in config:
                    self.entry_analista.delete(0, "end")
                    self.entry_analista.insert(0, config["analista"])

                if "turno" in config:
                    self.option_turno.set(config["turno"])
                
                if config.get("salvar_planilha") and "planilha" in config:
                    self.entry_planilha.delete(0, "end")
                    self.entry_planilha.insert(0, config["planilha"])
                
                self.var_salvar_usuario.set(config.get("salvar_usuario", True))
                self.var_salvar_planilha.set(config.get("salvar_planilha", True))
                self.var_adicionar_tasks.set(config.get("adicionar_tasks", False))
                self.var_manter_aberto.set(config.get("manter_aberto", True))
            except Exception as e:
                print(f"Erro ao carregar configurações: {e}")

    def salvar_configuracoes(self):
        config = {
            "salvar_usuario": self.var_salvar_usuario.get(),
            "salvar_planilha": self.var_salvar_planilha.get(),
            "adicionar_tasks": self.var_adicionar_tasks.get(),
            "manter_aberto": self.var_manter_aberto.get()
        }
        if config["salvar_usuario"]:
            config["usuario"] = self.entry_usuario.get().strip()
        
        config["analista"] = self.entry_analista.get().strip()
        config["turno"] = self.option_turno.get()

        if config["salvar_planilha"]:
            config["planilha"] = self.entry_planilha.get().strip()
        
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.log(f"Erro ao salvar configurações: {e}")

    def iniciar_automacao(self):
        usuario = self.entry_usuario.get().strip()
        analista = self.entry_analista.get().strip()
        turno = self.option_turno.get()
        planilha = self.entry_planilha.get().strip()
        manter_aberto = self.var_manter_aberto.get()
        adicionar_tasks = self.var_adicionar_tasks.get()

        if not usuario:
            messagebox.showwarning("Atenção", "Por favor, preencha o Usuário Microsoft!")
            return

        self.salvar_configuracoes()
        self.toggle_ui_state("disabled")
        self.btn_iniciar.configure(text="Executando...", fg_color="#94A3B8")
        
        self.evento_fechar.clear()
        self.log("=== Iniciando Robô ===")
        self.progress_bar.set(0)
        self.label_status_progresso.configure(text="Iniciando robô...")
        
        bot = PowerAppsBot(self.log, self.evento_fechar, self.habilitar_botao_fechar)
        threading.Thread(target=self.rodar_robo_thread, args=(bot, usuario, analista, turno, planilha, manter_aberto, adicionar_tasks), daemon=True).start()

    def rodar_robo_thread(self, bot, usuario, analista, turno, planilha, manter_aberto, adicionar_tasks):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(bot.run(usuario, analista, turno, planilha, manter_aberto, adicionar_tasks, self.atualizar_progresso))
        except Exception as e:
            self.log(f"Erro fatal: {e}")
        finally:
            self.log("=== Fim da Execução ===")
            loop.close()
            self.after(0, self.resetar_ui)

    def habilitar_botao_fechar(self):
        self.after(0, lambda: self.btn_iniciar.configure(state="normal", text="Encerrar Navegador", fg_color="#DC2626", hover_color="#B91C1C", command=self.acionar_fechar_navegador))

    def acionar_fechar_navegador(self):
        self.btn_iniciar.configure(state="disabled", text="Encerrando...")
        self.evento_fechar.set()

    def toggle_ui_state(self, state):
        self.entry_usuario.configure(state=state)
        self.entry_analista.configure(state=state)
        self.option_turno.configure(state=state)
        self.entry_planilha.configure(state=state)
        self.check_manter_aberto.configure(state=state)
        self.check_salvar_usuario.configure(state=state)
        self.check_salvar_planilha.configure(state=state)
        self.check_adicionar_tasks.configure(state=state)

    def resetar_ui(self):
        self.toggle_ui_state("normal")
        self.btn_iniciar.configure(state="normal", text="Iniciar Automação", fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"], hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"], command=self.iniciar_automacao)
