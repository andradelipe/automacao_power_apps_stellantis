import asyncio
import customtkinter as ctk
from tkinter import messagebox
import threading
from playwright.async_api import async_playwright
import csv
import urllib.request
import io
import json
import os
from datetime import datetime

# Configurações de Aparência do CustomTkinter
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue")

class PowerAppsBot:
    def __init__(self, log_func, event_fechar, habilitar_fechar_func):
        self.log = log_func
        self.evento_fechar = event_fechar
        self.habilitar_fechar = habilitar_fechar_func

    async def run(self, usuario, planilha_google, manter_aberto):
        HOSTNAME = ""
        MODELO = ""
        OFICINA = ""
        LOCALIZACAO = ""
        SOLICITANTE = ""
        DESCRICAO_RESUMIDA = "PR:TRYOUT"
        DESCREVA_ATENDIMENTO = ""

        self.log("Iniciando contexto do navegador...")
        
        async with async_playwright() as p:
            browser = await p.chromium.launch_persistent_context(
                user_data_dir="./dados_navegador",
                channel="msedge",
                headless=False,
                args=["--disable-blink-features=AutomationControlled", "--start-maximized"],
                no_viewport=True,
                slow_mo=50
            )

            page = browser.pages[0] if browser.pages else await browser.new_page()
            self.log("Acessando a página do PowerApps...")
            await page.goto("https://apps.powerapps.com/play/e/default-d852d5cd-724c-4128-8812-ffa5db3f8507/a/0244c3cf-c487-4af9-963c-fd5728387792?tenantId=d852d5cd-724c-4128-8812-ffa5db3f8507&sourcetime=1737550430441")
            
            meu_iframe = page.frame_locator('iframe#fullscreen-app-host')
            input_campo = meu_iframe.locator('input[appmagic-control="TextInput1textbox"]')
            
            self.log("Aguardando o site carregar e o campo aparecer (Timeout de 90s)...")
            
            try:
                await input_campo.fill(usuario, timeout=90000)
                self.log("Campo de usuário preenchido!")
                
                self.log("Esperando o Microsoft achar o usuário na lista...")
                item_lista = meu_iframe.locator('div[data-control-part="gallery-item"]').first
                
                await item_lista.click(timeout=50000)
                self.log("Item da lista clicado!")
                
                self.log("Procurando o botão de Entrar...")
                botao_entrar = meu_iframe.get_by_text("Entrar", exact=False).first
                await botao_entrar.click(timeout=50000)
                self.log("Botão Entrar clicado!")

                linhas_csv = []
                
                if planilha_google:
                    try:
                        self.log("Tentando baixar tarefas do Google Sheets...")
                        csv_url = planilha_google.replace('/edit?usp=sharing', '/export?format=csv').replace('/edit', '/export?format=csv')
                        
                        req = urllib.request.Request(csv_url, headers={'User-Agent': 'Mozilla/5.0'})
                        resposta = urllib.request.urlopen(req)
                        conteudo = resposta.read().decode('utf-8')
                        
                        leitor = csv.DictReader(io.StringIO(conteudo), delimiter=',')
                        if leitor.fieldnames:
                            linhas_csv = list(leitor)
                        self.log(f"Sucesso! {len(linhas_csv)} tarefas carregadas da nuvem.")
                    except Exception as e:
                        self.log(f"Erro ao acessar Google Sheets. Detalhe: {e}")
                
                if not linhas_csv:
                    self.log("Buscando tarefas do arquivo de excel local (dados.csv)...")
                    try:
                        with open("dados.csv", mode="r", encoding="utf-8-sig") as f:
                            linhas_csv = list(csv.DictReader(f, delimiter=';'))
                        self.log(f"Carregado localmente {len(linhas_csv)} tarefas.")
                    except Exception as e:
                        self.log(f"Aviso: Falha ao carregar dados. Usando teste vazio. Erro: {e}")
                        linhas_csv = [{}]
                    
                numero_repeticoes = len(linhas_csv) if linhas_csv else 1
                
                for r, linha in enumerate(linhas_csv):
                    self.log(f"--- Iniciando tarefa {r+1} de {numero_repeticoes} ---")
                    
                    v_hostname = linha.get("HOSTNAME") or HOSTNAME
                    v_modelo = linha.get("MODELO") or MODELO
                    v_oficina = linha.get("OFICINA") or OFICINA
                    v_localizacao = linha.get("LOCALIZACAO") or LOCALIZACAO
                    v_solicitante = linha.get("SOLICITANTE") or SOLICITANTE
                    v_descricao_resumida = linha.get("DESCRICAO_RESUMIDA") or DESCRICAO_RESUMIDA
                    v_descreva_atendimento = linha.get("DESCREVA_ATENDIMENTO") or DESCREVA_ATENDIMENTO
                    
                    if r > 0:
                        self.log("Aguardando 3 segundos entre tarefas...")
                        await asyncio.sleep(3)

                    self.log("Procurando o botão de 'Inserir Manualmente'...")
                    botao_inserir = meu_iframe.get_by_text("Inserir Manualmente", exact=False).first
                    await botao_inserir.click(timeout=50000)
                    
                    self.log("Preenchendo a primeira parte do formulário...")
                    await meu_iframe.locator('input[appmagic-control="Hostnametextbox"]').fill(v_hostname, timeout=50000)
                    await meu_iframe.locator('input[appmagic-control="Modelotextbox"]').fill(v_modelo, timeout=50000)
                    await meu_iframe.locator('input[appmagic-control="Oficinatextbox"]').fill(v_oficina, timeout=50000)
                    await meu_iframe.locator('input[appmagic-control="Localizaçãotextbox"]').fill(v_localizacao, timeout=50000)
                    
                    self.log("Clicando no botão de 'Confirmar'...")
                    botao_confirmar = meu_iframe.get_by_text("Confirmar", exact=False).first
                    await botao_confirmar.click(timeout=50000)
                    
                    self.log("Preenchendo a segunda parte do formulário...")
                    await meu_iframe.locator('[appmagic-control="TextInput5textbox"]').fill(v_solicitante, timeout=50000)
                    await meu_iframe.locator('[appmagic-control="TextInput5_1textbox"]').fill(v_descricao_resumida, timeout=50000)
                    await meu_iframe.locator('[appmagic-control="TextInput4textarea"]').fill(v_descreva_atendimento, timeout=50000)
                    
                    self.log("Clicando em 'Enviar'...")
                    botao_enviar = meu_iframe.get_by_text("Enviar", exact=False).first
                    await botao_enviar.click(timeout=100000)
                    
                    self.log("Aguardando o sistema registrar o envio...")
                    await asyncio.sleep(5) 
                    
                    self.log(f">>> Tarefa {r+1} enviada com sucesso! <<<")
                
            except Exception as e:
                self.log(f"Erro de automação: {e}")

            if manter_aberto:
                self.log("Navegação concluída. O navegador permanecerá aberto.")
                self.log("Clique no botão vermelho 'Encerrar Navegador' quando terminar.")
                self.habilitar_fechar()
                while not self.evento_fechar.is_set():
                    await asyncio.sleep(1)

            self.log("Encerrando e fechando o navegador...")
            await browser.close()
            self.log("Fim do processamento Playwright.")

class AplicativoRobo(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Robô de Automação - PowerApps")
        self.geometry("700x750")
        self.config_file = "settings.json"
        self.evento_fechar = threading.Event()

        # Configuração de Grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Frame Principal (Card)
        self.main_frame = ctk.CTkFrame(self, corner_radius=15)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Título
        self.label_titulo = ctk.CTkLabel(self.main_frame, text="Configuração do Robô", font=ctk.CTkFont(size=22, weight="bold"))
        self.label_titulo.pack(pady=(20, 20))

        # --- Seção Usuário ---
        self.label_usuario = ctk.CTkLabel(self.main_frame, text="Usuário Microsoft (Ex: sg06951):", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_usuario.pack(anchor="w", padx=30)
        
        self.entry_usuario = ctk.CTkEntry(self.main_frame, placeholder_text="Digite o usuário...", width=400, height=35)
        self.entry_usuario.pack(fill="x", padx=30, pady=(5, 5))

        self.var_salvar_usuario = ctk.BooleanVar(value=True)
        self.check_salvar_usuario = ctk.CTkCheckBox(self.main_frame, text="Salvar usuário para a próxima vez", variable=self.var_salvar_usuario, font=ctk.CTkFont(size=12))
        self.check_salvar_usuario.pack(anchor="w", padx=30, pady=(0, 15))

        # --- Seção Planilha ---
        self.label_planilha = ctk.CTkLabel(self.main_frame, text="Link da Planilha Google (Opcional):", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_planilha.pack(anchor="w", padx=30)
        
        self.entry_planilha = ctk.CTkEntry(self.main_frame, placeholder_text="Link da planilha...", width=400, height=35)
        self.entry_planilha.pack(fill="x", padx=30, pady=(5, 5))
        self.entry_planilha.insert(0, "https://docs.google.com/spreadsheets/d/1_OAJyc98f6eeWDTwp3C2j9XfMdkwV4L7ZYORk59iZuw/edit?usp=sharing")

        self.var_salvar_planilha = ctk.BooleanVar(value=True)
        self.check_salvar_planilha = ctk.CTkCheckBox(self.main_frame, text="Salvar link da planilha", variable=self.var_salvar_planilha, font=ctk.CTkFont(size=12))
        self.check_salvar_planilha.pack(anchor="w", padx=30, pady=(0, 15))

        # --- Opções Adicionais ---
        self.var_manter_aberto = ctk.BooleanVar(value=False)
        self.check_manter_aberto = ctk.CTkCheckBox(self.main_frame, text="Manter navegador aberto ao finalizar", variable=self.var_manter_aberto)
        self.check_manter_aberto.pack(anchor="w", padx=30, pady=(0, 20))

        # --- Botão Iniciar ---
        self.btn_iniciar = ctk.CTkButton(self.main_frame, text="Iniciar Automação", command=self.iniciar_automacao, height=45, font=ctk.CTkFont(size=15, weight="bold"))
        self.btn_iniciar.pack(fill="x", padx=30, pady=(0, 20))

        # --- Área de Log ---
        self.label_log = ctk.CTkLabel(self.main_frame, text="Log de Execução:", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_log.pack(anchor="w", padx=30)
        
        self.log_area = ctk.CTkTextbox(self.main_frame, height=200, font=("Consolas", 12), text_color="#10B981", fg_color="#0F172A")
        self.log_area.pack(fill="both", expand=True, padx=30, pady=(5, 20))
        self.log_area.configure(state="disabled")

        # Carregar configurações salvas
        self.carregar_configuracoes()

    def log(self, mensagem):
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        self.after(0, self._inserir_log, f"{timestamp} {mensagem}")

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
                
                if config.get("salvar_planilha") and "planilha" in config:
                    self.entry_planilha.delete(0, "end")
                    self.entry_planilha.insert(0, config["planilha"])
                
                self.var_salvar_usuario.set(config.get("salvar_usuario", True))
                self.var_salvar_planilha.set(config.get("salvar_planilha", True))
            except Exception as e:
                print(f"Erro ao carregar configurações: {e}")

    def salvar_configuracoes(self):
        config = {
            "salvar_usuario": self.var_salvar_usuario.get(),
            "salvar_planilha": self.var_salvar_planilha.get()
        }
        if config["salvar_usuario"]:
            config["usuario"] = self.entry_usuario.get().strip()
        if config["salvar_planilha"]:
            config["planilha"] = self.entry_planilha.get().strip()
        
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.log(f"Erro ao salvar configurações: {e}")

    def iniciar_automacao(self):
        usuario = self.entry_usuario.get().strip()
        planilha = self.entry_planilha.get().strip()
        manter_aberto = self.var_manter_aberto.get()

        if not usuario:
            messagebox.showwarning("Atenção", "Por favor, preencha o Usuário Microsoft!")
            return

        self.salvar_configuracoes()

        # Desabilitar UI
        self.toggle_ui_state("disabled")
        self.btn_iniciar.configure(text="Executando...", fg_color="#94A3B8")
        
        self.evento_fechar.clear()
        self.log("=== Iniciando Robô ===")
        
        bot = PowerAppsBot(self.log, self.evento_fechar, self.habilitar_botao_fechar)
        threading.Thread(target=self.rodar_robo_thread, args=(bot, usuario, planilha, manter_aberto), daemon=True).start()

    def rodar_robo_thread(self, bot, usuario, planilha, manter_aberto):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(bot.run(usuario, planilha, manter_aberto))
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
        self.entry_planilha.configure(state=state)
        self.check_manter_aberto.configure(state=state)
        self.check_salvar_usuario.configure(state=state)
        self.check_salvar_planilha.configure(state=state)

    def resetar_ui(self):
        self.toggle_ui_state("normal")
        self.btn_iniciar.configure(state="normal", text="Iniciar Automação", fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"], hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"], command=self.iniciar_automacao)

if __name__ == "__main__":
    app = AplicativoRobo()
    app.mainloop()