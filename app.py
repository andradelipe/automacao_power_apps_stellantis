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
import hashlib
from datetime import date
from datetime import datetime
import webbrowser

# Configurações de Aparência do CustomTkinter
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue")

class PowerAppsBot:
    def __init__(self, log_func, event_fechar, habilitar_fechar_func):
        self.log = log_func
        self.evento_fechar = event_fechar
        self.habilitar_fechar = habilitar_fechar_func

    async def run(self, usuario, analista, turno, planilha_google, manter_aberto, adicionar_tasks, progress_callback):
        HOSTNAME = ""
        MODELO = ""
        OFICINA = ""
        LOCALIZACAO = ""
        SOLICITANTE = ""
        DESCRICAO_RESUMIDA = "PR:TRYOUT"
        DESCREVA_ATENDIMENTO = ""
        TASK = ""

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
                    if self.evento_fechar.is_set(): break
                    
                    self.log(f"--- Iniciando tarefa {r+1} de {numero_repeticoes} ---")
                    # Atualiza para o início da tarefa atual
                    progress_callback(r / numero_repeticoes, f"Tarefa {r+1}/{numero_repeticoes}: Iniciando...")
                    
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
                    await asyncio.sleep(3) 
                    
                    self.log(f">>> Tarefa {r+1} enviada com sucesso! <<<")

                    # --- NOVA LÓGICA: ADICIONAR TASKS ---
                    v_task = linha.get("TASK") or TASK
                    if adicionar_tasks and v_task:
                        self.log(f"Iniciando inclusão de task: {v_task}")
                        
                        try:
                            # Prioriza clicar pelo nome para manter o padrão e a robustez
                            botao_listar = meu_iframe.get_by_text("Listar atividades", exact=False).first
                            
                            # Se não achar pelo nome, tenta o XPath como plano B
                            if await botao_listar.count() == 0:
                                xpath_listar = "/html/body/div[1]/div/div/div/div[3]/div/div/div[7]/div/div/div/div/button/div"
                                botao_listar = meu_iframe.locator(f"xpath={xpath_listar}").first
                            await botao_listar.click(timeout=30000)
                            self.log("Botão 'Listar atividades' clicado. Aguardando 10s para carregar...")
                            
                            await asyncio.sleep(3) # Aumentado de 3 para 10 segundos
                            
                            # --- SELEÇÃO DO TURNO (Otimizado: 1° Turno é o padrão) ---
                            if turno != "1° Turno":
                                self.log(f"Verificando se o turno '{turno}' está selecionado...")
                                try:
                                    dropdown_turno = meu_iframe.locator('.appmagic-dropdownLabelText').filter(has_text="Turno").first
                                    
                                    if await dropdown_turno.count() > 0:
                                        texto_atual = await dropdown_turno.inner_text()
                                        if turno.strip() not in texto_atual.strip():
                                            self.log(f"Turno atual é '{texto_atual}'. Alterando para '{turno}'...")
                                            await dropdown_turno.click()
                                            await asyncio.sleep(2)
                                            
                                            opcao_menu = meu_iframe.get_by_role("option", name=turno, exact=False).first
                                            if await opcao_menu.count() == 0:
                                                opcao_menu = meu_iframe.get_by_text(turno, exact=False).last
                                            
                                            await opcao_menu.click(timeout=10000)
                                            self.log(f"Turno '{turno}' selecionado com sucesso!")
                                            await asyncio.sleep(2)
                                        else:
                                            self.log(f"Turno '{turno}' já está selecionado.")
                                    else:
                                        self.log("Aviso: Campo de seleção de turno não encontrado.")
                                except Exception as e_turno:
                                    self.log(f"Aviso ao selecionar turno: {e_turno}")
                            else:
                                self.log("Usando 1° Turno (Padrão do PowerApps).")

                            # --- SELEÇÃO DO ANALISTA (AGORA OPCIONAL SE JÁ ESTIVER NA TELA) ---
                            v_analista = linha.get("ANALISTA") or analista
                             
                            if v_analista:
                                self.log(f"Verificando se analista '{v_analista}' já está selecionado...")
                                
                                # XPath do botão Gerenciar para ver se ele já apareceu
                                xpath_gerenciar = '//*[@id="publishedCanvas"]/div/div[4]/div/div/div[8]/div/div/div/div/button/div/div'
                                botao_gerenciar = meu_iframe.locator(f"xpath={xpath_gerenciar}").first
                                
                                # Se o botão de Gerenciar NÃO estiver visível, precisamos selecionar o analista
                                if await botao_gerenciar.count() == 0:
                                    self.log(f"Selecionando analista: {v_analista}")
                                    
                                    # Tenta o campo pelo texto ou pelo XPath
                                    campo_analista = meu_iframe.get_by_text("Selecione o Analista", exact=False).first
                                    if await campo_analista.count() == 0:
                                        xpath_campo_analista = "/html/body/div[1]/div/div/div/div[4]/div/div/div[6]/div/div/div/div[1]/div[2]"
                                        campo_analista = meu_iframe.locator(f"xpath={xpath_campo_analista}").first
                                    
                                    if await campo_analista.count() > 0:
                                        await campo_analista.click(timeout=10000)
                                        await page.keyboard.type(v_analista) # Removido o delay para ser instantâneo
                                        await asyncio.sleep(2)
                                        await page.keyboard.press("Enter")
                                        await asyncio.sleep(1)
                                        
                                        # Seleção na lista flutuante
                                        try:
                                            xpath_exato = '//*[@id="powerapps-flyout-react-combobox-view-6"]/div/ul/li/div/span'
                                            item_analista = meu_iframe.locator(f"xpath={xpath_exato}").first
                                            if await item_analista.count() == 0:
                                                xpath_item = '//*[contains(@id, "powerapps-flyout-react-combobox-view")]//ul/li//span'
                                                item_analista = meu_iframe.locator(f"xpath={xpath_item}").get_by_text(v_analista, exact=False).first
                                            
                                            await item_analista.click(timeout=10000)
                                            self.log(f"Analista {v_analista} selecionado!")
                                            
                                            # Fecha menu
                                            try:
                                                await meu_iframe.get_by_text("Nome do Analista", exact=False).first.click(timeout=5000)
                                            except: pass
                                            await asyncio.sleep(2)
                                        except: pass
                                    else:
                                        self.log("Aviso: Campo de analista não encontrado, tentando seguir...")
                                else:
                                    self.log("Analista já parece estar selecionado (botão Gerenciar visível).")

                                # Agora clica no Gerenciar (que garantimos que existe ou tentamos selecionar)
                                self.log("Aguardando 5s antes de clicar em Gerenciar...")
                                await asyncio.sleep(5)
                                
                                # Re-localiza para garantir que temos o elemento fresco
                                botao_gerenciar = meu_iframe.locator(f"xpath={xpath_gerenciar}").first
                                if await botao_gerenciar.count() == 0:
                                    botao_gerenciar = meu_iframe.get_by_text("Gerenciar Atendimento", exact=False).first
                                    
                                await botao_gerenciar.click(timeout=20000)
                                self.log("Botão 'Gerenciar Atendimento' clicado. Aguardando 20s...")
                                await asyncio.sleep(20)

                                # --- LÓGICA: EDITAR APENAS O ÚLTIMO ITEM DA LISTA (O MAIS RECENTE) ---
                                self.log("Acessando a lista de atividades para editar o ÚLTIMO item...")
                                
                                seletor_galeria = 'div[data-control-part="gallery-item"]'
                                
                                tentativas_reload = 0
                                while tentativas_reload < 3:
                                    itens_galeria = meu_iframe.locator(seletor_galeria)
                                    count = await itens_galeria.count()
                                    if count > 0:
                                        break
                                    
                                    tentativas_reload += 1
                                    self.log(f"Lista vazia. Recarregando (Tentativa {tentativas_reload}/3)...")
                                    botao_reload = meu_iframe.locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_Reload"])').first
                                    if await botao_reload.count() > 0:
                                        await botao_reload.click()
                                        await asyncio.sleep(10)
                                    else:
                                        await asyncio.sleep(10)

                                if count > 0:
                                    self.log("Editando o ÚLTIMO item da lista (o recém-criado)...")
                                    item_recente = itens_galeria.last # Pega sempre o ÚLTIMO da lista
                                    
                                    # Clica no ícone de edição do último item
                                    botao_editar = item_recente.locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_Edit"])').first
                                    
                                    if await botao_editar.count() > 0:
                                        await botao_editar.click(timeout=15000)
                                        await asyncio.sleep(5)
                                        
                                        campo_input = meu_iframe.locator('input[appmagic-control="DataCardValue17textbox"]').first
                                        if await campo_input.count() == 0:
                                            campo_input = meu_iframe.get_by_title("Chamado", exact=False).first

                                        if await campo_input.count() > 0:
                                            await campo_input.click()
                                            await page.keyboard.press("Control+A")
                                            await page.keyboard.press("Backspace")
                                            await campo_input.fill(v_task)
                                            self.log(f"Task '{v_task}' preenchida no novo chamado.")
                                            
                                            botao_salvar_task = meu_iframe.get_by_text("Salvar", exact=False).first
                                            if await botao_salvar_task.count() == 0:
                                                botao_salvar_task = meu_iframe.locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_Save"])').first

                                            await botao_salvar_task.click(timeout=15000)
                                            self.log(f"Salvo com sucesso. Aguardando 20s...")
                                            await asyncio.sleep(20)
                                        else:
                                            self.log("Aviso: Campo 'Chamado' não encontrado.")
                                            await page.keyboard.press("Escape")
                                    else:
                                        self.log("Aviso: Ícone de edição não encontrado no primeiro item.")
                                else:
                                    self.log("Aviso: Nenhum item encontrado na lista para editar.")

                                self.log(f"Processamento da task concluído.")
                                
                                # --- VOLTAR PARA A TELA INICIAL (USANDO O ÍCONE FORNECIDO) ---
                                self.log("Retornando para a tela inicial via ícone de seta...")
                                try:
                                    # Busca o ícone de 'Basel_BackArrow' que você mandou
                                    botao_voltar_seta = meu_iframe.locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_BackArrow"])').first
                                    
                                    if await botao_voltar_seta.count() > 0:
                                        await botao_voltar_seta.click(timeout=10000)
                                        self.log("Voltando com sucesso...")
                                        await asyncio.sleep(5)
                                    else:
                                        # Plano B se o ícone não for achado (texto ou escape)
                                        self.log("Ícone de seta não encontrado, tentando botão 'Sair' ou 'Voltar'...")
                                        botao_voltar_texto = meu_iframe.get_by_text("Sair", exact=False).first
                                        if await botao_voltar_texto.count() == 0:
                                            botao_voltar_texto = meu_iframe.get_by_text("Voltar", exact=False).first
                                        
                                        if await botao_voltar_texto.count() > 0:
                                            await botao_voltar_texto.click()
                                            await asyncio.sleep(5)
                                        else:
                                            await page.keyboard.press("Escape")
                                            await asyncio.sleep(3)
                                except Exception as e_voltar:
                                    self.log(f"Aviso ao tentar voltar: {e_voltar}")
                                            
                                except Exception as e:
                                    self.log(f"Erro ao processar lista de tasks ou retornar: {e}")
                                    
                        except Exception as e:
                            self.log(f"Erro no fluxo de analista/gerenciamento: {e}")

                    # --- ATUALIZAÇÃO DE PROGRESSO AO FINAL DA TAREFA ---
                    progress_callback((r + 1) / numero_repeticoes, f"Tarefa {r+1}/{numero_repeticoes} concluída!")
                
            except Exception as e:
                self.log(f"Erro de automação no loop da planilha: {e}")

            # --- SINALIZAÇÃO DE CONCLUSÃO TOTAL ---
            progress_callback(1.0, "Automação finalizada com sucesso!")

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
    # Metadados do Aplicativo
    APP_VERSION = "1.1"
    DEVELOPER_NAME = "Felipe Andrade"
    CONTACT_EMAIL = "adsalfsa@gmail.com"

    def __init__(self):
        super().__init__()

        self.title(f"Robô de Automação - PowerApps v{self.APP_VERSION} | Dev: {self.DEVELOPER_NAME}")
        self.geometry("700x740") # Reduzido pois ganhamos espaço vertical
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
        self.entry_analista = ctk.CTkEntry(self.frame_analista_turno, placeholder_text="Ex: Andrade Felipe...", height=35)
        self.entry_analista.grid(row=1, column=0, sticky="nsew", padx=(0, 10))

        self.label_turno = ctk.CTkLabel(self.frame_analista_turno, text="Turno:", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_turno.grid(row=0, column=1, sticky="w")
        self.option_turno = ctk.CTkOptionMenu(self.frame_analista_turno, values=["1° Turno", "2° Turno", "3° Turno"], height=35)
        self.option_turno.grid(row=1, column=1, sticky="nsew")

        # --- Seção Usuário ---
        self.label_usuario = ctk.CTkLabel(self.tab_automacao, text="Usuário Microsoft:", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_usuario.pack(anchor="w", padx=20)
        self.entry_usuario = ctk.CTkEntry(self.tab_automacao, placeholder_text="Digite o usuário...", height=35)
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

        # --- CONTEÚDO DA ABA SOBRE (Design Profissional) ---
        
        # Banner de Título e Versão
        self.label_sobre_titulo = ctk.CTkLabel(self.tab_sobre, text="Robô PowerApps", font=ctk.CTkFont(size=28, weight="bold"), text_color=("#3B82F6", "#60A5FA"))
        self.label_sobre_titulo.pack(pady=(40, 5))
        
        self.label_versao = ctk.CTkLabel(self.tab_sobre, text=f"Versão {self.APP_VERSION}", font=ctk.CTkFont(size=12), text_color="gray")
        self.label_versao.pack(pady=(0, 20))

        # Card Centralizado de Informações (Visual Outlined Moderno)
        self.card_info = ctk.CTkFrame(self.tab_sobre, fg_color="transparent", border_width=2, border_color=("#3B82F6", "#60A5FA"), corner_radius=15)
        self.card_info.pack(padx=60, pady=10, fill="x")
        
        # Seção: Desenvolvedor
        self.label_dev_header = ctk.CTkLabel(self.card_info, text="DESENVOLVEDOR", font=ctk.CTkFont(size=10, weight="bold"), text_color="gray")
        self.label_dev_header.pack(pady=(20, 0))
        
        self.label_dev_nome = ctk.CTkLabel(self.card_info, text=self.DEVELOPER_NAME, font=ctk.CTkFont(size=18, weight="bold"))
        self.label_dev_nome.pack(pady=(0, 15))
        
        # Seção: Contato/Suporte
        self.label_contato_header = ctk.CTkLabel(self.card_info, text="SUPORTE E CONTATO", font=ctk.CTkFont(size=10, weight="bold"), text_color="gray")
        self.label_contato_header.pack(pady=(5, 0))
        
        self.label_email = ctk.CTkLabel(self.card_info, text=self.CONTACT_EMAIL, font=ctk.CTkFont(size=14, underline=True), text_color=("#2563EB", "#93C5FD"), cursor="hand2")
        self.label_email.pack(pady=(0, 20))
        self.label_email.bind("<Button-1>", lambda e: self.abrir_email())
        
        # Rodapé de Informações do Sistema
        self.label_info = ctk.CTkLabel(self.tab_sobre, text="Automação especializada para passagem de turno no PowerApps.\nSistema Seguro | Stellantis Automotive", font=ctk.CTkFont(size=11), text_color="gray")
        self.label_info.pack(side="bottom", pady=30)

        # Carregar configurações salvas
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

        # Desabilitar UI
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

if __name__ == "__main__":
    app = AplicativoRobo()
    app.mainloop()