import asyncio
from playwright.async_api import async_playwright
import csv
import urllib.request
import io
import os
import hashlib

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
                
                # --- FASE 1: CRIAÇÃO DOS ITENS ---
                self.log("INICIANDO FASE 1: Criação dos itens em lote...")
                for r, linha in enumerate(linhas_csv):
                    if self.evento_fechar.is_set(): break
                    
                    self.log(f"--- [FASE 1] Criando item {r+1} de {numero_repeticoes} ---")
                    progress_callback((r * 0.4) / numero_repeticoes, f"Fase 1: Criando {r+1}/{numero_repeticoes}...")
                    
                    v_hostname = linha.get("HOSTNAME") or HOSTNAME
                    v_modelo = linha.get("MODELO") or MODELO
                    v_oficina = linha.get("OFICINA") or OFICINA
                    v_localizacao = linha.get("LOCALIZACAO") or LOCALIZACAO
                    v_solicitante = linha.get("SOLICITANTE") or SOLICITANTE
                    v_descricao_resumida = linha.get("DESCRICAO_RESUMIDA") or DESCRICAO_RESUMIDA
                    v_descreva_atendimento = linha.get("DESCREVA_ATENDIMENTO") or DESCREVA_ATENDIMENTO
                    
                    if r > 0:
                        await asyncio.sleep(1)

                    if self.evento_fechar.is_set(): break
                    self.log("Clicando em 'Inserir Manualmente'...")
                    botao_inserir = meu_iframe.get_by_text("Inserir Manualmente", exact=False).first
                    await botao_inserir.click(timeout=50000)
                    
                    if self.evento_fechar.is_set(): break
                    self.log("Preenchendo formulário (Parte 1)...")
                    await meu_iframe.locator('input[appmagic-control="Hostnametextbox"]').fill(v_hostname, timeout=50000)
                    await meu_iframe.locator('input[appmagic-control="Modelotextbox"]').fill(v_modelo, timeout=50000)
                    await meu_iframe.locator('input[appmagic-control="Oficinatextbox"]').fill(v_oficina, timeout=50000)
                    await meu_iframe.locator('input[appmagic-control="Localizaçãotextbox"]').fill(v_localizacao, timeout=50000)
                    
                    botao_confirmar = meu_iframe.get_by_text("Confirmar", exact=False).first
                    await botao_confirmar.click(timeout=50000)
                    
                    self.log("Preenchendo formulário (Parte 2)...")
                    await meu_iframe.locator('[appmagic-control="TextInput5textbox"]').fill(v_solicitante, timeout=50000)
                    await meu_iframe.locator('[appmagic-control="TextInput5_1textbox"]').fill(v_descricao_resumida, timeout=50000)
                    await meu_iframe.locator('[appmagic-control="TextInput4textarea"]').fill(v_descreva_atendimento, timeout=50000)
                    
                    if self.evento_fechar.is_set(): break
                    self.log("Enviando...")
                    botao_enviar = meu_iframe.get_by_text("Enviar", exact=False).first
                    await botao_enviar.click(timeout=100000)
                    
                    await asyncio.sleep(3) 
                    self.log(f"Item {r+1} criado.")

                # --- FASE 2: VINCULAÇÃO DE TASKS ---
                if adicionar_tasks and not self.evento_fechar.is_set():
                    self.log("INICIANDO FASE 2: Vinculação de Tasks em lote...")
                    
                    try:
                        self.log("Navegando para 'Listar atividades'...")
                        botao_listar = meu_iframe.get_by_text("Listar atividades", exact=False).first
                        if await botao_listar.count() == 0:
                            xpath_listar = "/html/body/div[1]/div/div/div/div[3]/div/div/div[7]/div/div/div/div/button/div"
                            botao_listar = meu_iframe.locator(f"xpath={xpath_listar}").first
                        await botao_listar.click(timeout=30000)
                        await asyncio.sleep(2)
                        
                        # Configuração de Turno (apenas uma vez)
                        if turno != "1° Turno":
                            self.log(f"Configurando turno: {turno}")
                            try:
                                dropdown_turno = meu_iframe.locator('.appmagic-dropdownLabelText').filter(has_text="Turno").first
                                if await dropdown_turno.count() > 0:
                                    texto_atual = await dropdown_turno.inner_text()
                                    if turno.strip() not in texto_atual.strip():
                                        await dropdown_turno.click()
                                        await asyncio.sleep(1)
                                        opcao_menu = meu_iframe.get_by_role("option", name=turno, exact=False).first
                                        if await opcao_menu.count() == 0:
                                            opcao_menu = meu_iframe.get_by_text(turno, exact=False).last
                                        await opcao_menu.click(timeout=10000)
                                        await asyncio.sleep(2)
                            except Exception as e: self.log(f"Aviso Turno: {e}")

                        # Seleção de Analista (apenas uma vez)
                        v_analista_padrao = analista
                        self.log(f"Configurando analista: {v_analista_padrao}")
                        campo_analista = meu_iframe.get_by_text("Selecione o Analista", exact=False).first
                        if await campo_analista.count() == 0:
                            xpath_campo_analista = "/html/body/div[1]/div/div/div/div[4]/div/div/div[6]/div/div/div/div[1]/div[2]"
                            campo_analista = meu_iframe.locator(f"xpath={xpath_campo_analista}").first
                        
                        if await campo_analista.count() > 0:
                            await campo_analista.click(timeout=10000)
                            await page.keyboard.type(v_analista_padrao)
                            await asyncio.sleep(2)
                            await page.keyboard.press("Enter")
                            await asyncio.sleep(2)
                            try:
                                xpath_item = '//*[contains(@id, "powerapps-flyout-react-combobox-view")]//ul/li//span'
                                item_analista = meu_iframe.locator(f"xpath={xpath_item}").get_by_text(v_analista_padrao, exact=False).first
                                await item_analista.click(timeout=10000)
                                await asyncio.sleep(2)
                                # Forçar o fechamento do menu clicando no rótulo
                                try:
                                    await meu_iframe.get_by_text("Nome do Analista", exact=False).first.click(timeout=5000)
                                except: pass
                                await asyncio.sleep(2)
                            except: pass

                        # Entrar na Gestão de Atendimento
                        self.log("Entrando em 'Gerenciar Atendimento'...")
                        xpath_gerenciar = '//*[@id="publishedCanvas"]/div/div[4]/div/div/div[8]/div/div/div/div/button/div/div'
                        botao_gerenciar = meu_iframe.locator(f"xpath={xpath_gerenciar}").first
                        if await botao_gerenciar.count() == 0:
                            botao_gerenciar = meu_iframe.get_by_text("Gerenciar Atendimento", exact=False).first
                        
                        # Usar force=True para garantir que o clique ocorra mesmo com overlays
                        await botao_gerenciar.click(timeout=20000, force=True)
                        await asyncio.sleep(5)

                        # Loop de Vinculação
                        for r, linha in enumerate(linhas_csv):
                            if self.evento_fechar.is_set(): break
                            
                            v_task = linha.get("TASK") or TASK
                            if not v_task:
                                self.log(f"Item {r+1} sem task definida. Pulando...")
                                continue
                            
                            self.log(f"--- [FASE 2] Vinculando task ao item {r+1} ---")
                            progress_callback(0.4 + (r * 0.6) / numero_repeticoes, f"Fase 2: Vinculando {r+1}/{numero_repeticoes}...")
                            
                            seletor_galeria = 'div[data-control-part="gallery-item"]'
                            seletor_janela = 'div[data-control-part="gallery-window"]'
                            seletor_item_especifico = f'div[data-control-part="gallery-item"][aria-posinset="{r+1}"]'
                            
                            item_encontrado = False
                            for tentativa in range(1, 11):
                                if self.evento_fechar.is_set(): break
                                
                                # Scroll proativo
                                if r >= 2:
                                    try:
                                        await meu_iframe.locator(seletor_janela).evaluate("el => { el.scrollLeft = el.scrollWidth; el.scrollTop = el.scrollHeight; }")
                                        await page.keyboard.press("End")
                                    except: pass

                                item_alvo = meu_iframe.locator(seletor_item_especifico).first
                                try:
                                    await item_alvo.wait_for(state="attached", timeout=3000)
                                    await item_alvo.scroll_into_view_if_needed()
                                    await asyncio.sleep(1)
                                    item_encontrado = True
                                    break
                                except:
                                    if tentativa % 2 == 0:
                                        botao_reload = meu_iframe.locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_Reload"])').first
                                        if await botao_reload.count() > 0:
                                            await botao_reload.click()
                                            await asyncio.sleep(5)

                            if item_encontrado:
                                botao_editar = meu_iframe.locator(seletor_item_especifico).locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_Edit"])').first
                                await botao_editar.click(timeout=15000)
                                await asyncio.sleep(3)
                                
                                campo_input = meu_iframe.locator('input[appmagic-control="DataCardValue17textbox"]').first
                                if await campo_input.count() == 0:
                                    campo_input = meu_iframe.get_by_title("Chamado", exact=False).first
                                
                                if await campo_input.count() > 0:
                                    await campo_input.fill(v_task)
                                    await asyncio.sleep(3)
                                    botao_salvar = meu_iframe.get_by_text("Salvar", exact=False).first
                                    if await botao_salvar.count() == 0:
                                        botao_salvar = meu_iframe.locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_Save"])').first
                                    await botao_salvar.click(timeout=15000)
                                    await botao_salvar.wait_for(state="hidden", timeout=15000)
                                    self.log(f"Task {v_task} vinculada ao item {r+1}.")
                                    await asyncio.sleep(3)
                                else:
                                    await page.keyboard.press("Escape")
                                    self.log(f"Aviso: Campo de task não encontrado para item {r+1}.")
                            else:
                                self.log(f"ERRO: Não foi possível encontrar o item {r+1} na lista.")

                        # Retorno final
                        self.log("Processamento em lote concluído. Retornando...")
                        botao_voltar = meu_iframe.locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_BackArrow"])').first
                        if await botao_voltar.count() > 0:
                            await botao_voltar.click()
                            await asyncio.sleep(2)

                    except Exception as e:
                        self.log(f"Erro na Fase 2: {e}")

                progress_callback(1.0, "Automação em lote finalizada!")
                
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
