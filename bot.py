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
                
                for r, linha in enumerate(linhas_csv):
                    if self.evento_fechar.is_set(): break
                    
                    self.log(f"--- Iniciando tarefa {r+1} de {numero_repeticoes} ---")
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

                    v_task = linha.get("TASK") or TASK
                    if adicionar_tasks and v_task:
                        self.log(f"Iniciando inclusão de task: {v_task}")
                        
                        try:
                            botao_listar = meu_iframe.get_by_text("Listar atividades", exact=False).first
                            
                            if await botao_listar.count() == 0:
                                xpath_listar = "/html/body/div[1]/div/div/div/div[3]/div/div/div[7]/div/div/div/div/button/div"
                                botao_listar = meu_iframe.locator(f"xpath={xpath_listar}").first
                            await botao_listar.click(timeout=30000)
                            self.log("Botão 'Listar atividades' clicado. Aguardando 3s para carregar...")
                            
                            await asyncio.sleep(3)
                            
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

                            v_analista = linha.get("ANALISTA") or analista
                             
                            if v_analista:
                                self.log(f"Verificando se analista '{v_analista}' já está selecionado...")
                                
                                xpath_gerenciar = '//*[@id="publishedCanvas"]/div/div[4]/div/div/div[8]/div/div/div/div/button/div/div'
                                botao_gerenciar = meu_iframe.locator(f"xpath={xpath_gerenciar}").first
                                
                                if await botao_gerenciar.count() == 0:
                                    self.log(f"Selecionando analista: {v_analista}")
                                    
                                    campo_analista = meu_iframe.get_by_text("Selecione o Analista", exact=False).first
                                    if await campo_analista.count() == 0:
                                        xpath_campo_analista = "/html/body/div[1]/div/div/div/div[4]/div/div/div[6]/div/div/div/div[1]/div[2]"
                                        campo_analista = meu_iframe.locator(f"xpath={xpath_campo_analista}").first
                                    
                                    if await campo_analista.count() > 0:
                                        await campo_analista.click(timeout=10000)
                                        await page.keyboard.type(v_analista)
                                        await asyncio.sleep(2)
                                        await page.keyboard.press("Enter")
                                        await asyncio.sleep(1)
                                        
                                        try:
                                            xpath_exato = '//*[@id="powerapps-flyout-react-combobox-view-6"]/div/ul/li/div/span'
                                            item_analista = meu_iframe.locator(f"xpath={xpath_exato}").first
                                            if await item_analista.count() == 0:
                                                xpath_item = '//*[contains(@id, "powerapps-flyout-react-combobox-view")]//ul/li//span'
                                                item_analista = meu_iframe.locator(f"xpath={xpath_item}").get_by_text(v_analista, exact=False).first
                                            
                                            await item_analista.click(timeout=10000)
                                            self.log(f"Analista {v_analista} selecionado!")
                                            
                                            try:
                                                await meu_iframe.get_by_text("Nome do Analista", exact=False).first.click(timeout=5000)
                                            except: pass
                                            await asyncio.sleep(2)
                                        except: pass
                                    else:
                                        self.log("Aviso: Campo de analista não encontrado, tentando seguir...")
                                else:
                                    self.log("Analista já parece estar selecionado (botão Gerenciar visível).")

                                self.log("Aguardando 5s antes de clicar em Gerenciar...")
                                await asyncio.sleep(5)
                                
                                botao_gerenciar = meu_iframe.locator(f"xpath={xpath_gerenciar}").first
                                if await botao_gerenciar.count() == 0:
                                    botao_gerenciar = meu_iframe.get_by_text("Gerenciar Atendimento", exact=False).first
                                    
                                await botao_gerenciar.click(timeout=20000)
                                self.log("Botão 'Gerenciar Atendimento' clicado. Aguardando a lista carregar...")
                                
                                # Espera inteligente: aguarda o primeiro item da lista aparecer (max 30s)
                                try:
                                    await meu_iframe.locator(seletor_galeria).first.wait_for(state="visible", timeout=30000)
                                    self.log("Lista detectada! Aguardando 3s para sincronização final...")
                                    await asyncio.sleep(3) # Respiro de segurança para o último item aparecer
                                    self.log("Lista carregada com sucesso!")
                                except:
                                    self.log("Aviso: Tempo esgotado aguardando a lista, tentando prosseguir...")

                                self.log("Acessando a lista de atividades para editar o ÚLTIMO item...")
                                
                                seletor_galeria = 'div[data-control-part="gallery-item"]'
                                
                                tentativas_reload = 0
                                count = 0
                                while tentativas_reload < 6:
                                    itens_galeria = meu_iframe.locator(seletor_galeria)
                                    count = await itens_galeria.count()
                                    # Garante que esperamos aparecer a quantidade certa de tarefas (pelo menos a atual)
                                    if count >= (r + 1):
                                        break
                                    
                                    tentativas_reload += 1
                                    self.log(f"Aguardando a tarefa {r+1} aparecer na lista... (Tentativa {tentativas_reload}/6)")
                                    botao_reload = meu_iframe.locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_Reload"])').first
                                    if await botao_reload.count() > 0:
                                        await botao_reload.click()
                                        await asyncio.sleep(12) # Espera o reload processar
                                    else:
                                        await asyncio.sleep(12)

                                if count > 0:
                                    self.log("Editando o ÚLTIMO item da lista (o recém-criado)...")
                                    item_recente = itens_galeria.last
                                    
                                    botao_editar = item_recente.locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_Edit"])').first
                                    
                                    if await botao_editar.count() > 0:
                                        await botao_editar.click(timeout=15000)
                                        await asyncio.sleep(2)
                                        
                                        campo_input = meu_iframe.locator('input[appmagic-control="DataCardValue17textbox"]').first
                                        if await campo_input.count() == 0:
                                            campo_input = meu_iframe.get_by_title("Chamado", exact=False).first

                                        if await campo_input.count() > 0:
                                            self.log(f"Campo de input encontrado. Limpando e preenchendo com: {v_task}")
                                            await campo_input.click()
                                            await page.keyboard.press("Control+A")
                                            await page.keyboard.press("Backspace")
                                            await campo_input.fill(v_task)
                                            await asyncio.sleep(1) # Pequena pausa para o PowerApps registrar o texto
                                            
                                            botao_salvar_task = meu_iframe.get_by_text("Salvar", exact=False).first
                                            if await botao_salvar_task.count() == 0:
                                                botao_salvar_task = meu_iframe.locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_Save"])').first

                                            self.log("Clicando em Salvar Task...")
                                            await botao_salvar_task.click(timeout=15000)
                                            await asyncio.sleep(10) # Aumentado para 10s para garantia total
                                            self.log("Salvamento enviado! Aguardando a janela fechar...")
                                            
                                            # Espera inteligente: aguarda o botão salvar sumir (indica que a janela fechou)
                                            try:
                                                await botao_salvar_task.wait_for(state="hidden", timeout=30000)
                                                self.log("Janela de edição fechada. Task salva com sucesso!")
                                                await asyncio.sleep(3) # Respiro extra antes de voltar
                                            except:
                                                self.log("Aviso: A janela de edição demorou a fechar, prosseguindo...")
                                                await asyncio.sleep(5)
                                        else:
                                            self.log("Aviso: Campo 'Chamado' não encontrado.")
                                            await page.keyboard.press("Escape")
                                    else:
                                        self.log("Aviso: Ícone de edição não encontrado no primeiro item.")
                                else:
                                    self.log("Aviso: Nenhum item encontrado na lista para editar.")

                                self.log(f"Processamento da task concluído.")
                                
                                self.log("Retornando para a tela inicial via ícone de seta...")
                                try:
                                    botao_voltar_seta = meu_iframe.locator('div.powerapps-icon:has(svg[data-appmagic-icon-name="Basel_BackArrow"])').first
                                    
                                    if await botao_voltar_seta.count() > 0:
                                        await botao_voltar_seta.click(timeout=10000)
                                        self.log("Voltando com sucesso...")
                                        await asyncio.sleep(5)
                                    else:
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
                                    
                    progress_callback((r + 1) / numero_repeticoes, f"Tarefa {r+1}/{numero_repeticoes} concluída!")
                
            except Exception as e:
                self.log(f"Erro de automação no loop da planilha: {e}")

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
