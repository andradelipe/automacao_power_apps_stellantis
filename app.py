import asyncio
import tkinter as tk
from tkinter import scrolledtext, messagebox
import threading
from playwright.async_api import async_playwright
import csv
import urllib.request
import io

class AplicativoRobo:
    def __init__(self, root):
        self.root = root
        self.root.title("Robô de Automação - PowerApps")
        self.root.geometry("650x600")
        
        # Cores mais modernas (Clean UI)
        bg_color = "#F0F4F8"
        card_color = "#FFFFFF"
        text_color = "#1E293B"
        accent_color = "#2563EB"
        accent_hover = "#1D4ED8"
        
        self.root.configure(bg=bg_color)
        
        # Estilos de Fonte mais grossos ("bold") e limpos
        font_label = ("Segoe UI", 11, "bold")
        font_input = ("Segoe UI", 11)
        font_btn = ("Segoe UI", 12, "bold")
        
        # Frame principal simulando um "card" moderno com padding maior no fundo
        frame = tk.Frame(root, padx=30, pady=25, bg=card_color, relief="flat", bd=0)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Título da tela
        tk.Label(frame, text="Configuração do Robô", font=("Segoe UI", 16, "bold"), bg=card_color, fg=text_color).pack(anchor=tk.W, pady=(0, 20))

        # Usuário Microsoft
        tk.Label(frame, text="Usuário Microsoft (Ex: sg06951):", font=font_label, bg=card_color, fg=text_color).pack(anchor=tk.W)
        self.entry_usuario = tk.Entry(frame, width=50, font=font_input, bg="#F8FAFC", fg=text_color, relief="solid", bd=1)
        self.entry_usuario.pack(fill=tk.X, pady=(5, 20), ipady=5)

        # Planilha Google
        tk.Label(frame, text="Link da Planilha Google (Deixe vazio p/ usar dados.csv):", font=font_label, bg=card_color, fg=text_color).pack(anchor=tk.W)
        self.entry_planilha = tk.Entry(frame, width=50, font=font_input, bg="#F8FAFC", fg=text_color, relief="solid", bd=1)
        # Preenche o link padrão
        self.entry_planilha.insert(0, "https://docs.google.com/spreadsheets/d/1_OAJyc98f6eeWDTwp3C2j9XfMdkwV4L7ZYORk59iZuw/edit?usp=sharing")
        self.entry_planilha.pack(fill=tk.X, pady=(5, 10), ipady=5)

        # Checkbox para manter o navegador aberto
        self.var_manter_aberto = tk.BooleanVar(value=False)
        self.check_manter_aberto = tk.Checkbutton(frame, text="Manter navegador aberto ao finalizar", 
                                                  variable=self.var_manter_aberto, bg=card_color, fg=text_color, font=font_input,
                                                  activebackground=card_color, selectcolor=card_color)
        self.check_manter_aberto.pack(anchor=tk.W, pady=(0, 15))

        # Botão Iniciar Modernizado
        self.btn_iniciar = tk.Button(frame, text="Iniciar Automação", command=self.iniciar_automacao, 
                                     bg=accent_color, fg="white", font=font_btn, relief="flat", 
                                     activebackground=accent_hover, activeforeground="white", cursor="hand2")
        self.btn_iniciar.pack(fill=tk.X, pady=(0, 20), ipady=8)

        self.evento_fechar = threading.Event()

        # Log de execução com efeito "hacker/terminal" mas num visual mais agradável
        tk.Label(frame, text="Log de Execução:", font=font_label, bg=card_color, fg=text_color).pack(anchor=tk.W)
        self.log_area = scrolledtext.ScrolledText(frame, width=70, height=12, state='disabled', 
                                                  bg="#0F172A", fg="#10B981", font=("Consolas", 10), relief="flat")
        self.log_area.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

    def log(self, mensagem):
        # Como o log será chamado de outra thread, usamos root.after para atualizar a interface gráfica com segurança
        self.root.after(0, self._inserir_log, mensagem)

    def _inserir_log(self, mensagem):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, mensagem + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def iniciar_automacao(self):
        usuario = self.entry_usuario.get().strip()
        planilha = self.entry_planilha.get().strip()
        manter_aberto = self.var_manter_aberto.get()

        if not usuario:
            messagebox.showwarning("Atenção", "Por favor, preencha o Usuário Microsoft antes de iniciar!")
            return

        # Desabilita botões e campos durante a execução
        self.btn_iniciar.config(state='disabled', bg="#94A3B8", text="Executando Automação...", cursor="wait")
        self.entry_usuario.config(state='disabled')
        self.entry_planilha.config(state='disabled')
        self.check_manter_aberto.config(state='disabled')
        
        self.evento_fechar.clear()

        self.log("=== Iniciando Robô ===")
        
        # Inicia o robô em uma thread separada para não travar a interface
        threading.Thread(target=self.rodar_robo_thread, args=(usuario, planilha, manter_aberto), daemon=True).start()

    def rodar_robo_thread(self, usuario, planilha, manter_aberto):
        # Cria um novo event loop para essa thread
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        wrapper_habilitar_fechar = lambda: self.root.after(0, self.habilitar_botao_fechar)
        
        try:
            loop.run_until_complete(robo_main(usuario, planilha, manter_aberto, self.evento_fechar, self.log, wrapper_habilitar_fechar))
        except Exception as e:
            self.log(f"Erro fatal: {e}")
        finally:
            self.log("=== Fim da Execução ===")
            loop.close()
            # Reabilita os campos após terminar
            self.root.after(0, self.reabilitar_botoes)

    def habilitar_botao_fechar(self):
        self.btn_iniciar.config(state='normal', bg="#DC2626", text="Encerrar Navegador", cursor="hand2", command=self.acionar_fechar_navegador)

    def acionar_fechar_navegador(self):
        self.btn_iniciar.config(state='disabled', bg="#94A3B8", text="Encerrando...", cursor="wait")
        self.evento_fechar.set()

    def reabilitar_botoes(self):
        # Quando termina (mesmo por erro ou sucesso), ele reseta o botão para você poder clicar de novo.
        self.btn_iniciar.config(state='normal', bg="#2563EB", text="Iniciar Automação", cursor="hand2", command=self.iniciar_automacao)
        self.entry_usuario.config(state='normal')
        self.entry_planilha.config(state='normal')
        self.check_manter_aberto.config(state='normal')


async def robo_main(usuario, planilha_google, manter_aberto, evento_fechar, func_log, func_habilitar_fechar):
    HOSTNAME = ""
    MODELO = ""
    OFICINA = ""
    LOCALIZACAO = ""
    SOLICITANTE = ""
    DESCRICAO_RESUMIDA ="PR:TRYOUT"
    DESCREVA_ATENDIMENTO=""

    func_log("Iniciando contexto do navegador...")
    
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
        func_log("Acessando a página do PowerApps...")
        await page.goto("https://apps.powerapps.com/play/e/default-d852d5cd-724c-4128-8812-ffa5db3f8507/a/0244c3cf-c487-4af9-963c-fd5728387792?tenantId=d852d5cd-724c-4128-8812-ffa5db3f8507&sourcetime=1737550430441")
        
        meu_iframe = page.frame_locator('iframe#fullscreen-app-host')
        input_campo = meu_iframe.locator('input[appmagic-control="TextInput1textbox"]')
        
        func_log("Aguardando o site carregar e o campo aparecer (Timeout de 90s)...")
        
        try:
            await input_campo.fill(usuario, timeout=90000)
            func_log("Campo de usuário preenchido!")
            
            func_log("Esperando o Microsoft achar o usuário na lista...")
            item_lista = meu_iframe.locator('div[data-control-part="gallery-item"]').first
            
            await item_lista.click(timeout=50000)
            func_log("Item da lista clicado!")
            
            func_log("Procurando o botão de Entrar...")
            botao_entrar = meu_iframe.get_by_text("Entrar", exact=False).first
            await botao_entrar.click(timeout=50000)
            func_log("Botão Entrar clicado!")

            linhas_csv = []
            
            if planilha_google:
                try:
                    func_log("\nTentando baixar tarefas do Google Sheets...")
                    csv_url = planilha_google.replace('/edit?usp=sharing', '/export?format=csv').replace('/edit', '/export?format=csv')
                    
                    req = urllib.request.Request(csv_url, headers={'User-Agent': 'Mozilla/5.0'})
                    resposta = urllib.request.urlopen(req)
                    conteudo = resposta.read().decode('utf-8')
                    
                    leitor = csv.DictReader(io.StringIO(conteudo), delimiter=',')
                    if leitor.fieldnames:
                        linhas_csv = list(leitor)
                    func_log(f"Sucesso! {len(linhas_csv)} tarefas carregadas da nuvem.")
                except Exception as e:
                    func_log(f"Erro ao acessar Google Sheets. O envio falhou ou o link não é público. Detalhe: {e}")
            
            if not linhas_csv:
                func_log("\nBuscando tarefas do arquivo de excel local (dados.csv)...")
                try:
                    with open("dados.csv", mode="r", encoding="utf-8-sig") as f:
                        linhas_csv = list(csv.DictReader(f, delimiter=';'))
                    func_log(f"Carregado localmente {len(linhas_csv)} tarefas.")
                except Exception as e:
                    func_log(f"Aviso: Nem nuvem nem arquivo 'dados.csv' funcionaram. Usando variáveis vazias de teste. Erro: {e}")
                    linhas_csv = [{}]
                
            numero_repeticoes = len(linhas_csv) if linhas_csv else 1
            
            for r, linha in enumerate(linhas_csv):
                func_log(f"\n--- Iniciando tarefa {r+1} de {numero_repeticoes} ---")
                
                v_hostname = linha.get("HOSTNAME") or HOSTNAME
                v_modelo = linha.get("MODELO") or MODELO
                v_oficina = linha.get("OFICINA") or OFICINA
                v_localizacao = linha.get("LOCALIZACAO") or LOCALIZACAO
                v_solicitante = linha.get("SOLICITANTE") or SOLICITANTE
                v_descricao_resumida = linha.get("DESCRICAO_RESUMIDA") or DESCRICAO_RESUMIDA
                v_descreva_atendimento = linha.get("DESCREVA_ATENDIMENTO") or DESCREVA_ATENDIMENTO
                
                if r > 0:
                    func_log("Aguardando 3 segundos entre tarefas...")
                    await asyncio.sleep(3)

                func_log("Procurando o botão de 'Inserir Manualmente'...")
                botao_inserir = meu_iframe.get_by_text("Inserir Manualmente", exact=False).first
                await botao_inserir.click(timeout=50000)
                
                func_log("Preenchendo a primeira parte do formulário...")
                await meu_iframe.locator('input[appmagic-control="Hostnametextbox"]').fill(v_hostname, timeout=50000)
                await meu_iframe.locator('input[appmagic-control="Modelotextbox"]').fill(v_modelo, timeout=50000)
                await meu_iframe.locator('input[appmagic-control="Oficinatextbox"]').fill(v_oficina, timeout=50000)
                await meu_iframe.locator('input[appmagic-control="Localizaçãotextbox"]').fill(v_localizacao, timeout=50000)
                
                func_log("Clicando no botão de 'Confirmar'...")
                botao_confirmar = meu_iframe.get_by_text("Confirmar", exact=False).first
                await botao_confirmar.click(timeout=50000)
                
                func_log("Preenchendo a segunda parte do formulário...")
                await meu_iframe.locator('[appmagic-control="TextInput5textbox"]').fill(v_solicitante, timeout=50000)
                await meu_iframe.locator('[appmagic-control="TextInput5_1textbox"]').fill(v_descricao_resumida, timeout=50000)
                await meu_iframe.locator('[appmagic-control="TextInput4textarea"]').fill(v_descreva_atendimento, timeout=50000)
                
                func_log("Clicando em 'Enviar'...")
                botao_enviar = meu_iframe.get_by_text("Enviar", exact=False).first
                await botao_enviar.click(timeout=100000)
                
                func_log("Aguardando o sistema registrar o envio...")
                await asyncio.sleep(5) # Aguarda 5 segundos para o PowerApps processar o clique
                
                func_log(">>> Tarefa enviada com sucesso! <<<")
            
        except Exception as e:
            func_log(f"Erro de automação em alguma etapa: {e}")

        if manter_aberto:
            func_log("\nNavegação concluída. O navegador permanecerá aberto.")
            func_log("Você pode revisar o resultado. Quando terminar, clique no botão vermelho 'Encerrar Navegador'.")
            func_habilitar_fechar()
            # Fica aguardando o usuário clicar no botão para disparar o evento
            while not evento_fechar.is_set():
                await asyncio.sleep(1)

        # Fechamento automático como o usuário pediu
        func_log("\nEncerrando e fechando o navegador de forma segura...")
        await browser.close()
        func_log("Fim do processamento Playwright.")

if __name__ == "__main__":
    root = tk.Tk()
    app = AplicativoRobo(root)
    root.mainloop()