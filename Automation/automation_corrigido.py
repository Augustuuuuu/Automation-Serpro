import pyautogui as pg
import logging
import time
from datetime import datetime
import os
import sys
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

# ==============================================================================
# CONFIGURAÇÃO
# ==============================================================================
data_hoje = datetime.now().strftime('%d/%m/%Y')
logging.basicConfig(level=logging.INFO, format=f'{data_hoje} - %(message)s')

# ==============================================================================
# FUNÇÕES UTILITÁRIAS
# ==============================================================================
def obter_link_do_usuario():
    logging.info("Solicitando o link do ALM ao usuário...")
    titulo = "Automação ALM"
    texto = "Por favor, insira o link completo do ALM e clique em OK:"
    
    link_inserido = pg.prompt(text=texto, title=titulo)

    if link_inserido:
        logging.info(f"Link recebido: {link_inserido}")
        if not (link_inserido.startswith("http://") or link_inserido.startswith("https://")):
            link_inserido = "https://" + link_inserido
        return link_inserido
    else:
        logging.warning("O usuário cancelou a caixa de diálogo.")
        return None


def iniciar_navegador_com_perfil_usuario(url):
    if not url:
        return None
    try:
        logging.info("Configurando o driver para o Microsoft Edge com perfil de usuário...")
        edge_options = EdgeOptions()
        user_data_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Microsoft", "Edge", "User Data")
        edge_options.add_argument(f"user-data-dir={user_data_dir}")
        service = EdgeService(EdgeChromiumDriverManager().install())
        driver = webdriver.Edge(service=service, options=edge_options)
        driver.get(url)
        driver.maximize_window()
        logging.info("Navegador iniciado e site acessado com sucesso.")
        return driver
    except Exception as e:
        logging.error(f"Erro ao iniciar o navegador: {e}")
        pg.alert(f"Ocorreu um erro ao iniciar o navegador: {e}", "Erro de Automação")
        return None


# ==============================================================================
# CLASSE DA PÁGINA ALM
# ==============================================================================
class automation:
    def __init__(self, driver):
        self.driver = driver
        self.locators = {
            "resumo": (By.CSS_SELECTOR, ".RichTextEditorWidget.cke_editable.cke_contents_ltr"),
            "numero_demanda": (By.CLASS_NAME, "TitleText"),
            "solicitante": (By.XPATH, "//span[@class='ValueLabelHolder']"),
            "data_criacao": (By.CLASS_NAME, "TimeLabel"),
            "codigo_servico": (By.XPATH, "//div[@aria-label='Código de Serviço']"),
            "responsavel": (By.XPATH, "//span[@class='ValueLabelHolder']"),
            "tipo_demanda": (By.XPATH, "//span[@class='ValueLabelHolder']"),
            "aba_atendimento": (By.XPATH, "//a[@title='Atendimento']"),
            "aba_demanda": (By.XPATH, "//span[contains(@class, 'nav-label') and contains(text(), 'Demanda')]"),
            "aba_incluirDemanda": (By.ID, "IncluirDemanda"),
            "nome_demanda": (By.ID, "nome"),
            "descricao_demanda":  (By.ID, "descricao"),
            "nomeResponsavel": (By.ID, "nomeResponsavel"),
            "numero_da_demanda": (By.ID, "numeroDemanda"),
            "salvar": (By.ID, "confirmar"),
            "incluirFuncoes": (By.ID, "incluirFuncoes"),
            "criarContagem": (By.CSS_SELECTOR,"button.swal2-confirm.btn.btn-primary.btn-pills.ml-2"),
            "descricao_contagem": (By.ID, "descricao"),
            "proposito": (By.ID, "proposito"),
            "escopo": (By.ID, "escopo"),
            "titulo": (By.CLASS_NAME, "title-5 align-middle"),
            "url": (By.CSS_SELECTOR, "input[dojoattachpoint='_urlField']"),
            "rotulo" : (By.CSS_SELECTOR, "input[dojoattachpoint='_textField']"),
            "tamanhoPF": (By.XPATH, "//input[@aria-label='Tamanho (PF)']"),
            "aba_visaogeral": (By.XPATH, "//a[@title='Visão Geral']"),
            "salvar_link": (By.CSS_SELECTOR, "button[dojoattachpoint='_okButton']"),
            "comentario": (By.XPATH, "//div[contains(@class, 'RichTextEditorWidget') and contains(@aria-label, 'Coment')]"),

        }

    def obter_textoElemento(self, nome_do_localizador, timeout=20):
        try:
            localizador = self.locators.get(nome_do_localizador)
            if not localizador:
                logging.error(f"❌ Localizador '{nome_do_localizador}' não definido.")
                return None

            logging.info(f"🔎 Aguardando campo '{nome_do_localizador}' (timeout={timeout}s)...")
            wait = WebDriverWait(self.driver, timeout)
            elementos = wait.until(EC.presence_of_all_elements_located(localizador))

            # Lógica por nome
            indices = {
                "solicitante": 1,
                "tipo_demanda": 3,
                "responsavel": 2,
            }
            index = indices.get(nome_do_localizador, 0)

            if len(elementos) > index:
                texto_elemento = elementos[index].text.strip()
            elif elementos:
                texto_elemento = elementos[0].text.strip()
            else:
                logging.error(f"⚠️ Nenhum elemento encontrado para '{nome_do_localizador}'")
                return None

            logging.info(f"✅ Texto de '{nome_do_localizador}': {texto_elemento[:100]}...")
            return texto_elemento

        except TimeoutException:
            logging.error(f"⏰ Timeout: elemento '{nome_do_localizador}' não foi encontrado após {timeout} segundos.")
            return None
        except Exception as e:
            logging.error(f"💥 Erro inesperado ao buscar '{nome_do_localizador}': {e}")
            return None
    def clicar_botao(self, nome_do_botao, timeout=20):
        try:
            localizador = self.locators.get(nome_do_botao)
            if not localizador:
                logging.error(f"Localizador '{nome_do_botao}' não definido.")
                return None
    
            logging.info(f"Aguardando botão '{nome_do_botao}' ficar clicável (timeout={timeout}s)...")
            wait = WebDriverWait(self.driver, timeout)
    
            # Espera o botão estar visível e clicável diretamente
            elemento = wait.until(EC.element_to_be_clickable(localizador))
    
            # Rola até o botão (opcional, mas ajuda)
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
            # sleep removido, já está aguardando via WebDriverWait
    
            try:
                elemento.click()
                logging.info(f"✅ Clique realizado no botão '{nome_do_botao}' com .click().")
            except Exception as e_click:
                logging.warning(f"⚠️ Falha no .click(): {e_click}. Tentando via JavaScript.")
                self.driver.execute_script("arguments[0].click();", elemento)
                logging.info(f"✅ Clique forçado com JavaScript no botão '{nome_do_botao}'.")
    
            return True
    
        except TimeoutException:
            logging.error(f"❌ Botão '{nome_do_botao}' não clicável após {timeout} segundos.")
            return None
        except Exception as e:
            logging.error(f"❌ Erro inesperado ao clicar no botão '{nome_do_botao}': {e}")
            return None
    
    def preencher_campo(self, nome_do_campo, texto, timeout=20):
        try:
            localizador = self.locators.get(nome_do_campo)
            if not localizador:
                logging.error(f"Localizador '{nome_do_campo}' não definido.")
                return None

            logging.info(f"Aguardando campo '{nome_do_campo}' (timeout={timeout}s)...")
            wait = WebDriverWait(self.driver, timeout)
            campo = wait.until(EC.element_to_be_clickable(localizador))
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)

            campo.clear()
            campo.send_keys(texto)
            logging.info(f"✅ Preenchido '{nome_do_campo}' com: {texto}")
            return True

        except TimeoutException:
            logging.error(f"❌ Campo '{nome_do_campo}' não encontrado após {timeout} segundos.")
            return None
        except Exception as e:
            logging.error(f"❌ Erro ao preencher o campo '{nome_do_campo}': {e}")
            return None
    def selecionar_Dropdown(self, placeholder_texto, texto_opcao, timeout=10):
        try:
            wait = WebDriverWait(self.driver, timeout)

            # Clica no dropdown pelo placeholder
            ng_select = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, f"ng-select[placeholder='{placeholder_texto}']"))
            )
            ng_select.click()

            # Localiza o input de pesquisa dentro do dropdown aberto
            input_pesquisa = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, f"ng-select[placeholder='{placeholder_texto}'] input[type='text']"))
            )
            input_pesquisa.clear()
            input_pesquisa.send_keys(texto_opcao)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.ng-option")))  # aguarda opções aparecerem

            # Pressiona ENTER para confirmar a seleção
            input_pesquisa.send_keys(Keys.ENTER)

            logging.info(f"✅ Opção '{texto_opcao}' selecionada no dropdown '{placeholder_texto}' com ENTER.")
            return True

        except Exception as e:
            logging.error(f"❌ Erro ao selecionar '{texto_opcao}' no dropdown '{placeholder_texto}': {e}")
            return False
    def preencher_dataIndice(self, index, data, timeout=10):
        try:
            wait = WebDriverWait(self.driver, timeout)
            campos = wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input[placeholder='__/__/____']"))
            )
            if index >= len(campos):
                logging.error(f"❌ Índice {index} inválido. Só existem {len(campos)} campos.")
                return False    

            campos[index].clear()
            campos[index].send_keys(data)
            campos[index].send_keys(Keys.TAB)
            logging.info(f"📅 Data '{data}' preenchida no campo de índice {index}")
            return True
        except Exception as e:
            logging.error(f"❌ Erro ao preencher campo de data no índice {index}: {e}")
            return False
    def selecionar_dropdown_padrão(self, id_do_select, texto_ou_valor, por_valor=True, timeout=10):
        try:
            wait = WebDriverWait(self.driver, timeout)
            select_element = wait.until(EC.presence_of_element_located((By.ID, id_do_select)))
            select = Select(select_element) 

            if por_valor:
                select.select_by_value(texto_ou_valor)
                logging.info(f"✅ Selecionado valor '{texto_ou_valor}' no select '{id_do_select}'")
            else:
                select.select_by_visible_text(texto_ou_valor)
                logging.info(f"✅ Selecionado texto '{texto_ou_valor}' no select '{id_do_select}'")
            return True
        except Exception as e:
            logging.error(f"❌ Erro ao selecionar no dropdown '{id_do_select}': {e}")
            return False

# Função de execução
def executar_automacao(link_alm):
    navegador = iniciar_navegador_com_perfil_usuario(link_alm)
    if not navegador:
        return

    try:
        automacao = automation(navegador)

        print("✅ [1/12] Aguardando o carregamento da página inicial...")
        WebDriverWait(navegador, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "TitleText"))
        )

        print("✅ [2/12] Extraindo informações do ALM...")
        resumo = automacao.obter_textoElemento("resumo")
        numero_demanda = automacao.obter_textoElemento("numero_demanda")
        numeroDemanda = numero_demanda[15:].strip()
        solicitante = automacao.obter_textoElemento("solicitante")
        data_criacao = automacao.obter_textoElemento("data_criacao")

        meses = {
            'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04', 'mai': '05',
            'jun': '06', 'jul': '07', 'ago': '08', 'set': '09', 'out': '10',
            'nov': '11', 'dez': '12'
        }
        partes = data_criacao.split()
        dia = partes[0]
        mes = meses.get(partes[2], partes[2])
        ano = partes[4]
        data_formatada = f"{dia}/{mes}/{ano}"

        codigo_servico = automacao.obter_textoElemento("codigo_servico")
        tipo_demanda = automacao.obter_textoElemento("tipo_demanda")

        print("✅ [3/12] Clicando na aba de atendimento...")
        automacao.clicar_botao("aba_atendimento")

        responsavel = automacao.obter_textoElemento("responsavel")

        print("✅ [4/12] Acessando o Pontua em nova aba...")
        navegador.execute_script("window.open('https://pontua.estaleiro.serpro.gov.br/pontua-web/#/dashboard','_blank');")
        navegador.switch_to.window(navegador.window_handles[-1])

        print("✅ [5/12] Inserindo nova demanda no Pontua...")
        automacao.clicar_botao("aba_demanda")
        automacao.clicar_botao("aba_incluirDemanda")
        automacao.preencher_campo("nome_demanda", f"{numeroDemanda}: {resumo}")
        automacao.selecionar_Dropdown("selecionar Fronteira/Aplicação", f"{codigo_servico[-5:]}")
        automacao.preencher_campo("descricao_demanda", f"Solicitação: {resumo}")
        automacao.selecionar_Dropdown("selecionar processo", "Ágil")

        if tipo_demanda == "Apuração":
            automacao.selecionar_Dropdown("selecionar tipo de demanda", "Apuração Especial (AESP)")
        elif tipo_demanda == "Melhoria":
            automacao.selecionar_Dropdown("selecionar tipo de demanda", "Manutenção Corretiva")
        else:
            automacao.selecionar_Dropdown("selecionar tipo de demanda", f"{pg.prompt('Digite o tipo de demanda e aperte OK:', 'Automação')}")

        automacao.preencher_dataIndice(0, data_formatada)
        data_atual = datetime.now().strftime("%d/%m/%Y")
        automacao.preencher_dataIndice(1, data_atual)
        automacao.preencher_campo("nomeResponsavel", responsavel)
        automacao.preencher_campo("numero_da_demanda", numeroDemanda)
        automacao.selecionar_Dropdown("selecionar plataforma", "Web")

        if codigo_servico[-5:] == "80728":
            automacao.selecionar_Dropdown("selecionar linguagem", "JAVA")
        else:
            automacao.selecionar_Dropdown("selecionar linguagem", "Low-Code")

        automacao.selecionar_Dropdown("selecionar banco de dados", "MySql")

        print("✅ [6/12] Confirmando antes de criar contagem...")
        resposta = pg.confirm("Deseja continuar com a automação?", "Confirmação", ["OK", "Cancelar"])
        if resposta != 'OK':
            print("🚫 Operação cancelada pelo usuário.")
            sys.exit()
        print("✅ [7/12] Criando contagem...")
        automacao.clicar_botao("salvar")
        automacao.clicar_botao("criarContagem")
        WebDriverWait(navegador, 20).until(
            EC.visibility_of_element_located((By.ID, "descricao"))
        )
        
        automacao.selecionar_dropdown_padrão("tipoContagem", "1: MANUTENCAO", por_valor=True)
        automacao.selecionar_Dropdown("selecionar Roteiro", "SERPRO V3")
        automacao.selecionar_dropdown_padrão("metodoContagem", "5: CONTAGEM_SFP", por_valor=True)
        automacao.preencher_campo("descricao_contagem", f"Contagem da {numero_demanda}")
        automacao.preencher_campo("proposito", "Fornecer o tamanho funcional de uma demanda de manutenção da aplicação.")
        automacao.preencher_campo("escopo", "Fornecer o tamanho funcional de uma demanda de manutenção da aplicação.")

        print("✅ [8/12] Confirmando antes de finalizar a contagem...")
        resposta = pg.confirm("Deseja continuar com a automação?", "Confirmação", ["OK", "Cancelar"])
        if resposta != 'OK':
            print("🚫 Operação cancelada pelo usuário.")
            sys.exit()

        print("✅ [9/12] Salvando contagem...")
        automacao.clicar_botao("salvar")
        automacao.clicar_botao("incluirFuncoes")
        link_pontua = navegador.current_url

        print("✅ [10/12] Retornando ao ALM...")
        navegador.switch_to.window(navegador.window_handles[0])
        automacao.clicar_botao("aba_visaogeral")
        pf = pg.prompt("Quantidade de PF: ", "Pontos de função")
        print("✅ [11/12] Preenchendo informações finais no ALM...")
        mensagem = f"Contagem da {numero_demanda} em método SFP = {pf} PF.\n"
        mensagem += f"Estimativa realizada em {data_hoje} pelo estagiário Augusto Saboia\n"
        automacao.preencher_campo("comentario", mensagem)
        automacao.preencher_campo("comentario", f"{mensagem} {Keys.CONTROL + "l"}")
        automacao.preencher_campo("url", link_pontua)
        automacao.preencher_campo("rotulo", "Link do Pontua")
        automacao.clicar_botao("salvar_link")
        automacao.clicar_botao("aba_atendimento")
        automacao.preencher_campo("tamanhoPF", pf)

        print("✅ [12/12] Finalizando...")
        logging.info("--- EXTRAÇÃO FINALIZADA ---")
        pg.alert("Execução finalizada! O navegador será fechado.", "Encerrando")

    except Exception as e:
        logging.error(f"Erro durante a automação: {e}")
        pg.alert(f"Ocorreu um erro durante a execução: {e}", "Erro")
    finally:
        navegador.close()
        navegador.quit()

# EXECUÇÃO PRINCIPAL
if __name__ == "__main__":
 link = obter_link_do_usuario()
 if link:
     executar_automacao(link)