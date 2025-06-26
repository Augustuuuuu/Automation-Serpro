import pyautogui as pg
import logging
import time
import os
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==============================================================================
# CONFIGURAÇÃO
# ==============================================================================
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ==============================================================================
# FUNÇÕES UTILITÁRIAS
# ==============================================================================
def obter_link_do_usuario():
    logging.info("Solicitando o link do ALM ao usuário...")
    titulo = "Automação ALM"
    texto = "Por favor, insira o link completo do ALM e clique em OK:"
    
    # link_inserido = pg.prompt(text=texto, title=titulo)
    link_inserido = 'https://alm.serpro/ccm/web/projects/Gest%C3%A3o%20de%20Demandas%20Internas#action=com.ibm.team.workitem.viewWorkItem&id=4792741'

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
class AlmPage:
    def __init__(self, driver):
        self.driver = driver
        self.locators = {
            "resumo": (By.CSS_SELECTOR, ".RichTextEditorWidget.cke_editable.cke_contents_ltr"),
            "numero_demanda": (By.CLASS_NAME, "TitleText"),
            "solicitante": (By.XPATH, "//span[@class='ValueLabelHolder']")  # pega o primeiro
        }

    def _obter_texto_do_elemento(self, nome_do_localizador, timeout=20):
        try:
            localizador = self.locators.get(nome_do_localizador)
            if not localizador:
                logging.error(f"Localizador '{nome_do_localizador}' não definido.")
                return None
            
            logging.info(f"Aguardando campo '{nome_do_localizador}' (timeout={timeout}s)...")
            wait = WebDriverWait(self.driver, timeout)
    
            if nome_do_localizador == "solicitante":
                # Exemplo: pegar o segundo ValueLabelHolder, que você disse ser o correto
                # Ajuste aqui se quiser outro índice ou seletor mais específico
                elementos = wait.until(EC.presence_of_all_elements_located(localizador))
                if len(elementos) >= 2:
                    texto_elemento = elementos[1].text.strip()  # índice 1 = segundo elemento
                elif elementos:
                    texto_elemento = elementos[0].text.strip()
                else:
                    logging.error(f"Nenhum elemento encontrado para '{nome_do_localizador}'")
                    return None
            else:
                # Para os demais, pega o único elemento visível
                elemento = wait.until(EC.visibility_of_element_located(localizador))
                texto_elemento = elemento.text.strip()
            
            logging.info(f"Texto de '{nome_do_localizador}': {texto_elemento[:100]}...")
            return texto_elemento
    
        except TimeoutException:
            logging.error(f"Elemento '{nome_do_localizador}' não foi encontrado após {timeout} segundos.")
            return None
        except Exception as e:
            logging.error(f"Erro inesperado ao buscar '{nome_do_localizador}': {e}")
            return None


    def obter_resumo(self):
        return self._obter_texto_do_elemento("resumo")

    def obter_numero_demanda(self):
        return self._obter_texto_do_elemento("numero_demanda")

    def obter_solicitante(self):
        return self._obter_texto_do_elemento("solicitante")


# ==============================================================================
# EXECUÇÃO PRINCIPAL
# ==============================================================================
if __name__ == "__main__":
    logging.info(">>> INICIANDO AUTOMAÇÃO ALM <<<")
    
    link_alm = obter_link_do_usuario()
    
    if link_alm:
        navegador = iniciar_navegador_com_perfil_usuario(link_alm)

        if navegador:
            try:
                pagina_alm = AlmPage(navegador)

                logging.info("Aguardando 8 segundos para a página carregar...")
                time.sleep(8)

                logging.info("--- INICIANDO EXTRAÇÃO DE DADOS ---")
                resumo = pagina_alm.obter_resumo()
                numero_demanda = pagina_alm.obter_numero_demanda()
                solicitante = pagina_alm.obter_solicitante()
                logging.info("--- EXTRAÇÃO FINALIZADA ---")

                mensagem_final = "INFORMAÇÕES EXTRAÍDAS DA DEMANDA:\n\n"
                mensagem_final += f"Número da Demanda: {numero_demanda or 'Não encontrado'}\n"
                mensagem_final += f"Solicitante: {solicitante or 'Não encontrado'}\n"
                mensagem_final += "--------------------------------------\n"
                mensagem_final += f"Resumo: {resumo or 'Não encontrado'}\n"

                pg.alert(mensagem_final, "Relatório da Demanda")
                pg.alert("Execução finalizada! O navegador será fechado.", "Encerrando")
                time.sleep(2)

            except Exception as e:
                logging.error(f"Erro durante a automação: {e}")
            finally:
                logging.info("Fechando o navegador.")
                navegador.quit()

    logging.info(">>> FIM DA EXECUÇÃO <<<")
