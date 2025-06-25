import pyautogui as pg
import logging
import time
import os
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.common.exceptions import WebDriverException, NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==============================================================================
# 1. SETUP DE LOGGING
# ==============================================================================
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def obter_link_do_usuario():
    """Usa PyAutoGUI para solicitar o link ao usuário."""
    logging.info("Solicitando o link do ALM ao usuário...")
    titulo = "Automação ALM"
    texto = "Por favor, insira o link completo do ALM e clique em OK:"
    
    link_inserido = pg.prompt(text=texto, title=titulo)

    if link_inserido:
        logging.info(f"Link recebido: {link_inserido}")
        if not (link_inserido.startswith("http://") or link_inserido.startswith("https://")):
            link_inserido = "http://" + link_inserido
        return link_inserido
    else:
        logging.warning("O usuário cancelou a caixa de diálogo.")
        return None


def iniciar_navegador_com_perfil_usuario(url):
    """Inicia o Edge com o perfil de usuário padrão."""
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
        logging.error(f"Ocorreu um erro ao iniciar o navegador: {e}")
        pg.alert(f"Ocorreu um erro ao iniciar o navegador: {e}", "Erro de Automação")
        return None

# ==============================================================================
# 2. CLASSE DE INTERAÇÃO COM A PÁGINA
# Esta classe vai conter todos os métodos para interagir com a página do ALM.
# ==============================================================================
class AlmPage:
    """
    Representa a página do ALM, encapsulando os localizadores de elementos
    e as ações que podem ser realizadas na página.
    """
    def __init__(self, driver):
        """
        O construtor da classe recebe a instância do driver do Selenium.
        """
        self.driver = driver
        # Define um localizador para o campo de resumo.
        # Usamos CSS_SELECTOR porque o elemento tem múltiplas classes.
        # Pegamos as classes mais descritivas para criar um seletor único.
        self.resumo_locator = (By.CSS_SELECTOR, ".RichTextEditorWidget.cke_editable.cke_contents_ltr")

    def obter_resumo(self, timeout=10):
        """
        Encontra o campo de resumo na página, espera ele ficar visível
        e retorna o seu texto.

        Args:
            timeout (int): O tempo máximo em segundos para esperar o elemento aparecer.

        Returns:
            str: O texto do campo de resumo, ou None se não for encontrado.
        """
        try:
            logging.info(f"Aguardando o campo de resumo ficar visível (timeout={timeout}s)...")
            
            # Espera Inteligente: o script vai esperar até que o elemento seja
            # visível na página, até o limite de timeout.
            wait = WebDriverWait(self.driver, timeout)
            campo_resumo = wait.until(EC.visibility_of_element_located(self.resumo_locator))
            
            texto_resumo = campo_resumo.text
            logging.info("Campo de resumo encontrado com sucesso!")
            logging.info(f"Texto extraído: '{texto_resumo[:100]}...'")
            return texto_resumo

        except TimeoutException:
            logging.error(f"Elemento 'Resumo' não foi encontrado na página após {timeout} segundos.")
            return None
        except Exception as e:
            logging.error(f"Ocorreu um erro inesperado ao tentar obter o resumo: {e}")
            return None

# ==============================================================================
# 3. BLOCO DE EXECUÇÃO PRINCIPAL
# ==============================================================================
if __name__ == "__main__":
    logging.info(">>> INICIANDO AUTOMAÇÃO ALM <<<")
    
    link_alm = obter_link_do_usuario()
    
    if link_alm:
        navegador = iniciar_navegador_com_perfil_usuario(link_alm)

        if navegador:
            try:
                # 1. Cria uma instância da nossa nova classe, passando o navegador
                pagina_alm = AlmPage(navegador)

                # 2. Chama o método para obter a informação desejada
                # Damos uma pausa inicial para a página carregar
                logging.info("Aguardando 5 segundos para a página carregar antes de procurar o resumo...")
                time.sleep(5) 
                
                resumo = pagina_alm.obter_resumo()

                # 3. Usa a informação
                if resumo:
                    pg.alert(f"O resumo encontrado foi:\n\n{resumo}", "Informação Extraída")
                else:
                    pg.alert("Não foi possível encontrar o campo de resumo na página.", "Erro de Extração")

                logging.info("A automação de demonstração será encerrada em 10 segundos.")
                time.sleep(10)

            except Exception as e:
                logging.error(f"Um erro ocorreu durante a automação: {e}")
            finally:
                logging.info("Fechando o navegador.")
                navegador.quit()

    logging.info(">>> FIM DA EXECUÇÃO <<<")
