import pyautogui as pg
import logging
import time
import os
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.common.exceptions import WebDriverException

# ==============================================================================
# 1. SETUP DE LOGGING
# Configura um sistema de log limpo para exibir informações no console.
# ==============================================================================
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def obter_link_do_usuario():
    """
    Usa PyAutoGUI para exibir uma caixa de diálogo e solicitar o link ao usuário.
    É mais amigável que um input no terminal.

    Returns:
        str: A URL inserida pelo usuário, ou None se o usuário cancelar.
    """
    logging.info("Solicitando o link do ALM ao usuário...")
    titulo = "Automação ALM"
    texto = "Por favor, insira o link completo do ALM e clique em OK:"
    
    link_inserido = pg.prompt(text=texto, title=titulo)

    if link_inserido:
        logging.info(f"Link recebido: {link_inserido}")
        # Pequena validação para garantir que parece um link
        if not (link_inserido.startswith("http://") or link_inserido.startswith("https://")):
            link_inserido = "http://" + link_inserido
            logging.info(f"Adicionado 'http://' ao link: {link_inserido}")
        return link_inserido
    else:
        logging.warning("O usuário cancelou a caixa de diálogo. Encerrando o programa.")
        return None


def iniciar_navegador_com_perfil_usuario(url):
    """
    Inicia uma nova instância do Edge, mas carrega o perfil de usuário padrão,
    mantendo logins, cookies e sessões.
    Esta é a abordagem mais robusta e automática.

    Args:
        url (str): A URL completa do site a ser acessado.
    
    Returns:
        webdriver: A instância do driver do navegador, ou None se falhar.
    """
    if not url:
        return None

    try:
        logging.info("Configurando o driver para o Microsoft Edge com perfil de usuário...")
        
        edge_options = EdgeOptions()
        
        # --- MUDANÇA PRINCIPAL ---
        # Define o caminho para a pasta de perfil do usuário do Edge.
        # Isto faz com que o navegador inicie com os seus logins, cookies, etc.
        user_data_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Microsoft", "Edge", "User Data")
        edge_options.add_argument(f"user-data-dir={user_data_dir}")
        
        # Opcional: pode especificar um perfil específico se usar vários, ex: "Profile 1"
        # edge_options.add_argument("profile-directory=Default")
        
        # O webdriver-manager tratará do download/cache do driver correto.
        service = EdgeService(EdgeChromiumDriverManager().install())
        
        logging.info("Iniciando o navegador Edge com o seu perfil...")
        driver = webdriver.Edge(service=service, options=edge_options)
        
        logging.info(f"Navegando para a URL de destino: {url}")
        driver.get(url)
        driver.maximize_window()
        
        logging.info("Site acessado com sucesso! A automação continuará no navegador.")
        return driver

    except WebDriverException as e:
        logging.error(f"Ocorreu um erro do Selenium ao tentar iniciar o Edge: {e}")
        pg.alert("ERRO: Não foi possível iniciar o navegador Edge. Verifique se ele não está aberto em segundo plano.", "Erro de Automação")
        return None
    except Exception as e:
        logging.error(f"Um erro inesperado ocorreu: {e}")
        pg.alert(f"Ocorreu um erro inesperado: {e}", "Erro de Automação")
        return None


# ==============================================================================
# BLOCO DE EXECUÇÃO PRINCIPAL
# ==============================================================================
if __name__ == "__main__":
    logging.info(">>> INICIANDO AUTOMAÇÃO HÍBRIDA (PYAUTOGUI + SELENIUM) <<<")
    
    # Passo 1: Obter o link usando a interface gráfica do PyAutoGUI
    link_alm = obter_link_do_usuario()
    
    if link_alm:
        # Passo 2: Inicia o Edge com o perfil do usuário
        navegador = iniciar_navegador_com_perfil_usuario(link_alm)

        if navegador:
            try:
                # Agora o 'navegador' controla a janela do Edge com o seu perfil
                logging.info("A automação de login e outras tarefas podem começar aqui...")
                
                # Para demonstração, o script espera antes de fechar o navegador
                logging.info("O navegador será fechado em 300 segundos (5 minutos).")
                time.sleep(300)

            except Exception as e:
                logging.error(f"Um erro ocorreu durante a automação no navegador: {e}")
            finally:
                logging.info("Fechando o navegador iniciado pela automação.")
                navegador.quit()

    logging.info(">>> FIM DA EXECUÇÃO <<<")
