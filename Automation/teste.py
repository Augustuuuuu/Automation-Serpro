import os
import logging
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.keys import Keys # Importar Keys para usar atalhos

# --- Configuração do Logging ---
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def iniciar_navegador():
    """Configura e inicia o navegador Microsoft Edge com um perfil de usuário."""
    logging.info("🚀 Iniciando o navegador Microsoft Edge...")
    options = EdgeOptions()
    user_data_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Microsoft", "Edge", "User Data")
    options.add_argument(f"user-data-dir={user_data_dir}")
    
    service = EdgeService(EdgeChromiumDriverManager().install())
    driver = webdriver.Edge(service=service, options=options)
    driver.maximize_window()
    return driver

def executar_automacao(driver, url):
    """Executa a automação usando o atalho de teclado."""
    try:
        # 1. Acessar a URL
        logging.info(f"🔗 Acessando a URL: {url}")
        driver.get(url)
        
        wait = WebDriverWait(driver, 30)

        # 2. Preencher o campo de comentário
        logging.info("✍️ Procurando o campo de comentário...")
        xpath_comentario = "//div[contains(@class, 'RichTextEditorWidget') and contains(@aria-label, 'Coment')]"
        campo_comentario = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_comentario)))
        
        logging.info("⌨️ Inserindo o texto 'teste teste'...")
        campo_comentario.click()
        campo_comentario.send_keys(f"{"Olá"} {Keys.CONTROL + "l"}")
        logging.info("✅ Comentário inserido.")
        # 4. Clicar no botão OK
        logging.info("🖱️ Procurando e clicando no botão OK...")
        seletor_css_ok = "button[dojoattachpoint='_okButton']"
        botao_ok = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, seletor_css_ok)))
        botao_ok.click()
        logging.info("✅ Botão OK clicado com sucesso.")

        logging.info("\n🎉 Automação concluída com sucesso! 🎉")

    except Exception as e:
        logging.error(f"❌ Ocorreu um erro durante a automação: {e}")

    finally:
        input("\nPressione Enter para fechar o navegador...")
        driver.quit()
        logging.info("🏁 Navegador fechado.")


# --- Ponto de Entrada do Script ---
if __name__ == "__main__":
    url_alvo = "https://alm.serpro/ccm/web/projects/Gest%C3%A3o%20de%20Demandas%20Internas#action=com.ibm.team.workitem.viewWorkItem&id=4817285"
    
    navegador = iniciar_navegador()
    
    if navegador:
        executar_automacao(navegador, url_alvo)