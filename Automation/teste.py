from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import os
import tkinter as tk
from tkinter import messagebox

# --- Configuração do Tkinter para a caixa de diálogo ---
# Cria uma janela raiz principal para o tkinter
root = tk.Tk()
# Esconde a janela raiz, pois só queremos a caixa de diálogo
root.withdraw()

# --- Configuração do Selenium ---
edge_options = EdgeOptions()
user_data_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Microsoft", "Edge", "User Data")
edge_options.add_argument(f"user-data-dir={user_data_dir}")

service = EdgeService(EdgeChromiumDriverManager().install())
driver = webdriver.Edge(service=service, options=edge_options)

driver.get("https://pontua.estaleiro.serpro.gov.br/pontua-web/#/contagem/detalhar/73746/MANUTENCAO")
driver.maximize_window()
try:
    tim
    elemento = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, "title-5 align-middle"))
    )
    print("📦 Texto extraído:", elemento.text)       
    # Define a mensagem de sucesso para a caixa de diálogo
    mensagem_final = "Processo concluído com sucesso!"

except Exception as e:
    print("❌ Erro ao localizar o elemento:", e)
    # Define a mensagem de erro para a caixa de diálogo
    mensagem_final = f"Ocorreu um erro:\n{e}"

finally:
    # Mostra a caixa de diálogo e espera o usuário clicar em OK
    messagebox.showinfo("Aviso", f"{mensagem_final}\n\nClique em OK para fechar o navegador.")
    
    # Após o clique em OK, o driver.quit() é executado
    driver.quit()