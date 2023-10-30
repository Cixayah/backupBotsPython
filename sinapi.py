from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Caminho para o ChromeDriver (substitua pelo seu caminho)
chrome_driver_path = r"C:\Users\Gabriel V Costa\Desktop\chrome-win\chromedriver.exe"

# Configurar as opções do Chrome
chrome_options = webdriver.ChromeOptions()
# Adicione as opções adicionais do Chrome, se necessário
# chrome_options.add_argument("--headless")  # Exemplo: Executar em modo headless

# Inicializar o serviço do ChromeDriver
driver_service = webdriver.chrome.service.Service(chrome_driver_path)

# Inicializar o driver com o serviço e as opções configuradas
driver = webdriver.Chrome(service=driver_service, options=chrome_options)

# URL da página que você deseja visitar
url = "https://www.caixa.gov.br/site/Paginas/downloads.aspx"

# Estados requeridos
estados = ["AL", "BA", "CE", "DF", "MA", "MG", "MS", "PA", "PB", "PE", "PR", "RJ", "RO", "SE", "SP"]

# Variável de data
mesAno = "092023"

# Iniciar a sessão do navegador e abrir a página
driver.get(url)

# Encontre o campo de busca
busca_textbox = driver.find_element(By.ID, "txtBusca")

# Iterar pelos estados e fazer a busca por cada estado
for estado in estados:
    termo_de_busca = f"{estado}_{mesAno}_NaoDesonerado"
    busca_textbox.clear()
    busca_textbox.send_keys(termo_de_busca)
    
    # Localize e clique no botão de busca usando o XPath fornecido
    search_button = driver.find_element(By.XPATH, "/html/body/form/div[4]/div/div/div[3]/div/span/div[1]/div[2]/div/div/div[1]/div[1]/div[1]/div/button/i")
    search_button.click()
    time.sleep(5)
    # Aguardar até que a página de resultados esteja completamente carregada (você pode ajustar o tempo limite)
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.XPATH, f"//a[contains(@href, '{termo_de_busca}')]")))
    
    # Encontrar todos os links que contêm o termo de busca no atributo href
    links = driver.find_elements(By.XPATH, f"//a[contains(@href, '{termo_de_busca}')]")
    
    # Iterar sobre os links e efetuar um clique com um intervalo de 1 segundo entre cada um
    for link in links:
        link.click()
        time.sleep(1)

# Fechar o navegador após a automação
time.sleep(30)
driver.quit()
