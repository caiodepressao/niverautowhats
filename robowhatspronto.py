import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import time 
from time import sleep
import random
from datetime import datetime
import phonenumbers
import traceback
import os


options = Options() #salva uma pasta com um perfil no chrome para carregar cookies
options.add_argument("--user-data-dir=C://Users//Pax Nacional//IAN//Interface")

                    # vai verificar o driver atual e sempre que inicializar ele compara e se desatualizado ele atualiza       
service = Service(ChromeDriverManager().install())

                    # Inicializar um novo driver do Chrome com a opção user-data-dir especificada e um objeto Service 
driver = webdriver.Chrome(service = service, options = options)

                    # Acessar o site do WhatsApp Web
driver.get("https://web.whatsapp.com")

                    # Obter a data atual do sistema
data_atual = datetime.now()

                    # Formatar a data como uma string no formato "dd-mm-aa"
data_str = data_atual.strftime("%d-%m-%y")

                    # Concatenar a data formatada com o sufixo "aniversario"
nome_arquivo = data_str + "aniversario.xlsx"

                    # Ler os dados do arquivo Excel
df = pd.read_excel(nome_arquivo)

                    # Converter a coluna de data de nascimento para o formato de data
df["Nascimento"] = pd.to_datetime(df["Nascimento"], format="%d/%m/%Y").dt.strftime("%d/%m")

                    # Data atual do sistema
today = time.strftime("%d/%m")

                    # Filtrar o DataFrame por data de nascimento igual à data atual
birthday_df = df[df["Nascimento"] == today]

                    # Função para enviar mensagem pelo WhatsApp
def send_whatsapp_message(driver, phone_number, message, nome_cliente=None):
    formatted_phone_number = "{:.0f}".format(phone_number)
       
                    # Acessar a página do WhatsApp Web
    driver.get("https://web.whatsapp.com")
    wait = WebDriverWait(driver, 600)
                    # Número máximo de tentativas para enviar a mensagem
    max_attempts = 2
                    # Número de tentativas para enviar a mensagem
    attempts = 0   

    while attempts < max_attempts:

        try:  
                        # Esperar pelo QR Code ser escaneado
             wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[4]/header/div[2]/div/span/div[1]/div/span')))
             time.sleep(5)

                        # Enviar a combinação de teclas CTRL + ALT + S
             actions = ActionChains(driver)
             actions.key_down(Keys.CONTROL).key_down(Keys.ALT).send_keys('s').key_up(Keys.ALT).key_up(Keys.CONTROL).perform()
             time.sleep(5)

                        # Digitar o número de telefone no campo de pesquisa após usar o ctrl+alt+s
             search_box = driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/div/input')
             for char in formatted_phone_number:
                 search_box.send_keys(char)
                 time.sleep(0.5)
             search_box.send_keys(Keys.ENTER)
             time.sleep(5)

                        # Localize o botão de anexo
             attachment_box = driver.find_element(By.XPATH, '//div[@title = "Anexar"]')
             attachment_box.click()
             time.sleep(2)

                        # Localize o botão de entrada de documento
             document_box = driver.find_element(By.XPATH, '//input[@accept="*"]')
             document_box.send_keys(r'C:\Users\Pax Nacional\IAN\Interface\boleto-teste.pdf')
             time.sleep(2)

                        # Clicar no botão de envio de mensagem
             send_button = driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[3]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span')
             send_button.click()
             time.sleep(5)
             break
                        # Obter data atual no formato desejado
        except Exception as e:
            data_atual = datetime.now().strftime("%d-%m-%y") 

                        # Nome do arquivo com a data atual e o sufixo "_com_erro.txt"
            nome_arquivo = f"{data_atual}numeros-com-erro.txt" 

                        # Adicionar número de telefone ao arquivo de texto
            with open(nome_arquivo, "a") as f: 
                if nome_cliente:
                    f.write(f"{formatted_phone_number} - {nome_cliente}\n")
                else:
                    f.write(f"{formatted_phone_number}\n")

                        # Pressionar ESC duas vezes e tentar novamente
            actions = ActionChains(driver) 
            actions.send_keys(Keys.ESCAPE).send_keys(Keys.ESCAPE).perform()
            attempts += 1

                        # manda a mensagem, se tirar a mensagem ele nao salva com nome no txt de erro
for index, row in df.iterrows():
    if row["Nascimento"] == today:
        phone = row["Telefone"]
        nome = row["Nome"]
        message = f"Olá {nome}, Nós da  , desejamos a você um feliz aniversário, tenha um dia especial !"
        send_whatsapp_message(driver, phone, message, nome)
        time.sleep(random.randint(2, 25))
                
driver.quit() # Fechar o driver do Chrome
