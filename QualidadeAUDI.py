import time
import pandas as pd
import urllib
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys  
from sqlalchemy import create_engine
from datetime import date

# Configurando o driver do Chrome
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

# URL e credenciais
url = "https://cem-portal-audi.ttr-group.de/login/auth"
login = "CEM_L1_Joyce.Milene"
senha = "Parvi.2024"

# Acessar a URL
navegador.get(url)
navegador.maximize_window()

def send_multiple_keys(navegador, key, times):
    for _ in range(times):
        navegador.switch_to.active_element.send_keys(key)
        time.sleep(1)

# Login
campo_login = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#username")))
campo_login.click()
campo_login.send_keys(login)
time.sleep(3)

campo_senha = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#password")))
campo_senha.click()
campo_senha.send_keys(senha)
time.sleep(3)

cliquelogin = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#submit")))
cliquelogin.click()
time.sleep(15)

send_multiple_keys(navegador, Keys.TAB, 3)
time.sleep(5)
send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(4)

nota = navegador.find_element(By.CSS_SELECTOR, "#overall-satisfaction_rating > div.panel-body > div").text
time.sleep(3)

# Criar DataFrame
dados = {'Satisfacao_Global': ["Audi"], 'Nota': [nota]}
df1 = pd.DataFrame(data=dados)
df3 = pd.DataFrame({"Segmento": ["Pos Vendas"]})
df = pd.concat([df1.reset_index(drop=True), df3.reset_index(drop=True)], axis=1)
df['data_atualizacao'] = date.today()

# Fechar o navegador
print("Fechando o navegador...")
print(df)
navegador.quit()
print("Navegador fechado com sucesso.")

# Conexão com banco de dados SQL Server
print("Conectando ao banco de dados...")
user = 'rpa_bi'
password = 'Rp@_B&_P@rvi'
host = '10.0.10.243'
port = '54949'
database = 'stage'

params = urllib.parse.quote_plus(
    f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={host},{port};DATABASE={database};UID={user};PWD={password}')
connection_str = f'mssql+pyodbc:///?odbc_connect={params}'
engine = create_engine(connection_str)
table_name = "QualidadeAudi"

# Inserir dados no banco
with engine.connect() as connection:
    df.to_sql(table_name, con=connection, if_exists='replace', index=False)

print(f"Dados inseridos com sucesso na tabela '{table_name}'!")

# Fechando o navegador
time.sleep(10)
print("Fechando o navegador...")
navegador.quit()
print("Navegador fechado com sucesso.")

# Fechar conexão com o banco
engine.dispose()
print("Conexão com o banco encerrada.")
