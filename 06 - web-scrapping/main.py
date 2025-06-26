import requests
from bs4 import BeautifulSoup
import sqlite3
from datetime import datetime


moto = 'https://www.honda.com.br/motos/street/naked/cb-300f-twister'
headers = {'User-Agent': 'Mozilla/5.0'}  # boa prática pra evitar bloqueio

response = requests.get(moto)

soup = BeautifulSoup(response.text, 'html.parser')

preco = soup.find('strong', class_='price')

# conecta ou cria banco
conn = sqlite3.connect('precos.db')
c = conn.cursor()

# cria tabela se não existir
c.execute('''CREATE TABLE IF NOT EXISTS precos
             (data TEXT, url TEXT, preco TEXT)''')

# insere dados
c.execute("INSERT INTO precos VALUES (?, ?, ?)", 
          (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "CB300", preco.text))

conn.commit()
conn.close()