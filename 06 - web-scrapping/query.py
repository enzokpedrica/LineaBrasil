import sqlite3


path = 'C:\git\LineaBrasil\\06 - web-scrapping\precos.db'

conn = sqlite3.connect(path)
c = conn.cursor ()

for row in c.execute('SELECT * FROM precos'):
    print(row)

conn.close()