import sqlite3

conn = sqlite3.connect("ocorrencias.db")
cursor = conn.cursor()

cursor.execute("DELETE FROM usuarios WHERE usuario = 'admin'")
conn.commit()
conn.close()

print("Usuário admin antigo apagado!")
