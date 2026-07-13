import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "historico_produccion.db")

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS produccion_extra (
                    proceso TEXT,
                    anio TEXT,
                    mes TEXT,
                    produccion REAL
                )''')
    conn.commit()
    conn.close()

def get_extra_prod(proceso_name):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT anio, mes, produccion FROM produccion_extra WHERE proceso=? ORDER BY anio, mes', (proceso_name,))
    rows = c.fetchall()
    conn.close()
    return rows

def save_extra_prod(proceso_name, anio, mes, produccion):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DELETE FROM produccion_extra WHERE proceso=? AND anio=? AND mes=?', (proceso_name, anio, mes))
    c.execute('INSERT INTO produccion_extra VALUES (?, ?, ?, ?)', (proceso_name, anio, mes, produccion))
    conn.commit()
    conn.close()

def clear_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DELETE FROM produccion_extra')
    conn.commit()
    conn.close()

init_db()
