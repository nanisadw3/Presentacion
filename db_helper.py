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
    c.execute('''CREATE TABLE IF NOT EXISTS modificaciones_generales (
                    proceso TEXT,
                    tabla TEXT,
                    clave TEXT,
                    valor REAL
                )''')
    c.execute('''CREATE TABLE IF NOT EXISTS coordenadas_override (
                    proceso TEXT PRIMARY KEY,
                    diaria_filas TEXT,
                    diaria_cols TEXT,
                    programa_filas TEXT,
                    programa_cols TEXT,
                    historica_filas TEXT,
                    historica_cols TEXT
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
    c.execute('DELETE FROM modificaciones_generales')
    c.execute('DELETE FROM coordenadas_override')
    conn.commit()
    conn.close()

def save_modificacion(proceso, tabla, clave, valor):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DELETE FROM modificaciones_generales WHERE proceso=? AND tabla=? AND clave=?', (proceso, tabla, clave))
    if valor is not None:
        c.execute('INSERT INTO modificaciones_generales VALUES (?, ?, ?, ?)', (proceso, tabla, clave, valor))
    conn.commit()
    conn.close()

def delete_modificacion(proceso, tabla, clave):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DELETE FROM modificaciones_generales WHERE proceso=? AND tabla=? AND clave=?', (proceso, tabla, clave))
    conn.commit()
    conn.close()

def get_modificaciones(proceso):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT tabla, clave, valor FROM modificaciones_generales WHERE proceso=?', (proceso,))
    rows = c.fetchall()
    conn.close()
    
    mods = {"diaria": {}, "programa": {}, "historica": {}}
    for r in rows:
        t, cl, v = r
        if t in mods:
            mods[t][cl] = v
    return mods

def save_coordenadas_override(proceso, d_filas, d_cols, p_filas, p_cols, h_filas, h_cols):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DELETE FROM coordenadas_override WHERE proceso=?', (proceso,))
    c.execute('INSERT INTO coordenadas_override VALUES (?, ?, ?, ?, ?, ?, ?)', 
              (proceso, d_filas, d_cols, p_filas, p_cols, h_filas, h_cols))
    conn.commit()
    conn.close()

def delete_coordenadas_override(proceso):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DELETE FROM coordenadas_override WHERE proceso=?', (proceso,))
    conn.commit()
    conn.close()

def get_coordenadas_override(proceso):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT diaria_filas, diaria_cols, programa_filas, programa_cols, historica_filas, historica_cols FROM coordenadas_override WHERE proceso=?', (proceso,))
    row = c.fetchone()
    conn.close()
    if row:
        return {
            "diaria_filas": row[0],
            "diaria_cols": row[1],
            "programa_filas": row[2],
            "programa_cols": row[3],
            "historica_filas": row[4],
            "historica_cols": row[5]
        }
    return None

init_db()
