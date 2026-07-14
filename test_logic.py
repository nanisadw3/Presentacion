import os
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import pandas as pd

def test_logic():
    df_prod_current = pd.DataFrame([
        ['2023', 0],
        ['Ene', 10],
        ['Mar', 30],
        ['2025', 0],
        ['Ene', 10],
        ['nan', None],
        ['', None]
    ], columns=["Año/Mes", "Produccion"])

    extra_rows = [
        ('2024', 'AÑO', 0),
        ('2024', 'Feb', 20),
        ('2023', 'Feb', 20),
        ('2025', 'Feb', 20)
    ]

    cleaned_data = []
    for _, row in df_prod_current.iterrows():
        val0 = str(row.iloc[0]).strip()
        if val0.lower() in ['', 'nan', 'nat', 'none']:
            continue
        cleaned_data.append((val0, row.iloc[1]))

    for r in extra_rows:
        anio, mes, prod = r
        anio = str(anio).strip()
        mes = str(mes).strip()
        
        year_idx = -1
        for i, (c_name, c_val) in enumerate(cleaned_data):
            if c_name == anio:
                year_idx = i
                break
                
        if year_idx == -1:
            insert_pos = 0
            for i, (c_name, c_val) in enumerate(cleaned_data):
                if c_name.isdigit() and len(c_name) == 4:
                    if int(c_name) > int(anio):
                        # Found a year greater than ours. insert_pos is here.
                        # Wait, we don't update insert_pos inside the loop unless we break.
                        # Actually we must track if we found it.
                        insert_pos = i
                        break
                    else:
                        insert_pos = len(cleaned_data)
            
            if insert_pos == 0 and len(cleaned_data) > 0 and not (cleaned_data[0][0].isdigit() and len(cleaned_data[0][0])==4):
                # if there's no year at all? just append.
                pass
                
            cleaned_data.insert(insert_pos, (anio, 0 if mes != "AÑO" else prod))
            year_idx = insert_pos
        else:
            if mes == "AÑO":
                cleaned_data[year_idx] = (anio, prod)
                
        if mes != "AÑO":
            meses_orden = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
            try:
                target_month_idx = meses_orden.index(mes.lower()[:3])
            except ValueError:
                target_month_idx = 99
                
            insert_idx = year_idx + 1
            replaced = False
            
            while insert_idx < len(cleaned_data):
                name_at_idx = cleaned_data[insert_idx][0]
                if name_at_idx.isdigit() and len(name_at_idx) == 4:
                    break
                    
                curr_month_str = name_at_idx.lower()[:3]
                try:
                    curr_m_idx = meses_orden.index(curr_month_str)
                except ValueError:
                    curr_m_idx = -1
                    
                if curr_m_idx == target_month_idx:
                    cleaned_data[insert_idx] = (mes, prod)
                    replaced = True
                    break
                elif curr_m_idx > target_month_idx:
                    break
                    
                insert_idx += 1
                
            if not replaced:
                cleaned_data.insert(insert_idx, (mes, prod))
                
    for row in cleaned_data:
        print(row)

test_logic()
