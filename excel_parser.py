import os
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import traceback
import pandas as pd
import db_helper

def load_data(app, file_path):
    try:
        import warnings
        warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

        # Función para eliminar filas donde la producción (última columna) sea 0 o vacía
        def filter_zero_rows(df):
            if df.empty: return df
            def is_not_zero(row):
                try:
                    val = row.iloc[-1]
                    if val == "" or val is None: return False
                    return float(val) != 0
                except:
                    return True
            return df[df.apply(is_not_zero, axis=1)]

        # Función segura para quitar decimales sin causar errores de tipo
        def remove_decimals(df_to_clean, skip_first=False, skip_last=False):
            def safe_round(val):
                try:
                    if pd.isna(val):
                        return ""
                except ValueError:
                    return ""
                try:
                    f_val = float(val)
                    import math
                    if math.isinf(f_val) or math.isnan(f_val):
                        return ""
                    return int(round(f_val))
                except:
                    return str(val).strip()

            def keep_raw(val):
                try:
                    if pd.isna(val):
                        return ""
                except ValueError:
                    return ""
                try:
                    f_val = float(val)
                    import math
                    if math.isinf(f_val) or math.isnan(f_val):
                        return ""
                    return val
                except:
                    return str(val).strip()

            if skip_first and df_to_clean.shape[1] > 1:
                first_col = df_to_clean.iloc[:, 0].map(keep_raw)
                rest = df_to_clean.iloc[:, 1:]
                if hasattr(rest, 'map'):
                    rest = rest.map(safe_round)
                else:
                    rest = rest.applymap(safe_round)
                return pd.concat([first_col, rest], axis=1)

            if skip_last and df_to_clean.shape[1] > 1:
                rest = df_to_clean.iloc[:, :-1]
                last_col = df_to_clean.iloc[:, -1].map(keep_raw)
                if hasattr(rest, 'map'):
                    rest = rest.map(safe_round)
                else:
                    rest = rest.applymap(safe_round)
                return pd.concat([rest, last_col], axis=1)

            if hasattr(df_to_clean, 'map'):
                return df_to_clean.map(safe_round)
            else:
                return df_to_clean.applymap(safe_round)

        # Autodetectar hoja
        xls = pd.ExcelFile(file_path)
        sheet_to_use = None
        for sheet in xls.sheet_names:
            normalizado = sheet.lower().replace(" ", "").replace("_", "")
            if "enviocalculopromedio" in normalizado:
                sheet_to_use = sheet
                break
        if not sheet_to_use:
            for sheet in xls.sheet_names:
                normalizado = sheet.lower().replace(" ", "").replace("_", "")
                if "enviocalculo" in normalizado:
                    sheet_to_use = sheet
                    break
        if not sheet_to_use:
            sheet_to_use = xls.sheet_names[0]

        # --- LEER LA HOJA COMPLETA UNA SOLA VEZ ---
        df_sheet = pd.read_excel(file_path, sheet_name=sheet_to_use, header=None)
        app.after(0, app.update_progress, 0.1, "Leyendo hoja de cálculo...")

        # Helper para formatear encabezados
        def get_clean_headers(row_idx, start_col, end_col):
            headers = []
            for val in df_sheet.iloc[row_idx, start_col:end_col]:
                if pd.isna(val):
                    headers.append("")
                else:
                    try:
                        headers.append(str(int(round(float(val)))))
                    except:
                        headers.append(str(val))
            return headers

        def merge_extra_prod(proceso_name, df_prod_current):
            extra_rows = db_helper.get_extra_prod(proceso_name)
            
            cleaned_data = []
            for _, row in df_prod_current.iterrows():
                val0 = str(row.iloc[0]).strip()
                if val0.endswith('.0'):
                    val0 = val0[:-2]
                if val0.lower() in ['', 'nan', 'nat', 'none']:
                    continue
                cleaned_data.append((val0, row.iloc[1]))

            if not extra_rows:
                if not cleaned_data:
                    return df_prod_current
                return pd.DataFrame(cleaned_data, columns=["Año/Mes", "Produccion"])

            for r in extra_rows:
                anio, mes, prod = r
                anio = str(anio).strip()
                mes = str(mes).strip()
                
                if mes == "AÑO":
                    insert_pos = 0
                    replaced = False
                    for i, (c_name, c_val) in enumerate(cleaned_data):
                        if c_name.isdigit() and len(c_name) == 4:
                            if c_name == anio:
                                cleaned_data[i] = (anio, prod)
                                replaced = True
                                break
                            elif int(c_name) > int(anio):
                                insert_pos = i
                                break
                            else:
                                insert_pos = i + 1
                        else:
                            insert_pos = i
                            break
                            
                    if not replaced:
                        cleaned_data.insert(insert_pos, (anio, prod))
                else:
                    month_block_start = len(cleaned_data)
                    for i, (c_name, c_val) in enumerate(cleaned_data):
                        if not (c_name.isdigit() and len(c_name) == 4):
                            month_block_start = i
                            break
                            
                    meses_orden = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
                    try:
                        target_month_idx = meses_orden.index(mes.lower()[:3])
                    except ValueError:
                        target_month_idx = 99
                        
                    insert_pos = month_block_start
                    replaced = False
                    
                    for i in range(month_block_start, len(cleaned_data)):
                        c_name = cleaned_data[i][0]
                        curr_month_str = c_name.lower()[:3]
                        try:
                            curr_m_idx = meses_orden.index(curr_month_str)
                        except ValueError:
                            curr_m_idx = -1
                            
                        if curr_m_idx == target_month_idx:
                            cleaned_data[i] = (mes, prod)
                            replaced = True
                            break
                        elif curr_m_idx > target_month_idx:
                            insert_pos = i
                            break
                        else:
                            insert_pos = i + 1
                            
                    if not replaced:
                        cleaned_data.insert(insert_pos, (mes, prod))
                        
            df_merged = pd.DataFrame(cleaned_data, columns=["Año/Mes", "Produccion"])
            return df_merged

        # --- 1. PROCESAR CRUDO (Cadereyta) ---
        app.after(0, app.update_progress, 0.2, "Procesando Crudo Cadereyta...")
        # --- 1. PROCESAR CRUDO (Estándar) ---
        app.after(0, app.update_progress, 0.2, "Procesando Crudo...")
        # Leer Tabla 1 (Rows 21-51 -> index 20:51), Cols A-H (0:8)
        headers_cad = get_clean_headers(0, 0, 8)
        df_data = df_sheet.iloc[20:51, 0:8].copy()
        df_data.columns = headers_cad
        df_data = df_data.dropna(how='all')
        df_data = remove_decimals(df_data)
        df_data = filter_zero_rows(df_data)

        # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Cols AE-AF (30:32)
        df_snr = df_sheet.iloc[73:104, 30:32].copy()
        df_snr.columns = ["CMP", "PODIM"]
        df_snr = df_snr.dropna(how='all').dropna(axis=1, how='all')
        df_snr = remove_decimals(df_snr, skip_first=True)
        df_snr_copy = df_snr.copy()

        # Leer Tabla 3 (Fecha y Producción, Rows 21-40 -> index 20:40), Cols AE-AF (30:32)
        df_prod_raw = df_sheet.iloc[20:40, 30:32].copy()
        df_prod_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_raw = df_prod_raw.dropna(how='all')
        
        dic_idx_crudo = -1
        for idx, row in df_prod_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_crudo = idx - 20
                break
        
        if dic_idx_crudo != -1:
            df_prod = df_prod_raw.iloc[:dic_idx_crudo + 1]
        else:
            df_prod = df_prod_raw.iloc[:20]
            
        df_prod = df_prod.dropna(axis=1, how='all')
        df_prod = remove_decimals(df_prod, skip_last=True)
        df_prod = merge_extra_prod("Crudo", df_prod)
        app.df_prod = df_prod.copy()
        df_prod_copy = df_prod.copy()

        num_dias_reales = len(df_data) if not df_data.empty else 31

        # --- 1.0 PROCESAR CRUDO CADEREYTA (Específico para Diapositiva 10) ---
        # Columnas A (0) y C (2), Filas 21-51
        df_data_cad = df_sheet.iloc[20:51, [0, 2]].copy()
        df_data_cad.columns = ["Crudo", "Cadereyta"]
        df_data_cad = df_data_cad.dropna(how='all')
        df_data_cad = remove_decimals(df_data_cad)
        df_data_cad = filter_zero_rows(df_data_cad)

        # Programa: Columna BI (60) repetida, Filas 74-104
        df_snr_cad = df_sheet.iloc[73:104, [60, 60]].copy()
        df_snr_cad.columns = ["CMP", "PODIM"]
        df_snr_cad = df_snr_cad.dropna(how='all').dropna(axis=1, how='all')
        df_snr_cad = remove_decimals(df_snr_cad, skip_first=True)
        df_snr_cad_copy = df_snr_cad.copy()
        
        # Fechas: Cols AW-AX (48:50), Filas 21-40
        df_prod_cad = df_sheet.iloc[20:40, 48:50].copy()
        df_prod_cad.columns = ["Año/Mes", "Produccion"]
        df_prod_cad = df_prod_cad.dropna(how='all')
        df_prod_cad = remove_decimals(df_prod_cad, skip_last=True)
        df_prod_cad = merge_extra_prod("Cadereyta -Crudo", df_prod_cad)
        df_prod_cad_copy = df_prod_cad.copy()
        
        # --- 1.1 PROCESAR GASOLINAS (Cadereyta) ---
        app.after(0, app.update_progress, 0.25, "Procesando Gasolinas Cadereyta...")
        headers_cad_gas = ["Cadereyta Gas - Día", "Cadereyta Gas - Producción"]
        # Leer Tabla 1 (Rows 22-51 -> index 21:51), Cols L y M (11, 12)
        df_gas_cad = df_sheet.iloc[21:51, [11, 12]].copy()
        df_gas_cad.columns = headers_cad_gas
        df_gas_cad = df_gas_cad.dropna(how='all')
        df_gas_cad = remove_decimals(df_gas_cad)
        df_gas_cad = filter_zero_rows(df_gas_cad)
        df_data_cad_gas = df_gas_cad.copy()

        # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Col BS repetida (70, 70)
        df_snr_cad_gas = df_sheet.iloc[73:104, [70, 70]].copy()
        df_snr_cad_gas = df_snr_cad_gas.dropna(how='all').dropna(axis=1, how='all')
        df_snr_cad_gas = remove_decimals(df_snr_cad_gas, skip_first=True)
        df_snr_cad_gas_copy = df_snr_cad_gas.copy()
        # Aplicar recorte de días reales basado en Crudo
        df_data_cad_gas = df_data_cad_gas.iloc[:num_dias_reales]

        # Leer Tabla 3 (Fecha y Producción Cadereyta Gas, Rows 21-51 -> index 20:51), Cols BO:BP (66:68)
        df_prod_cad_gas_raw = df_sheet.iloc[20:51, 66:68].copy()
        df_prod_cad_gas_raw = df_prod_cad_gas_raw.dropna(how='all')
        
        dic_idx_cad_gas = -1
        for idx, row in df_prod_cad_gas_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_cad_gas = idx - 20 
                break
        
        if dic_idx_cad_gas != -1:
            df_prod_cad_gas = df_prod_cad_gas_raw.iloc[:dic_idx_cad_gas + 1]
        else:
            df_prod_cad_gas = df_prod_cad_gas_raw.iloc[:31]
            
        df_prod_cad_gas = df_prod_cad_gas.dropna(axis=1, how='all')
        df_prod_cad_gas = remove_decimals(df_prod_cad_gas, skip_last=True)
        df_prod_cad_gas = merge_extra_prod("Cadereyta -Gasolinas", df_prod_cad_gas)
        df_prod_cad_gas_copy = df_prod_cad_gas.copy()

        # --- 1.2 PROCESAR DIESEL (Cadereyta) ---
        app.after(0, app.update_progress, 0.3, "Procesando Diesel Cadereyta...")
        headers_cad_die = ["Cadereyta Die - Día", "Cadereyta Die - Producción"]
        # Leer Tabla 1 (Rows 74-104 -> index 73:104), Cols A y B (0, 1)
        df_die_cad = df_sheet.iloc[73:104, [0, 1]].copy()
        df_die_cad.columns = headers_cad_die
        df_die_cad = df_die_cad.dropna(how='all')
        df_die_cad = remove_decimals(df_die_cad)
        df_die_cad = filter_zero_rows(df_die_cad)
        df_data_cad_die = df_die_cad.copy()

        # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Col CC repetida (80, 80)
        df_snr_cad_die = df_sheet.iloc[73:104, [80, 80]].copy()
        df_snr_cad_die = df_snr_cad_die.dropna(how='all').dropna(axis=1, how='all')
        df_snr_cad_die = remove_decimals(df_snr_cad_die, skip_first=True)
        df_snr_cad_die_copy = df_snr_cad_die.copy()
        # Aplicar recorte de días reales basado en Crudo
        df_data_cad_die = df_data_cad_die.iloc[:num_dias_reales]

        # Leer Tabla 3 (Fecha y Producción Cadereyta Die, Rows 21-40 -> index 20:40), Cols CG:CH (84:86)
        df_prod_cad_die_raw = df_sheet.iloc[20:40, 84:86].copy()
        df_prod_cad_die_raw = df_prod_cad_die_raw.dropna(how='all')
        
        dic_idx_cad_die = -1
        for idx, row in df_prod_cad_die_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_cad_die = idx - 20 
                break
        
        if dic_idx_cad_die != -1:
            df_prod_cad_die = df_prod_cad_die_raw.iloc[:dic_idx_cad_die + 1]
        else:
            df_prod_cad_die = df_prod_cad_die_raw.iloc[:20]
            
        df_prod_cad_die = df_prod_cad_die.dropna(axis=1, how='all')
        df_prod_cad_die = remove_decimals(df_prod_cad_die, skip_last=True)
        df_prod_cad_die = merge_extra_prod("Cadereyta -Diesel", df_prod_cad_die)
        df_prod_cad_die_copy = df_prod_cad_die.copy()

        # --- 1.4 PROCESAR COMBUSTOLEO (Cadereyta) ---
        app.after(0, app.update_progress, 0.35, "Procesando Combustoleo Cadereyta...")
        headers_cad_comb = ["Cadereyta Comb - Día", "Cadereyta Comb - Producción"]
        df_comb_cad = df_sheet.iloc[158:189, [18, 19]].copy()
        df_comb_cad.columns = headers_cad_comb
        df_comb_cad = df_comb_cad.dropna(how='all')
        df_comb_cad = remove_decimals(df_comb_cad)
        df_comb_cad = filter_zero_rows(df_comb_cad)
        df_data_cad_comb = df_comb_cad.copy()

        df_snr_cad_comb = df_sheet.iloc[73:104, [98, 98]].copy()
        df_snr_cad_comb = df_snr_cad_comb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_cad_comb = remove_decimals(df_snr_cad_comb, skip_first=True)
        df_snr_cad_comb_copy = df_snr_cad_comb.copy()
        df_data_cad_comb = df_data_cad_comb.iloc[:num_dias_reales]

        df_prod_cad_comb_raw = df_sheet.iloc[157:179, 30:32].copy()
        df_prod_cad_comb_raw = df_prod_cad_comb_raw.dropna(how='all')
        
        dic_idx_cad_comb = -1
        for idx, row in df_prod_cad_comb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_cad_comb = idx - 157 
                break
        
        if dic_idx_cad_comb != -1:
            df_prod_cad_comb = df_prod_cad_comb_raw.iloc[:dic_idx_cad_comb + 1]
        else:
            df_prod_cad_comb = df_prod_cad_comb_raw.iloc[:20]
            
        df_prod_cad_comb = df_prod_cad_comb.dropna(axis=1, how='all')
        df_prod_cad_comb = remove_decimals(df_prod_cad_comb, skip_last=True)
        df_prod_cad_comb = merge_extra_prod("Cadereyta -Combustoleo", df_prod_cad_comb)
        df_prod_cad_comb_copy = df_prod_cad_comb.copy()



        # --- 1.5 PROCESAR CRUDO MADERO ---
        app.after(0, app.update_progress, 0.36, "Procesando Crudo Madero...")
        df_data_mad_crud = df_sheet.iloc[20:51, [0, 3]].copy()
        df_data_mad_crud.columns = ["Crudo Día", "Madero Crudo"]
        df_data_mad_crud = df_data_mad_crud.dropna(how='all')
        df_data_mad_crud = remove_decimals(df_data_mad_crud)
        df_data_mad_crud = df_data_mad_crud.iloc[:num_dias_reales]

        df_snr_mad_crud = df_sheet.iloc[73:104, [61, 61]].copy()
        df_snr_mad_crud.columns = ["CMP", "PODIM"]
        df_snr_mad_crud = df_snr_mad_crud.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mad_crud = remove_decimals(df_snr_mad_crud, skip_first=True)
        df_snr_mad_crud_copy = df_snr_mad_crud.copy()

        df_prod_mad_crud_raw = df_sheet.iloc[20:40, 50:52].copy()
        df_prod_mad_crud_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_mad_crud_raw = df_prod_mad_crud_raw.dropna(how='all')

        dic_idx_mad_crud = -1
        for idx, row in df_prod_mad_crud_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_mad_crud = idx - 20
                break

        if dic_idx_mad_crud != -1:
            df_prod_mad_crud = df_prod_mad_crud_raw.iloc[:dic_idx_mad_crud + 1]
        else:
            df_prod_mad_crud = df_prod_mad_crud_raw.iloc[:20]

        df_prod_mad_crud = df_prod_mad_crud.dropna(axis=1, how='all')
        df_prod_mad_crud = remove_decimals(df_prod_mad_crud, skip_last=True)
        df_prod_mad_crud = merge_extra_prod("Madero -Crudo", df_prod_mad_crud)
        df_prod_mad_crud_copy = df_prod_mad_crud.copy()

        # --- 1.6 PROCESAR GASOLINAS MADERO ---
        app.after(0, app.update_progress, 0.37, "Procesando Gasolinas Madero...")
        df_data_mad_gas = df_sheet.iloc[20:51, [11, 13]].copy()
        df_data_mad_gas.columns = ["Gas Día", "Madero Gas"]
        df_data_mad_gas = df_data_mad_gas.dropna(how='all')
        df_data_mad_gas = remove_decimals(df_data_mad_gas)
        df_data_mad_gas = df_data_mad_gas.iloc[:num_dias_reales]

        df_snr_mad_gas = df_sheet.iloc[73:104, [71, 71]].copy()
        df_snr_mad_gas.columns = ["CMP", "PODIM"]
        df_snr_mad_gas = df_snr_mad_gas.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mad_gas = remove_decimals(df_snr_mad_gas, skip_first=True)
        df_snr_mad_gas_copy = df_snr_mad_gas.copy()

        df_prod_mad_gas_raw = df_sheet.iloc[20:40, 70:72].copy()
        df_prod_mad_gas_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_mad_gas_raw = df_prod_mad_gas_raw.dropna(how='all')

        dic_idx_mad_gas = -1
        for idx, row in df_prod_mad_gas_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_mad_gas = idx - 20
                break

        if dic_idx_mad_gas != -1:
            df_prod_mad_gas = df_prod_mad_gas_raw.iloc[:dic_idx_mad_gas + 1]
        else:
            df_prod_mad_gas = df_prod_mad_gas_raw.iloc[:20]

        df_prod_mad_gas = df_prod_mad_gas.dropna(axis=1, how='all')
        df_prod_mad_gas = remove_decimals(df_prod_mad_gas, skip_last=True)
        df_prod_mad_gas = merge_extra_prod("Madero -Gasolinas", df_prod_mad_gas)
        df_prod_mad_gas_copy = df_prod_mad_gas.copy()

        # --- 1.7 PROCESAR DIESEL MADERO ---
        app.after(0, app.update_progress, 0.38, "Procesando Diesel Madero...")
        df_data_mad_die = df_sheet.iloc[73:104, [0, 2]].copy()
        df_data_mad_die.columns = ["Diesel Día", "Madero Die"]
        df_data_mad_die = df_data_mad_die.dropna(how='all')
        df_data_mad_die = remove_decimals(df_data_mad_die)
        df_data_mad_die = df_data_mad_die.iloc[:num_dias_reales]

        df_snr_mad_die = df_sheet.iloc[73:104, [81, 81]].copy()
        df_snr_mad_die.columns = ["CMP", "PODIM"]
        df_snr_mad_die = df_snr_mad_die.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mad_die = remove_decimals(df_snr_mad_die, skip_first=True)
        df_snr_mad_die_copy = df_snr_mad_die.copy()

        df_prod_mad_die_raw = df_sheet.iloc[20:40, 86:88].copy()
        df_prod_mad_die_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_mad_die_raw = df_prod_mad_die_raw.dropna(how='all')

        dic_idx_mad_die = -1
        for idx, row in df_prod_mad_die_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_mad_die = idx - 20
                break

        if dic_idx_mad_die != -1:
            df_prod_mad_die = df_prod_mad_die_raw.iloc[:dic_idx_mad_die + 1]
        else:
            df_prod_mad_die = df_prod_mad_die_raw.iloc[:20]

        df_prod_mad_die = df_prod_mad_die.dropna(axis=1, how='all')
        df_prod_mad_die = remove_decimals(df_prod_mad_die, skip_last=True)
        df_prod_mad_die = merge_extra_prod("Madero -Diesel", df_prod_mad_die)
        df_prod_mad_die_copy = df_prod_mad_die.copy()

        # --- 1.8 PROCESAR TURBOSINA MADERO ---
        app.after(0, app.update_progress, 0.39, "Procesando Turbosina Madero...")
        df_data_mad_turb = df_sheet.iloc[73:104, [11, 12]].copy()
        df_data_mad_turb.columns = ["Turb Día", "Madero Turb"]
        df_data_mad_turb = df_data_mad_turb.dropna(how='all')
        df_data_mad_turb = remove_decimals(df_data_mad_turb)
        df_data_mad_turb = df_data_mad_turb.iloc[:num_dias_reales]

        df_snr_mad_turb = df_sheet.iloc[73:104, [90, 90]].copy()
        df_snr_mad_turb.columns = ["CMP", "PODIM"]
        df_snr_mad_turb = df_snr_mad_turb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mad_turb = remove_decimals(df_snr_mad_turb, skip_first=True)
        df_snr_mad_turb_copy = df_snr_mad_turb.copy()

        df_prod_mad_turb_raw = df_sheet.iloc[20:40, 101:103].copy()
        df_prod_mad_turb_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_mad_turb_raw = df_prod_mad_turb_raw.dropna(how='all')

        dic_idx_mad_turb = -1
        for idx, row in df_prod_mad_turb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_mad_turb = idx - 20
                break

        if dic_idx_mad_turb != -1:
            df_prod_mad_turb = df_prod_mad_turb_raw.iloc[:dic_idx_mad_turb + 1]
        else:
            df_prod_mad_turb = df_prod_mad_turb_raw.iloc[:20]

        df_prod_mad_turb = df_prod_mad_turb.dropna(axis=1, how='all')
        df_prod_mad_turb = remove_decimals(df_prod_mad_turb, skip_last=True)
        df_prod_mad_turb = merge_extra_prod("Madero -Turbosina", df_prod_mad_turb)
        df_prod_mad_turb_copy = df_prod_mad_turb.copy()

        # --- 1.9 PROCESAR COMBUSTOLEO MADERO ---
        app.after(0, app.update_progress, 0.40, "Procesando Combustoleo Madero...")
        df_data_mad_comb = df_sheet.iloc[158:189, [18, 20]].copy()
        df_data_mad_comb.columns = ["Comb Día", "Madero Comb"]
        df_data_mad_comb = df_data_mad_comb.dropna(how='all')
        df_data_mad_comb = remove_decimals(df_data_mad_comb)
        df_data_mad_comb = df_data_mad_comb.iloc[:num_dias_reales]

        df_snr_mad_comb = df_sheet.iloc[73:104, [99, 99]].copy()
        df_snr_mad_comb.columns = ["CMP", "PODIM"]
        df_snr_mad_comb = df_snr_mad_comb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mad_comb = remove_decimals(df_snr_mad_comb, skip_first=True)
        df_snr_mad_comb_copy = df_snr_mad_comb.copy()

        df_prod_mad_comb_raw = df_sheet.iloc[157:179, 36:38].copy()
        df_prod_mad_comb_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_mad_comb_raw = df_prod_mad_comb_raw.dropna(how='all')

        dic_idx_mad_comb = -1
        for idx, row in df_prod_mad_comb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_mad_comb = idx - 157
                break

        if dic_idx_mad_comb != -1:
            df_prod_mad_comb = df_prod_mad_comb_raw.iloc[:dic_idx_mad_comb + 1]
        else:
            df_prod_mad_comb = df_prod_mad_comb_raw.iloc[:22]

        df_prod_mad_comb = df_prod_mad_comb.dropna(axis=1, how='all')
        df_prod_mad_comb = remove_decimals(df_prod_mad_comb, skip_last=True)
        df_prod_mad_comb = merge_extra_prod("Madero -Combustoleo", df_prod_mad_comb)
        df_prod_mad_comb_copy = df_prod_mad_comb.copy()

        # --- 1.10 PROCESAR CRUDO MINATITLAN ---
        app.after(0, app.update_progress, 0.41, "Procesando Crudo Minatitlan...")
        df_data_mina_crud = df_sheet.iloc[20:51, [0, 4]].copy()
        df_data_mina_crud.columns = ["Crudo Día", "Minatitlan Crudo"]
        df_data_mina_crud = df_data_mina_crud.dropna(how='all')
        df_data_mina_crud = remove_decimals(df_data_mina_crud)
        df_data_mina_crud = df_data_mina_crud.iloc[:num_dias_reales]

        df_snr_mina_crud = df_sheet.iloc[73:104, [62, 62]].copy()
        df_snr_mina_crud.columns = ["CMP", "PODIM"]
        df_snr_mina_crud = df_snr_mina_crud.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mina_crud = remove_decimals(df_snr_mina_crud, skip_first=True)
        df_snr_mina_crud_copy = df_snr_mina_crud.copy()

        df_prod_mina_crud_raw = df_sheet.iloc[20:40, 54:56].copy()
        df_prod_mina_crud_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_mina_crud_raw = df_prod_mina_crud_raw.dropna(how='all')

        dic_idx_mina_crud = -1
        for idx, row in df_prod_mina_crud_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_mina_crud = idx - 20
                break

        if dic_idx_mina_crud != -1:
            df_prod_mina_crud = df_prod_mina_crud_raw.iloc[:dic_idx_mina_crud + 1]
        else:
            df_prod_mina_crud = df_prod_mina_crud_raw.iloc[:20]

        df_prod_mina_crud = df_prod_mina_crud.dropna(axis=1, how='all')
        df_prod_mina_crud = remove_decimals(df_prod_mina_crud, skip_last=True)
        df_prod_mina_crud = merge_extra_prod("Minatitlan -Crudo", df_prod_mina_crud)
        df_prod_mina_crud_copy = df_prod_mina_crud.copy()

        # --- 1.11 PROCESAR GASOLINAS MINATITLAN ---
        app.after(0, app.update_progress, 0.42, "Procesando Gasolinas Minatitlan...")
        df_data_mina_gas = df_sheet.iloc[20:51, [11, 14]].copy()
        df_data_mina_gas.columns = ["Gas Día", "Minatitlan Gas"]
        df_data_mina_gas = df_data_mina_gas.dropna(how='all')
        df_data_mina_gas = remove_decimals(df_data_mina_gas)
        df_data_mina_gas = df_data_mina_gas.iloc[:num_dias_reales]

        df_snr_mina_gas = df_sheet.iloc[73:104, [72, 72]].copy()
        df_snr_mina_gas.columns = ["CMP", "PODIM"]
        df_snr_mina_gas = df_snr_mina_gas.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mina_gas = remove_decimals(df_snr_mina_gas, skip_first=True)
        df_snr_mina_gas_copy = df_snr_mina_gas.copy()

        df_prod_mina_gas_raw = df_sheet.iloc[20:40, 72:74].copy()
        df_prod_mina_gas_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_mina_gas_raw = df_prod_mina_gas_raw.dropna(how='all')

        dic_idx_mina_gas = -1
        for idx, row in df_prod_mina_gas_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_mina_gas = idx - 20
                break

        if dic_idx_mina_gas != -1:
            df_prod_mina_gas = df_prod_mina_gas_raw.iloc[:dic_idx_mina_gas + 1]
        else:
            df_prod_mina_gas = df_prod_mina_gas_raw.iloc[:20]

        df_prod_mina_gas = df_prod_mina_gas.dropna(axis=1, how='all')
        df_prod_mina_gas = remove_decimals(df_prod_mina_gas, skip_last=True)
        df_prod_mina_gas = merge_extra_prod("Minatitlan -Gasolinas", df_prod_mina_gas)
        df_prod_mina_gas_copy = df_prod_mina_gas.copy()

        # --- 1.12 PROCESAR DIESEL MINATITLAN ---
        app.after(0, app.update_progress, 0.43, "Procesando Diesel Minatitlan...")
        df_data_mina_die = df_sheet.iloc[73:104, [0, 3]].copy()
        df_data_mina_die.columns = ["Diesel Día", "Minatitlan Die"]
        df_data_mina_die = df_data_mina_die.dropna(how='all')
        df_data_mina_die = remove_decimals(df_data_mina_die)
        df_data_mina_die = df_data_mina_die.iloc[:num_dias_reales]

        df_snr_mina_die = df_sheet.iloc[73:104, [82, 82]].copy()
        df_snr_mina_die.columns = ["CMP", "PODIM"]
        df_snr_mina_die = df_snr_mina_die.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mina_die = remove_decimals(df_snr_mina_die, skip_first=True)
        df_snr_mina_die_copy = df_snr_mina_die.copy()

        df_prod_mina_die_raw = df_sheet.iloc[20:40, 90:92].copy()
        df_prod_mina_die_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_mina_die_raw = df_prod_mina_die_raw.dropna(how='all')

        dic_idx_mina_die = -1
        for idx, row in df_prod_mina_die_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_mina_die = idx - 20
                break

        if dic_idx_mina_die != -1:
            df_prod_mina_die = df_prod_mina_die_raw.iloc[:dic_idx_mina_die + 1]
        else:
            df_prod_mina_die = df_prod_mina_die_raw.iloc[:20]

        df_prod_mina_die = df_prod_mina_die.dropna(axis=1, how='all')
        df_prod_mina_die = remove_decimals(df_prod_mina_die, skip_last=True)
        df_prod_mina_die = merge_extra_prod("Minatitlan -Diesel", df_prod_mina_die)
        df_prod_mina_die_copy = df_prod_mina_die.copy()

        # --- 1.13 PROCESAR COMBUSTOLEO MINATITLAN ---
        app.after(0, app.update_progress, 0.44, "Procesando Combustoleo Minatitlan...")
        df_data_mina_comb = df_sheet.iloc[158:189, [18, 21]].copy()
        df_data_mina_comb.columns = ["Comb Día", "Minatitlan Comb"]
        df_data_mina_comb = df_data_mina_comb.dropna(how='all')
        df_data_mina_comb = remove_decimals(df_data_mina_comb)
        df_data_mina_comb = df_data_mina_comb.iloc[:num_dias_reales]

        df_snr_mina_comb = df_sheet.iloc[73:104, [100, 100]].copy()
        df_snr_mina_comb.columns = ["CMP", "PODIM"]
        df_snr_mina_comb = df_snr_mina_comb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mina_comb = remove_decimals(df_snr_mina_comb, skip_first=True)
        df_snr_mina_comb_copy = df_snr_mina_comb.copy()

        df_prod_mina_comb_raw = df_sheet.iloc[157:179, 39:41].copy()
        df_prod_mina_comb_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_mina_comb_raw = df_prod_mina_comb_raw.dropna(how='all')

        dic_idx_mina_comb = -1
        for idx, row in df_prod_mina_comb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_mina_comb = idx - 157
                break

        if dic_idx_mina_comb != -1:
            df_prod_mina_comb = df_prod_mina_comb_raw.iloc[:dic_idx_mina_comb + 1]
        else:
            df_prod_mina_comb = df_prod_mina_comb_raw.iloc[:22]

        df_prod_mina_comb = df_prod_mina_comb.dropna(axis=1, how='all')
        df_prod_mina_comb = remove_decimals(df_prod_mina_comb, skip_last=True)
        df_prod_mina_comb = merge_extra_prod("Minatitlan -Combustoleo", df_prod_mina_comb)
        df_prod_mina_comb_copy = df_prod_mina_comb.copy()

        # --- 2. PROCESAR GASOLINAS ---
        app.after(0, app.update_progress, 0.35, "Procesando Gasolinas...")
        # Leer fila 1 (index 0) para encabezados, Cols L:S (11:19)
        clean_headers_gas = get_clean_headers(0, 11, 19)
        # Leer Tabla 1 (Rows 21-51 -> index 20:51), Cols L:S (11:19)
        df_gas = df_sheet.iloc[20:51, 11:19].copy()
        df_gas.columns = clean_headers_gas
        df_gas = df_gas.dropna(how='all').dropna(axis=1, how='all')
        df_gas = remove_decimals(df_gas)
        df_data_gasolinas = filter_zero_rows(df_gas).copy()

        # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Cols AK:AL (36:38)
        df_snr_gas = df_sheet.iloc[73:104, 36:38].copy()
        df_snr_gas = df_snr_gas.dropna(how='all').dropna(axis=1, how='all')
        df_snr_gas = remove_decimals(df_snr_gas, skip_first=True)
        df_snr_gas_copy = df_snr_gas.copy()

        # Recortar filas de fin de mes
        df_data_gasolinas = df_data_gasolinas.iloc[:num_dias_reales]

        # Leer Tabla 3 (Años, Rows 21-120 -> index 20:120), Cols AG:AH (32:34)
        df_prod_gas_raw = df_sheet.iloc[20:120, 32:34].copy()
        df_prod_gas_raw = df_prod_gas_raw.dropna(how='all')
        
        dic_idx_gas = -1
        for idx, row in df_prod_gas_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_gas = idx - 20
                break
        
        if dic_idx_gas != -1:
            df_prod_gas = df_prod_gas_raw.iloc[:dic_idx_gas + 1]
        else:
            df_prod_gas = df_prod_gas_raw.iloc[:20]
            
        df_prod_gas = df_prod_gas.dropna(axis=1, how='all')
        df_prod_gas = remove_decimals(df_prod_gas, skip_last=True)
        df_prod_gas = merge_extra_prod("Gasolinas", df_prod_gas)
        df_prod_gasolinas_copy = df_prod_gas.copy()


        # --- 3. PROCESAR DIESEL ---
        app.after(0, app.update_progress, 0.5, "Procesando Diesel...")
        # Leer fila 54 (index 53) para encabezados, Cols A:H (0:8)
        clean_headers_die = get_clean_headers(53, 0, 8)
        # Leer Tabla 1 (Rows 74-104 -> index 73:104), Cols A:H (0:8)
        df_die = df_sheet.iloc[73:104, 0:8].copy()
        df_die.columns = clean_headers_die
        df_die = df_die.dropna(how='all').dropna(axis=1, how='all')
        df_die = remove_decimals(df_die)
        df_data_diesel = filter_zero_rows(df_die).copy()

        # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Cols AQ:AR (42:44)
        df_snr_die = df_sheet.iloc[73:104, 42:44].copy()
        df_snr_die = df_snr_die.dropna(how='all').dropna(axis=1, how='all')
        df_snr_die = remove_decimals(df_snr_die, skip_first=True)
        df_snr_die_copy = df_snr_die.copy()

        # Recortar filas de fin de mes
        df_data_diesel = df_data_diesel.iloc[:num_dias_reales]

        # Leer Tabla 3 (Años, Rows 21-120 -> index 20:120), Cols AK:AL (36:38)
        df_prod_die_raw = df_sheet.iloc[20:120, 36:38].copy()
        df_prod_die_raw = df_prod_die_raw.dropna(how='all')
        
        dic_idx_die = -1
        for idx, row in df_prod_die_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_die = idx - 20
                break
        
        if dic_idx_die != -1:
            df_prod_die = df_prod_die_raw.iloc[:dic_idx_die + 1]
        else:
            df_prod_die = df_prod_die_raw.iloc[:20]
            
        df_prod_die = df_prod_die.dropna(axis=1, how='all')
        df_prod_die = remove_decimals(df_prod_die, skip_last=True)
        df_prod_die = merge_extra_prod("Diesel", df_prod_die)
        df_prod_diesel_copy = df_prod_die.copy()


        # --- 4. PROCESAR TURBOSINA ---
        app.after(0, app.update_progress, 0.62, "Procesando Turbosina...")
        # Leer fila 54 (index 53) para encabezados, Cols L:Q (11:17)
        clean_headers_turb = get_clean_headers(53, 11, 17)
        # Leer Tabla 1 (Rows 74-104 -> index 73:104), Cols L:Q (11:17)
        df_turb = df_sheet.iloc[73:104, 11:17].copy()
        df_turb.columns = clean_headers_turb
        df_turb = df_turb.dropna(how='all').dropna(axis=1, how='all')
        df_turb = remove_decimals(df_turb)
        df_data_turbosina = filter_zero_rows(df_turb).copy()

        # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Cols AW:AX (48:50)
        df_snr_turb = df_sheet.iloc[73:104, 48:50].copy()
        df_snr_turb = df_snr_turb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_turb = remove_decimals(df_snr_turb, skip_first=True)
        df_snr_turb_copy = df_snr_turb.copy()

        # Recortar filas de fin de mes
        df_data_turbosina = df_data_turbosina.iloc[:num_dias_reales]

        # Leer Tabla 3 (Años, Rows 21-120 -> index 20:120), Cols AM:AN (38:40)
        df_prod_turb_raw = df_sheet.iloc[20:120, 38:40].copy()
        df_prod_turb_raw = df_prod_turb_raw.dropna(how='all')
        
        dic_idx_turb = -1
        for idx, row in df_prod_turb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_turb = idx - 20
                break
        
        if dic_idx_turb != -1:
            df_prod_turb = df_prod_turb_raw.iloc[:dic_idx_turb + 1]
        else:
            df_prod_turb = df_prod_turb_raw.iloc[:20]
            
        df_prod_turb = df_prod_turb.dropna(axis=1, how='all')
        df_prod_turb = remove_decimals(df_prod_turb, skip_last=True)
        df_prod_turb = merge_extra_prod("Turbosina", df_prod_turb)
        df_prod_turbosina_copy = df_prod_turb.copy()


        # --- 5. PROCESAR ASFALTO ---
        app.after(0, app.update_progress, 0.74, "Procesando Asfalto...")
        # Leer Tabla 1 (Rows 122-152 -> index 121:152), Cols AM:AN (38:40)
        df_asf = df_sheet.iloc[121:152, 38:40].copy()
        df_asf.columns = ["Asfalto", "real"]
        df_asf = df_asf.dropna(how='all').dropna(axis=1, how='all')
        df_asf = remove_decimals(df_asf)
        df_data_asfalto = filter_zero_rows(df_asf).copy()

        # Leer Tabla 2 (Programa, Rows 122-152 -> index 121:152), Cols AK:AL (36:38)
        df_snr_asf = df_sheet.iloc[121:152, 36:38].copy()
        df_snr_asf = df_snr_asf.dropna(how='all').dropna(axis=1, how='all')
        df_snr_asf = remove_decimals(df_snr_asf, skip_first=True)
        df_snr_asf_copy = df_snr_asf.copy()

        # Recortar filas de fin de mes
        df_data_asfalto = df_data_asfalto.iloc[:num_dias_reales]

        # Leer Tabla 3 (Años, Rows 121-140 -> index 120:140), Cols AQ:AR (42:44)
        df_prod_asf_raw = df_sheet.iloc[120:140, 42:44].copy()
        df_prod_asf_raw = df_prod_asf_raw.dropna(how='all')

        dic_idx_asf = -1
        for idx, row in df_prod_asf_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_asf = idx - 120
                break

        if dic_idx_asf != -1:
            df_prod_asf = df_prod_asf_raw.iloc[:dic_idx_asf + 1]
        else:
            df_prod_asf = df_prod_asf_raw.iloc[:20]

        df_prod_asf = df_prod_asf.dropna(axis=1, how='all')
        df_prod_asf = remove_decimals(df_prod_asf, skip_last=True)
        df_prod_asf = merge_extra_prod("Asfalto", df_prod_asf)
        df_prod_asfalto_copy = df_prod_asf.copy()


        # --- 6. PROCESAR COMBUSTOLEO ---
        app.after(0, app.update_progress, 0.85, "Procesando Combustoleo...")
        # Leer Tabla 1 (Rows 74-104 -> index 73:104), Cols U:V (20:22)
        df_comb = df_sheet.iloc[73:104, 20:22].copy()
        df_comb.columns = ["SNR", "Combustoleo"]
        df_comb = df_comb.dropna(how='all').dropna(axis=1, how='all')
        df_comb = remove_decimals(df_comb)
        df_data_combustoleo = filter_zero_rows(df_comb).copy()

        # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Cols BC:BD (54:56)
        df_snr_comb = df_sheet.iloc[73:104, 54:56].copy()
        df_snr_comb = df_snr_comb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_comb = remove_decimals(df_snr_comb, skip_first=True)
        df_snr_comb_copy = df_snr_comb.copy()

        # Recortar filas de fin de mes
        df_data_combustoleo = df_data_combustoleo.iloc[:num_dias_reales]

        # Leer Tabla 3 (Años, Rows 21-40 -> index 20:40), Cols AQ:AR (42:44)
        df_prod_comb_raw = df_sheet.iloc[20:40, 42:44].copy()
        df_prod_comb_raw = df_prod_comb_raw.dropna(how='all')

        dic_idx_comb = -1
        for idx, row in df_prod_comb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_comb = idx - 20
                break

        if dic_idx_comb != -1:
            df_prod_comb = df_prod_comb_raw.iloc[:dic_idx_comb + 1]
        else:
            df_prod_comb = df_prod_comb_raw.iloc[:20]

        df_prod_comb = df_prod_comb.dropna(axis=1, how='all')
        df_prod_comb = remove_decimals(df_prod_comb, skip_last=True)
        df_prod_comb = merge_extra_prod("Combustoleo", df_prod_comb)
        df_prod_combustoleo_copy = df_prod_comb.copy()


        # Crear Tabla 4 (Simulación 12 meses)
        import calendar
        from datetime import datetime
        now = datetime.now()
        current_year = now.year
        days_passed = now.timetuple().tm_yday

        meses_nombres = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        meses_cortos = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
        dias_por_mes = [calendar.monthrange(current_year, m)[1] for m in range(1, 13)]

        # 1. Simulación Crudo
        prod_dict = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try:
                p = float(val_prod)
            except:
                p = 0.0
            
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict[meses_nombres[i]] += p
                    break

        sim_data = []
        suma_total = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict[mes]
            total = prod * dias
            suma_total += total
            sim_data.append([mes, int(prod), dias, int(total)])
            
        promedio = suma_total / days_passed if days_passed > 0 else 0
        sim_data.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total)} | Prom: {promedio:.2f}"])
        df_sim = pd.DataFrame(sim_data, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 2. Simulación Gasolinas
        prod_dict_gas = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_gas.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try:
                p = float(val_prod)
            except:
                p = 0.0
            
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_gas[meses_nombres[i]] += p
                    break

        sim_data_gas = []
        suma_total_gas = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_gas[mes]
            total = prod * dias
            suma_total_gas += total
            sim_data_gas.append([mes, int(prod), dias, int(total)])
            
        promedio_gas = suma_total_gas / days_passed if days_passed > 0 else 0
        sim_data_gas.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_gas)} | Prom: {promedio_gas:.2f}"])
        df_sim_gasolinas = pd.DataFrame(sim_data_gas, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 3. Simulación Diesel
        prod_dict_die = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_die.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try:
                p = float(val_prod)
            except:
                p = 0.0
            
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_die[meses_nombres[i]] += p
                    break

        sim_data_die = []
        suma_total_die = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_die[mes]
            total = prod * dias
            suma_total_die += total
            sim_data_die.append([mes, int(prod), dias, int(total)])
            
        promedio_die = suma_total_die / days_passed if days_passed > 0 else 0
        sim_data_die.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_die)} | Prom: {promedio_die:.2f}"])
        df_sim_diesel = pd.DataFrame(sim_data_die, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 4. Simulación Turbosina
        prod_dict_turb = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_turb.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try:
                p = float(val_prod)
            except:
                p = 0.0
            
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_turb[meses_nombres[i]] += p
                    break

        sim_data_turb = []
        suma_total_turb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_turb[mes]
            total = prod * dias
            suma_total_turb += total
            sim_data_turb.append([mes, int(prod), dias, int(total)])
            
        promedio_turb = suma_total_turb / days_passed if days_passed > 0 else 0
        sim_data_turb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_turb)} | Prom: {promedio_turb:.2f}"])
        df_sim_turbosina = pd.DataFrame(sim_data_turb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 5. Simulación Asfalto
        prod_dict_asf = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_asf.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try:
                p = float(val_prod)
            except:
                p = 0.0

            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_asf[meses_nombres[i]] += p
                    break

        sim_data_asf = []
        suma_total_asf = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_asf[mes]
            total = prod * dias
            suma_total_asf += total
            sim_data_asf.append([mes, int(prod), dias, int(total)])

        promedio_asf = suma_total_asf / days_passed if days_passed > 0 else 0
        sim_data_asf.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_asf)} | Prom: {promedio_asf:.2f}"])
        df_sim_asfalto = pd.DataFrame(sim_data_asf, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 6. Simulación Combustoleo
        prod_dict_comb = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_comb.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try:
                p = float(val_prod)
            except:
                p = 0.0

            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_comb[meses_nombres[i]] += p
                    break

        sim_data_comb = []
        suma_total_comb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_comb[mes]
            total = prod * dias
            suma_total_comb += total
            sim_data_comb.append([mes, int(prod), dias, int(total)])

        promedio_comb = suma_total_comb / days_passed if days_passed > 0 else 0
        sim_data_comb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_comb)} | Prom: {promedio_comb:.2f}"])
        df_sim_combustoleo = pd.DataFrame(sim_data_comb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])
 
        # 7. Simulación Cadereyta Gasolinas
        prod_dict_cad_gas = {m: 0.0 for m in meses_nombres}
        if df_prod_cad_gas is not None:
            for idx, row_data in df_prod_cad_gas.iterrows():
                val_anio = str(row_data.iloc[0]).strip().lower()
                val_prod = row_data.iloc[1]
                try:
                    p = float(val_prod)
                except:
                    p = 0.0
                
                for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                    if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                        prod_dict_cad_gas[meses_nombres[i]] += p
                        break
 
        sim_data_cad_gas = []
        suma_total_cad_gas = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_cad_gas[mes]
            total = prod * dias
            suma_total_cad_gas += total
            sim_data_cad_gas.append([mes, int(prod), dias, int(total)])
            
        promedio_cad_gas = suma_total_cad_gas / days_passed if days_passed > 0 else 0
        sim_data_cad_gas.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_cad_gas)} | Prom: {promedio_cad_gas:.2f}"])
        df_sim_cad_gas = pd.DataFrame(sim_data_cad_gas, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])
 
        # 8. Simulación Cadereyta Diesel
        prod_dict_cad_die = {m: 0.0 for m in meses_nombres}
        if df_prod_cad_die is not None:
            for idx, row_data in df_prod_cad_die.iterrows():
                val_anio = str(row_data.iloc[0]).strip().lower()
                val_prod = row_data.iloc[1]
                try:
                    p = float(val_prod)
                except:
                    p = 0.0
                
                for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                    if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                        prod_dict_cad_die[meses_nombres[i]] += p
                        break
 
        sim_data_cad_die = []
        suma_total_cad_die = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_cad_die[mes]
            total = prod * dias
            suma_total_cad_die += total
            sim_data_cad_die.append([mes, int(prod), dias, int(total)])
            
        promedio_cad_die = suma_total_cad_die / days_passed if days_passed > 0 else 0
        sim_data_cad_die.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_cad_die)} | Prom: {promedio_cad_die:.2f}"])
        df_sim_cad_die = pd.DataFrame(sim_data_cad_die, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        prod_dict_cad_comb = {m: 0.0 for m in meses_nombres}
        if df_prod_cad_comb is not None:
            for idx, row_data in df_prod_cad_comb.iterrows():
                val_anio = str(row_data.iloc[0]).strip().lower()
                try:
                    p = float(row_data.iloc[-1])
                except:
                    p = 0.0
                    
                for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                    if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                        if not pd.isna(p) and str(p).strip() != "":
                            prod_dict_cad_comb[meses_nombres[i]] += p
                        break

        sim_data_cad_comb = []
        suma_total_cad_comb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_cad_comb[mes]
            total = prod * dias
            suma_total_cad_comb += total
            sim_data_cad_comb.append([mes, int(prod), dias, int(total)])
        
        promedio_cad_comb = suma_total_cad_comb / days_passed if days_passed > 0 else 0
        sim_data_cad_comb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_cad_comb)} | Prom: {promedio_cad_comb:.2f}"])
        df_sim_cad_comb = pd.DataFrame(sim_data_cad_comb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])
        # Simulación Crudo Cadereyta
        prod_dict_cad = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_cad.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_cad[meses_nombres[i]] += p
                    break
        sim_data_cad = []
        suma_total_cad = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_cad[mes]
            total = prod * dias
            suma_total_cad += total
            sim_data_cad.append([mes, int(prod), dias, int(total)])
        promedio_cad = suma_total_cad / days_passed if days_passed > 0 else 0
        sim_data_cad.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_cad)} | Prom: {promedio_cad:.2f}"])
        df_sim_cad = pd.DataFrame(sim_data_cad, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])
 
        # 9. Simulación Madero Crudo
        prod_dict_mad_crud = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_mad_crud.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_mad_crud[meses_nombres[i]] += p
                    break
        sim_data_mad_crud = []
        suma_total_mad_crud = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_mad_crud[mes]
            total = prod * dias
            suma_total_mad_crud += total
            sim_data_mad_crud.append([mes, int(prod), dias, int(total)])
        promedio_mad_crud = suma_total_mad_crud / days_passed if days_passed > 0 else 0
        sim_data_mad_crud.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_mad_crud)} | Prom: {promedio_mad_crud:.2f}"])
        df_sim_mad_crud = pd.DataFrame(sim_data_mad_crud, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 10. Simulación Madero Gasolinas
        prod_dict_mad_gas = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_mad_gas.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_mad_gas[meses_nombres[i]] += p
                    break
        sim_data_mad_gas = []
        suma_total_mad_gas = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_mad_gas[mes]
            total = prod * dias
            suma_total_mad_gas += total
            sim_data_mad_gas.append([mes, int(prod), dias, int(total)])
        promedio_mad_gas = suma_total_mad_gas / days_passed if days_passed > 0 else 0
        sim_data_mad_gas.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_mad_gas)} | Prom: {promedio_mad_gas:.2f}"])
        df_sim_mad_gas = pd.DataFrame(sim_data_mad_gas, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 11. Simulación Madero Diesel
        prod_dict_mad_die = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_mad_die.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_mad_die[meses_nombres[i]] += p
                    break
        sim_data_mad_die = []
        suma_total_mad_die = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_mad_die[mes]
            total = prod * dias
            suma_total_mad_die += total
            sim_data_mad_die.append([mes, int(prod), dias, int(total)])
        promedio_mad_die = suma_total_mad_die / days_passed if days_passed > 0 else 0
        sim_data_mad_die.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_mad_die)} | Prom: {promedio_mad_die:.2f}"])
        df_sim_mad_die = pd.DataFrame(sim_data_mad_die, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 12. Simulación Madero Turbosina
        prod_dict_mad_turb = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_mad_turb.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_mad_turb[meses_nombres[i]] += p
                    break
        sim_data_mad_turb = []
        suma_total_mad_turb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_mad_turb[mes]
            total = prod * dias
            suma_total_mad_turb += total
            sim_data_mad_turb.append([mes, int(prod), dias, int(total)])
        promedio_mad_turb = suma_total_mad_turb / days_passed if days_passed > 0 else 0
        sim_data_mad_turb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_mad_turb)} | Prom: {promedio_mad_turb:.2f}"])
        df_sim_mad_turb = pd.DataFrame(sim_data_mad_turb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 13. Simulación Madero Combustoleo
        prod_dict_mad_comb = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_mad_comb.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_mad_comb[meses_nombres[i]] += p
                    break
        sim_data_mad_comb = []
        suma_total_mad_comb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_mad_comb[mes]
            total = prod * dias
            suma_total_mad_comb += total
            sim_data_mad_comb.append([mes, int(prod), dias, int(total)])
        promedio_mad_comb = suma_total_mad_comb / days_passed if days_passed > 0 else 0
        sim_data_mad_comb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_mad_comb)} | Prom: {promedio_mad_comb:.2f}"])
        df_sim_mad_comb = pd.DataFrame(sim_data_mad_comb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 14. Simulación Minatitlan Crudo
        prod_dict_mina_crud = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_mina_crud.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_mina_crud[meses_nombres[i]] += p
                    break
        sim_data_mina_crud = []
        suma_total_mina_crud = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_mina_crud[mes]
            total = prod * dias
            suma_total_mina_crud += total
            sim_data_mina_crud.append([mes, int(prod), dias, int(total)])
        promedio_mina_crud = suma_total_mina_crud / days_passed if days_passed > 0 else 0
        sim_data_mina_crud.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_mina_crud)} | Prom: {promedio_mina_crud:.2f}"])
        df_sim_mina_crud = pd.DataFrame(sim_data_mina_crud, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 15. Simulación Minatitlan Gasolinas
        prod_dict_mina_gas = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_mina_gas.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_mina_gas[meses_nombres[i]] += p
                    break
        sim_data_mina_gas = []
        suma_total_mina_gas = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_mina_gas[mes]
            total = prod * dias
            suma_total_mina_gas += total
            sim_data_mina_gas.append([mes, int(prod), dias, int(total)])
        promedio_mina_gas = suma_total_mina_gas / days_passed if days_passed > 0 else 0
        sim_data_mina_gas.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_mina_gas)} | Prom: {promedio_mina_gas:.2f}"])
        df_sim_mina_gas = pd.DataFrame(sim_data_mina_gas, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 16. Simulación Minatitlan Diesel
        prod_dict_mina_die = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_mina_die.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_mina_die[meses_nombres[i]] += p
                    break
        sim_data_mina_die = []
        suma_total_mina_die = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_mina_die[mes]
            total = prod * dias
            suma_total_mina_die += total
            sim_data_mina_die.append([mes, int(prod), dias, int(total)])
        promedio_mina_die = suma_total_mina_die / days_passed if days_passed > 0 else 0
        sim_data_mina_die.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_mina_die)} | Prom: {promedio_mina_die:.2f}"])
        df_sim_mina_die = pd.DataFrame(sim_data_mina_die, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 17. Simulación Minatitlan Combustoleo
        prod_dict_mina_comb = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_mina_comb.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_mina_comb[meses_nombres[i]] += p
                    break
        sim_data_mina_comb = []
        suma_total_mina_comb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_mina_comb[mes]
            total = prod * dias
            suma_total_mina_comb += total
            sim_data_mina_comb.append([mes, int(prod), dias, int(total)])
        promedio_mina_comb = suma_total_mina_comb / days_passed if days_passed > 0 else 0
        sim_data_mina_comb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_mina_comb)} | Prom: {promedio_mina_comb:.2f}"])
        df_sim_mina_comb = pd.DataFrame(sim_data_mina_comb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        app.after(0, app.update_progress, 0.97, "Finalizando...")

        # Pasar datos a la interfaz (main thread)
        app.after(0, app.on_load_success, file_path, df_data, df_snr_copy, df_prod_copy, df_sim,
                   df_data_gasolinas, df_snr_gas_copy, df_prod_gasolinas_copy, df_sim_gasolinas,
                   df_data_diesel, df_snr_die_copy, df_prod_diesel_copy, df_sim_diesel,
                   df_data_turbosina, df_snr_turb_copy, df_prod_turbosina_copy, df_sim_turbosina,
                   df_data_asfalto, df_snr_asf_copy, df_prod_asfalto_copy, df_sim_asfalto,
                   df_data_combustoleo, df_snr_comb_copy, df_prod_combustoleo_copy, df_sim_combustoleo,
                   df_data_cad_gas, df_snr_cad_gas_copy, df_prod_cad_gas_copy, df_sim_cad_gas,
                   df_data_cad_die, df_snr_cad_die_copy, df_prod_cad_die_copy, df_sim_cad_die,
                   df_data_cad, df_snr_cad_copy, df_prod_cad_copy, df_sim_cad,
                   df_data_cad_comb, df_snr_cad_comb_copy, df_prod_cad_comb_copy, df_sim_cad_comb,
                   df_data_mad_crud, df_snr_mad_crud_copy, df_prod_mad_crud_copy, df_sim_mad_crud,
                   df_data_mad_gas, df_snr_mad_gas_copy, df_prod_mad_gas_copy, df_sim_mad_gas,
                   df_data_mad_die, df_snr_mad_die_copy, df_prod_mad_die_copy, df_sim_mad_die,
                   df_data_mad_turb, df_snr_mad_turb_copy, df_prod_mad_turb_copy, df_sim_mad_turb,
                   df_data_mad_comb, df_snr_mad_comb_copy, df_prod_mad_comb_copy, df_sim_mad_comb,
                   df_data_mina_crud, df_snr_mina_crud_copy, df_prod_mina_crud_copy, df_sim_mina_crud,
                   df_data_mina_gas, df_snr_mina_gas_copy, df_prod_mina_gas_copy, df_sim_mina_gas,
                   df_data_mina_die, df_snr_mina_die_copy, df_prod_mina_die_copy, df_sim_mina_die,
                   df_data_mina_comb, df_snr_mina_comb_copy, df_prod_mina_comb_copy, df_sim_mina_comb)

    except Exception as e:
        err_details = traceback.format_exc()
        app.after(0, app.on_load_error, str(e), err_details)
