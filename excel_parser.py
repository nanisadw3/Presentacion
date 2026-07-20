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

        # Autodetectar y usar caché si es el mismo archivo para acelerar recargas de celdas
        use_cache = (
            hasattr(app, 'cached_file_path') 
            and app.cached_file_path == file_path 
            and app.cached_df_sheet is not None 
            and app.cached_sheet_name is not None
        )
        
        if use_cache:
            sheet_to_use = app.cached_sheet_name
            df_sheet = app.cached_df_sheet
            app.cmp_value = getattr(app, 'cached_cmp_value', "1234.8")
            app.after(0, app.update_progress, 0.1, "Usando caché en memoria para recarga rápida...")
        else:
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
            app.after(0, app.update_progress, 0.1, "Leyendo hoja de cálculo...")
            df_sheet = pd.read_excel(file_path, sheet_name=sheet_to_use, header=None)
            
            try:
                df_envio = pd.read_excel(file_path, sheet_name="EnvioProDiairo", header=None)
                app.cmp_value = str(df_envio.iloc[18, 32])
                
                def get_float(row, col):
                    try:
                        return float(df_envio.iloc[row, col])
                    except:
                        return 0.0

                app.cmp_gasolinas = str(round(get_float(28, 32) + get_float(38, 32) + get_float(48, 32), 1))
                app.cmp_turbosina = str(round(get_float(68, 32), 1))
                app.cmp_diesel = str(round(get_float(78, 32), 1))

            except Exception as e:
                print("Error leyendo EnvioProDiairo:", e)
                app.cmp_value = "1234.8"
                app.cmp_gasolinas = "513"
                app.cmp_turbosina = "312.4"
                app.cmp_diesel = "386.9"
            
            # Guardar en caché en la instancia app
            app.cached_file_path = file_path
            app.cached_sheet_name = sheet_to_use
            app.cached_df_sheet = df_sheet
            app.cached_cmp_value = app.cmp_value

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

        def col_to_num(col_str):
            col_str = col_str.upper().strip()
            num = 0
            for c in col_str:
                if 'A' <= c <= 'Z':
                    num = num * 26 + (ord(c) - ord('A') + 1)
            return num - 1

        def parse_cols(cols_str):
            if not cols_str: return []
            cols_str = str(cols_str).replace(" ", "")
            if "-" in cols_str:
                parts = cols_str.split("-")
                start = col_to_num(parts[0])
                end = col_to_num(parts[1])
                return list(range(start, end + 1))
            elif "," in cols_str:
                parts = cols_str.split(",")
                return [col_to_num(p) for p in parts]
            else:
                return [col_to_num(cols_str)]

        def parse_rows(rows_str):
            if not rows_str: return (0, 0)
            parts = str(rows_str).replace(" ", "").split("-")
            start = int(parts[0]) - 1
            end = int(parts[1])
            return start, end

        def get_coords(proceso_name, d_rows_d, d_cols_d, d_rows_s, d_cols_s, d_rows_h, d_cols_h):
            import db_helper
            override = db_helper.get_coordenadas_override(proceso_name)
            if override:
                try:
                    rows_d = parse_rows(override["diaria_filas"]) if override["diaria_filas"] else d_rows_d
                    cols_d = parse_cols(override["diaria_cols"]) if override["diaria_cols"] else d_cols_d
                    if len(cols_d) == 1 and 0 not in cols_d:
                        cols_d = [0] + cols_d
                    
                    rows_s = parse_rows(override["programa_filas"]) if override["programa_filas"] else d_rows_s
                    cols_s = parse_cols(override["programa_cols"]) if override["programa_cols"] else d_cols_s
                    if len(cols_s) == 1:
                        cols_s = cols_s * 2
                    
                    rows_h = parse_rows(override["historica_filas"]) if override["historica_filas"] else d_rows_h
                    cols_h = parse_cols(override["historica_cols"]) if override["historica_cols"] else d_cols_h
                    if len(cols_h) == 1:
                        cols_h = [max(0, cols_h[0] - 1)] + cols_h
                    
                    return rows_d, cols_d, rows_s, cols_s, rows_h, cols_h
                except Exception as e:
                    print(f"Error parsing coordinates overrides for {proceso_name}: {e}")
            return d_rows_d, d_cols_d, d_rows_s, d_cols_s, d_rows_h, d_cols_h

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
                
                # Proteger contra registros corruptos en la base de datos (donde el año no es de 4 dígitos)
                if not (anio.isdigit() and len(anio) == 4):
                    continue
                
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Cadereyta -Crudo", (20, 51), [0, 2], (73, 104), [60, 60], (20, 40), list(range(48, 50)))
        # Columnas A (0) y C (2), Filas 21-51
        df_data_cad = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_cad.columns = ["Crudo", "Cadereyta"]
        df_data_cad = df_data_cad.dropna(how='all')
        df_data_cad = remove_decimals(df_data_cad)
        df_data_cad = filter_zero_rows(df_data_cad)

        # Programa: Columna BI (60) repetida, Filas 74-104
        df_snr_cad = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_cad.columns = ["CMP", "PODIM"]
        df_snr_cad = df_snr_cad.dropna(how='all').dropna(axis=1, how='all')
        df_snr_cad = remove_decimals(df_snr_cad, skip_first=True)
        df_snr_cad_copy = df_snr_cad.copy()
        
        # Fechas: Cols AW-AX (48:50), Filas 21-40
        df_prod_cad = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_cad.columns = ["Año/Mes", "Produccion"]
        df_prod_cad = df_prod_cad.dropna(how='all')
        df_prod_cad = remove_decimals(df_prod_cad, skip_last=True)
        df_prod_cad = merge_extra_prod("Cadereyta -Crudo", df_prod_cad)
        df_prod_cad_copy = df_prod_cad.copy()
        
        # --- 1.1 PROCESAR GASOLINAS (Cadereyta) ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Cadereyta -Gasolinas", (21, 51), [11, 12], (73, 104), [70, 70], (20, 51), list(range(66, 68)))
        app.after(0, app.update_progress, 0.25, "Procesando Gasolinas Cadereyta...")
        headers_cad_gas = ["Cadereyta Gas - Día", "Cadereyta Gas - Producción"]
        # Leer Tabla 1 (Rows 22-51 -> index 21:51), Cols L y M (11, 12)
        df_gas_cad = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_gas_cad.columns = headers_cad_gas
        df_gas_cad = df_gas_cad.dropna(how='all')
        df_gas_cad = remove_decimals(df_gas_cad)
        df_gas_cad = filter_zero_rows(df_gas_cad)
        df_data_cad_gas = df_gas_cad.copy()

        # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Col BS repetida (70, 70)
        df_snr_cad_gas = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_cad_gas = df_snr_cad_gas.dropna(how='all').dropna(axis=1, how='all')
        df_snr_cad_gas = remove_decimals(df_snr_cad_gas, skip_first=True)
        df_snr_cad_gas_copy = df_snr_cad_gas.copy()
        # Aplicar recorte de días reales basado en Crudo
        df_data_cad_gas = df_data_cad_gas.iloc[:num_dias_reales]

        # Leer Tabla 3 (Fecha y Producción Cadereyta Gas, Rows 21-51 -> index 20:51), Cols BO:BP (66:68)
        df_prod_cad_gas_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Cadereyta -Diesel", (73, 104), [0, 1], (73, 104), [80, 80], (20, 40), list(range(84, 86)))
        app.after(0, app.update_progress, 0.3, "Procesando Diesel Cadereyta...")
        headers_cad_die = ["Cadereyta Die - Día", "Cadereyta Die - Producción"]
        # Leer Tabla 1 (Rows 74-104 -> index 73:104), Cols A y B (0, 1)
        df_die_cad = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_die_cad.columns = headers_cad_die
        df_die_cad = df_die_cad.dropna(how='all')
        df_die_cad = remove_decimals(df_die_cad)
        df_die_cad = filter_zero_rows(df_die_cad)
        df_data_cad_die = df_die_cad.copy()

        # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Col CC repetida (80, 80)
        df_snr_cad_die = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_cad_die = df_snr_cad_die.dropna(how='all').dropna(axis=1, how='all')
        df_snr_cad_die = remove_decimals(df_snr_cad_die, skip_first=True)
        df_snr_cad_die_copy = df_snr_cad_die.copy()
        # Aplicar recorte de días reales basado en Crudo
        df_data_cad_die = df_data_cad_die.iloc[:num_dias_reales]

        # Leer Tabla 3 (Fecha y Producción Cadereyta Die, Rows 21-40 -> index 20:40), Cols CG:CH (84:86)
        df_prod_cad_die_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Cadereyta -Combustoleo", (158, 189), [18, 19], (73, 104), [98, 98], (157, 179), list(range(30, 32)))
        app.after(0, app.update_progress, 0.35, "Procesando Combustoleo Cadereyta...")
        headers_cad_comb = ["Cadereyta Comb - Día", "Cadereyta Comb - Producción"]
        df_comb_cad = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_comb_cad.columns = headers_cad_comb
        df_comb_cad = df_comb_cad.dropna(how='all')
        df_comb_cad = remove_decimals(df_comb_cad)
        df_comb_cad = filter_zero_rows(df_comb_cad)
        df_data_cad_comb = df_comb_cad.copy()

        df_snr_cad_comb = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_cad_comb = df_snr_cad_comb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_cad_comb = remove_decimals(df_snr_cad_comb, skip_first=True)
        df_snr_cad_comb_copy = df_snr_cad_comb.copy()
        df_data_cad_comb = df_data_cad_comb.iloc[:num_dias_reales]

        df_prod_cad_comb_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Madero -Crudo", (20, 51), [0, 3], (73, 104), [61, 61], (20, 40), list(range(50, 52)))
        app.after(0, app.update_progress, 0.36, "Procesando Crudo Madero...")
        df_data_mad_crud = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_mad_crud.columns = ["Crudo Día", "Madero Crudo"]
        df_data_mad_crud = df_data_mad_crud.dropna(how='all')
        df_data_mad_crud = remove_decimals(df_data_mad_crud)
        df_data_mad_crud = df_data_mad_crud.iloc[:num_dias_reales]

        df_snr_mad_crud = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_mad_crud.columns = ["CMP", "PODIM"]
        df_snr_mad_crud = df_snr_mad_crud.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mad_crud = remove_decimals(df_snr_mad_crud, skip_first=True)
        df_snr_mad_crud_copy = df_snr_mad_crud.copy()

        df_prod_mad_crud_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Madero -Gasolinas", (20, 51), [11, 13], (73, 104), [71, 71], (20, 40), list(range(70, 72)))
        app.after(0, app.update_progress, 0.37, "Procesando Gasolinas Madero...")
        df_data_mad_gas = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_mad_gas.columns = ["Gas Día", "Madero Gas"]
        df_data_mad_gas = df_data_mad_gas.dropna(how='all')
        df_data_mad_gas = remove_decimals(df_data_mad_gas)
        df_data_mad_gas = df_data_mad_gas.iloc[:num_dias_reales]

        df_snr_mad_gas = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_mad_gas.columns = ["CMP", "PODIM"]
        df_snr_mad_gas = df_snr_mad_gas.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mad_gas = remove_decimals(df_snr_mad_gas, skip_first=True)
        df_snr_mad_gas_copy = df_snr_mad_gas.copy()

        df_prod_mad_gas_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Madero -Diesel", (73, 104), [0, 2], (73, 104), [81, 81], (20, 40), list(range(86, 88)))
        app.after(0, app.update_progress, 0.38, "Procesando Diesel Madero...")
        df_data_mad_die = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_mad_die.columns = ["Diesel Día", "Madero Die"]
        df_data_mad_die = df_data_mad_die.dropna(how='all')
        df_data_mad_die = remove_decimals(df_data_mad_die)
        df_data_mad_die = df_data_mad_die.iloc[:num_dias_reales]

        df_snr_mad_die = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_mad_die.columns = ["CMP", "PODIM"]
        df_snr_mad_die = df_snr_mad_die.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mad_die = remove_decimals(df_snr_mad_die, skip_first=True)
        df_snr_mad_die_copy = df_snr_mad_die.copy()

        df_prod_mad_die_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Madero -Turbosina", (73, 104), [11, 12], (73, 104), [90, 90], (20, 40), list(range(101, 103)))
        app.after(0, app.update_progress, 0.39, "Procesando Turbosina Madero...")
        df_data_mad_turb = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_mad_turb.columns = ["Turb Día", "Madero Turb"]
        df_data_mad_turb = df_data_mad_turb.dropna(how='all')
        df_data_mad_turb = remove_decimals(df_data_mad_turb)
        df_data_mad_turb = df_data_mad_turb.iloc[:num_dias_reales]

        df_snr_mad_turb = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_mad_turb.columns = ["CMP", "PODIM"]
        df_snr_mad_turb = df_snr_mad_turb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mad_turb = remove_decimals(df_snr_mad_turb, skip_first=True)
        df_snr_mad_turb_copy = df_snr_mad_turb.copy()

        df_prod_mad_turb_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Madero -Combustoleo", (158, 189), [18, 20], (73, 104), [99, 99], (157, 179), list(range(36, 38)))
        app.after(0, app.update_progress, 0.40, "Procesando Combustoleo Madero...")
        df_data_mad_comb = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_mad_comb.columns = ["Comb Día", "Madero Comb"]
        df_data_mad_comb = df_data_mad_comb.dropna(how='all')
        df_data_mad_comb = remove_decimals(df_data_mad_comb)
        df_data_mad_comb = df_data_mad_comb.iloc[:num_dias_reales]

        df_snr_mad_comb = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_mad_comb.columns = ["CMP", "PODIM"]
        df_snr_mad_comb = df_snr_mad_comb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mad_comb = remove_decimals(df_snr_mad_comb, skip_first=True)
        df_snr_mad_comb_copy = df_snr_mad_comb.copy()

        df_prod_mad_comb_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Minatitlan -Crudo", (20, 51), [0, 4], (73, 104), [62, 62], (20, 40), list(range(54, 56)))
        app.after(0, app.update_progress, 0.41, "Procesando Crudo Minatitlan...")
        df_data_mina_crud = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_mina_crud.columns = ["Crudo Día", "Minatitlan Crudo"]
        df_data_mina_crud = df_data_mina_crud.dropna(how='all')
        df_data_mina_crud = remove_decimals(df_data_mina_crud)
        df_data_mina_crud = df_data_mina_crud.iloc[:num_dias_reales]

        df_snr_mina_crud = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_mina_crud.columns = ["CMP", "PODIM"]
        df_snr_mina_crud = df_snr_mina_crud.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mina_crud = remove_decimals(df_snr_mina_crud, skip_first=True)
        df_snr_mina_crud_copy = df_snr_mina_crud.copy()

        df_prod_mina_crud_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Minatitlan -Gasolinas", (20, 51), [11, 14], (73, 104), [72, 72], (20, 40), list(range(72, 74)))
        app.after(0, app.update_progress, 0.42, "Procesando Gasolinas Minatitlan...")
        df_data_mina_gas = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_mina_gas.columns = ["Gas Día", "Minatitlan Gas"]
        df_data_mina_gas = df_data_mina_gas.dropna(how='all')
        df_data_mina_gas = remove_decimals(df_data_mina_gas)
        df_data_mina_gas = df_data_mina_gas.iloc[:num_dias_reales]

        df_snr_mina_gas = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_mina_gas.columns = ["CMP", "PODIM"]
        df_snr_mina_gas = df_snr_mina_gas.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mina_gas = remove_decimals(df_snr_mina_gas, skip_first=True)
        df_snr_mina_gas_copy = df_snr_mina_gas.copy()

        df_prod_mina_gas_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Minatitlan -Diesel", (73, 104), [0, 3], (73, 104), [82, 82], (20, 40), list(range(90, 92)))
        app.after(0, app.update_progress, 0.43, "Procesando Diesel Minatitlan...")
        df_data_mina_die = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_mina_die.columns = ["Diesel Día", "Minatitlan Die"]
        df_data_mina_die = df_data_mina_die.dropna(how='all')
        df_data_mina_die = remove_decimals(df_data_mina_die)
        df_data_mina_die = df_data_mina_die.iloc[:num_dias_reales]

        df_snr_mina_die = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_mina_die.columns = ["CMP", "PODIM"]
        df_snr_mina_die = df_snr_mina_die.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mina_die = remove_decimals(df_snr_mina_die, skip_first=True)
        df_snr_mina_die_copy = df_snr_mina_die.copy()

        df_prod_mina_die_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Minatitlan -Combustoleo", (158, 189), [18, 21], (73, 104), [100, 100], (157, 179), list(range(39, 41)))
        app.after(0, app.update_progress, 0.44, "Procesando Combustoleo Minatitlan...")
        df_data_mina_comb = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_mina_comb.columns = ["Comb Día", "Minatitlan Comb"]
        df_data_mina_comb = df_data_mina_comb.dropna(how='all')
        df_data_mina_comb = remove_decimals(df_data_mina_comb)
        df_data_mina_comb = df_data_mina_comb.iloc[:num_dias_reales]

        df_snr_mina_comb = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_mina_comb.columns = ["CMP", "PODIM"]
        df_snr_mina_comb = df_snr_mina_comb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_mina_comb = remove_decimals(df_snr_mina_comb, skip_first=True)
        df_snr_mina_comb_copy = df_snr_mina_comb.copy()

        df_prod_mina_comb_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
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

        # --- 1.14 PROCESAR CRUDO SALAMANCA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Salamanca -Crudo", (20, 51), [0, 5], (73, 104), [63, 63], (20, 40), list(range(56, 58)))
        app.after(0, app.update_progress, 0.45, "Procesando Crudo Salamanca...")
        df_data_sala_crud = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_sala_crud.columns = ["Crudo Día", "Salamanca Crudo"]
        df_data_sala_crud = df_data_sala_crud.dropna(how='all')
        df_data_sala_crud = remove_decimals(df_data_sala_crud)
        df_data_sala_crud = df_data_sala_crud.iloc[:num_dias_reales]

        df_snr_sala_crud = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_sala_crud.columns = ["CMP", "PODIM"]
        df_snr_sala_crud = df_snr_sala_crud.dropna(how='all').dropna(axis=1, how='all')
        df_snr_sala_crud = remove_decimals(df_snr_sala_crud, skip_first=True)
        df_snr_sala_crud_copy = df_snr_sala_crud.copy()

        df_prod_sala_crud_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_sala_crud_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_sala_crud_raw = df_prod_sala_crud_raw.dropna(how='all')

        dic_idx_sala_crud = -1
        for idx, row in df_prod_sala_crud_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_sala_crud = idx - 20
                break

        if dic_idx_sala_crud != -1:
            df_prod_sala_crud = df_prod_sala_crud_raw.iloc[:dic_idx_sala_crud + 1]
        else:
            df_prod_sala_crud = df_prod_sala_crud_raw.iloc[:20]

        df_prod_sala_crud = df_prod_sala_crud.dropna(axis=1, how='all')
        df_prod_sala_crud = remove_decimals(df_prod_sala_crud, skip_last=True)
        df_prod_sala_crud = merge_extra_prod("Salamanca -Crudo", df_prod_sala_crud)
        df_prod_sala_crud_copy = df_prod_sala_crud.copy()

        # --- 1.15 PROCESAR GASOLINAS SALAMANCA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Salamanca -Gasolinas", (20, 51), [11, 15], (73, 104), [73, 73], (20, 40), list(range(74, 76)))
        app.after(0, app.update_progress, 0.46, "Procesando Gasolinas Salamanca...")
        df_data_sala_gas = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_sala_gas.columns = ["Gas Día", "Salamanca Gas"]
        df_data_sala_gas = df_data_sala_gas.dropna(how='all')
        df_data_sala_gas = remove_decimals(df_data_sala_gas)
        df_data_sala_gas = df_data_sala_gas.iloc[:num_dias_reales]

        df_snr_sala_gas = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_sala_gas.columns = ["CMP", "PODIM"]
        df_snr_sala_gas = df_snr_sala_gas.dropna(how='all').dropna(axis=1, how='all')
        df_snr_sala_gas = remove_decimals(df_snr_sala_gas, skip_first=True)
        df_snr_sala_gas_copy = df_snr_sala_gas.copy()

        df_prod_sala_gas_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_sala_gas_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_sala_gas_raw = df_prod_sala_gas_raw.dropna(how='all')

        dic_idx_sala_gas = -1
        for idx, row in df_prod_sala_gas_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_sala_gas = idx - 20
                break

        if dic_idx_sala_gas != -1:
            df_prod_sala_gas = df_prod_sala_gas_raw.iloc[:dic_idx_sala_gas + 1]
        else:
            df_prod_sala_gas = df_prod_sala_gas_raw.iloc[:20]

        df_prod_sala_gas = df_prod_sala_gas.dropna(axis=1, how='all')
        df_prod_sala_gas = remove_decimals(df_prod_sala_gas, skip_last=True)
        df_prod_sala_gas = merge_extra_prod("Salamanca -Gasolinas", df_prod_sala_gas)
        df_prod_sala_gas_copy = df_prod_sala_gas.copy()

        # --- 1.16 PROCESAR DIESEL SALAMANCA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Salamanca -Diesel", (73, 104), [0, 4], (73, 104), [83, 83], (20, 40), list(range(92, 94)))
        app.after(0, app.update_progress, 0.47, "Procesando Diesel Salamanca...")
        df_data_sala_die = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_sala_die.columns = ["Diesel Día", "Salamanca Die"]
        df_data_sala_die = df_data_sala_die.dropna(how='all')
        df_data_sala_die = remove_decimals(df_data_sala_die)
        df_data_sala_die = df_data_sala_die.iloc[:num_dias_reales]

        df_snr_sala_die = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_sala_die.columns = ["CMP", "PODIM"]
        df_snr_sala_die = df_snr_sala_die.dropna(how='all').dropna(axis=1, how='all')
        df_snr_sala_die = remove_decimals(df_snr_sala_die, skip_first=True)
        df_snr_sala_die_copy = df_snr_sala_die.copy()

        df_prod_sala_die_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_sala_die_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_sala_die_raw = df_prod_sala_die_raw.dropna(how='all')

        dic_idx_sala_die = -1
        for idx, row in df_prod_sala_die_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_sala_die = idx - 20
                break

        if dic_idx_sala_die != -1:
            df_prod_sala_die = df_prod_sala_die_raw.iloc[:dic_idx_sala_die + 1]
        else:
            df_prod_sala_die = df_prod_sala_die_raw.iloc[:20]

        df_prod_sala_die = df_prod_sala_die.dropna(axis=1, how='all')
        df_prod_sala_die = remove_decimals(df_prod_sala_die, skip_last=True)
        df_prod_sala_die = merge_extra_prod("Salamanca -Diesel", df_prod_sala_die)
        df_prod_sala_die_copy = df_prod_sala_die.copy()

        # --- 1.17 PROCESAR TURBOSINA SALAMANCA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Salamanca -Turbosina", (73, 104), [11, 13], (73, 104), [91, 91], (20, 40), list(range(103, 105)))
        app.after(0, app.update_progress, 0.48, "Procesando Turbosina Salamanca...")
        df_data_sala_turb = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_sala_turb.columns = ["Turb Día", "Salamanca Turb"]
        df_data_sala_turb = df_data_sala_turb.dropna(how='all')
        df_data_sala_turb = remove_decimals(df_data_sala_turb)
        df_data_sala_turb = df_data_sala_turb.iloc[:num_dias_reales]

        df_snr_sala_turb = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_sala_turb.columns = ["CMP", "PODIM"]
        df_snr_sala_turb = df_snr_sala_turb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_sala_turb = remove_decimals(df_snr_sala_turb, skip_first=True)
        df_snr_sala_turb_copy = df_snr_sala_turb.copy()

        df_prod_sala_turb_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_sala_turb_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_sala_turb_raw = df_prod_sala_turb_raw.dropna(how='all')

        dic_idx_sala_turb = -1
        for idx, row in df_prod_sala_turb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_sala_turb = idx - 20
                break

        if dic_idx_sala_turb != -1:
            df_prod_sala_turb = df_prod_sala_turb_raw.iloc[:dic_idx_sala_turb + 1]
        else:
            df_prod_sala_turb = df_prod_sala_turb_raw.iloc[:20]

        df_prod_sala_turb = df_prod_sala_turb.dropna(axis=1, how='all')
        df_prod_sala_turb = remove_decimals(df_prod_sala_turb, skip_last=True)
        df_prod_sala_turb = merge_extra_prod("Salamanca -Turbosina", df_prod_sala_turb)
        df_prod_sala_turb_copy = df_prod_sala_turb.copy()

        # --- 1.18 PROCESAR COMBUSTOLEO SALAMANCA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Salamanca -Combustoleo", (158, 189), [18, 22], (73, 104), [102, 102], (157, 179), list(range(42, 44)))
        app.after(0, app.update_progress, 0.49, "Procesando Combustoleo Salamanca...")
        df_data_sala_comb = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_sala_comb.columns = ["Comb Día", "Salamanca Comb"]
        df_data_sala_comb = df_data_sala_comb.dropna(how='all')
        df_data_sala_comb = remove_decimals(df_data_sala_comb)
        df_data_sala_comb = df_data_sala_comb.iloc[:num_dias_reales]

        df_snr_sala_comb = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_sala_comb.columns = ["CMP", "PODIM"]
        df_snr_sala_comb = df_snr_sala_comb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_sala_comb = remove_decimals(df_snr_sala_comb, skip_first=True)
        df_snr_sala_comb_copy = df_snr_sala_comb.copy()

        df_prod_sala_comb_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_sala_comb_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_sala_comb_raw = df_prod_sala_comb_raw.dropna(how='all')

        dic_idx_sala_comb = -1
        for idx, row in df_prod_sala_comb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_sala_comb = idx - 157
                break

        if dic_idx_sala_comb != -1:
            df_prod_sala_comb = df_prod_sala_comb_raw.iloc[:dic_idx_sala_comb + 1]
        else:
            df_prod_sala_comb = df_prod_sala_comb_raw.iloc[:22]

        df_prod_sala_comb = df_prod_sala_comb.dropna(axis=1, how='all')
        df_prod_sala_comb = remove_decimals(df_prod_sala_comb, skip_last=True)
        df_prod_sala_comb = merge_extra_prod("Salamanca -Combustoleo", df_prod_sala_comb)
        df_prod_sala_comb_copy = df_prod_sala_comb.copy()

        # --- 1.19 PROCESAR CRUDO SALINA CRUZ ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Salina Cruz -Crudo", (20, 51), [0, 6], (73, 104), [64, 64], (20, 40), list(range(60, 62)))
        app.after(0, app.update_progress, 0.50, "Procesando Crudo Salina Cruz...")
        df_data_sal_crud = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_sal_crud.columns = ["Crudo Día", "Salina Cruz Crudo"]
        df_data_sal_crud = df_data_sal_crud.dropna(how='all')
        df_data_sal_crud = remove_decimals(df_data_sal_crud)
        df_data_sal_crud = df_data_sal_crud.iloc[:num_dias_reales]

        df_snr_sal_crud = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_sal_crud.columns = ["CMP", "PODIM"]
        df_snr_sal_crud = df_snr_sal_crud.dropna(how='all').dropna(axis=1, how='all')
        df_snr_sal_crud = remove_decimals(df_snr_sal_crud, skip_first=True)
        df_snr_sal_crud_copy = df_snr_sal_crud.copy()

        df_prod_sal_crud_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_sal_crud_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_sal_crud_raw = df_prod_sal_crud_raw.dropna(how='all')

        dic_idx_sal_crud = -1
        for idx, row in df_prod_sal_crud_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_sal_crud = idx - 20
                break

        if dic_idx_sal_crud != -1:
            df_prod_sal_crud = df_prod_sal_crud_raw.iloc[:dic_idx_sal_crud + 1]
        else:
            df_prod_sal_crud = df_prod_sal_crud_raw.iloc[:20]

        df_prod_sal_crud = df_prod_sal_crud.dropna(axis=1, how='all')
        df_prod_sal_crud = remove_decimals(df_prod_sal_crud, skip_last=True)
        df_prod_sal_crud = merge_extra_prod("Salina Cruz -Crudo", df_prod_sal_crud)
        df_prod_sal_crud_copy = df_prod_sal_crud.copy()

        # --- 1.20 PROCESAR GASOLINAS SALINA CRUZ ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Salina Cruz -Gasolinas", (20, 51), [11, 16], (73, 104), [74, 74], (20, 40), list(range(76, 78)))
        app.after(0, app.update_progress, 0.51, "Procesando Gasolinas Salina Cruz...")
        df_data_sal_gas = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_sal_gas.columns = ["Gas Día", "Salina Cruz Gas"]
        df_data_sal_gas = df_data_sal_gas.dropna(how='all')
        df_data_sal_gas = remove_decimals(df_data_sal_gas)
        df_data_sal_gas = df_data_sal_gas.iloc[:num_dias_reales]

        df_snr_sal_gas = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_sal_gas.columns = ["CMP", "PODIM"]
        df_snr_sal_gas = df_snr_sal_gas.dropna(how='all').dropna(axis=1, how='all')
        df_snr_sal_gas = remove_decimals(df_snr_sal_gas, skip_first=True)
        df_snr_sal_gas_copy = df_snr_sal_gas.copy()

        df_prod_sal_gas_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_sal_gas_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_sal_gas_raw = df_prod_sal_gas_raw.dropna(how='all')

        dic_idx_sal_gas = -1
        for idx, row in df_prod_sal_gas_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_sal_gas = idx - 20
                break

        if dic_idx_sal_gas != -1:
            df_prod_sal_gas = df_prod_sal_gas_raw.iloc[:dic_idx_sal_gas + 1]
        else:
            df_prod_sal_gas = df_prod_sal_gas_raw.iloc[:20]

        df_prod_sal_gas = df_prod_sal_gas.dropna(axis=1, how='all')
        df_prod_sal_gas = remove_decimals(df_prod_sal_gas, skip_last=True)
        df_prod_sal_gas = merge_extra_prod("Salina Cruz -Gasolinas", df_prod_sal_gas)
        df_prod_sal_gas_copy = df_prod_sal_gas.copy()

        # --- 1.21 PROCESAR DIESEL SALINA CRUZ ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Salina Cruz -Diesel", (73, 104), [0, 5], (73, 104), [84, 84], (20, 40), list(range(94, 96)))
        app.after(0, app.update_progress, 0.52, "Procesando Diesel Salina Cruz...")
        df_data_sal_die = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_sal_die.columns = ["Diesel Día", "Salina Cruz Die"]
        df_data_sal_die = df_data_sal_die.dropna(how='all')
        df_data_sal_die = remove_decimals(df_data_sal_die)
        df_data_sal_die = df_data_sal_die.iloc[:num_dias_reales]

        df_snr_sal_die = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_sal_die.columns = ["CMP", "PODIM"]
        df_snr_sal_die = df_snr_sal_die.dropna(how='all').dropna(axis=1, how='all')
        df_snr_sal_die = remove_decimals(df_snr_sal_die, skip_first=True)
        df_snr_sal_die_copy = df_snr_sal_die.copy()

        df_prod_sal_die_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_sal_die_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_sal_die_raw = df_prod_sal_die_raw.dropna(how='all')

        dic_idx_sal_die = -1
        for idx, row in df_prod_sal_die_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_sal_die = idx - 20
                break

        if dic_idx_sal_die != -1:
            df_prod_sal_die = df_prod_sal_die_raw.iloc[:dic_idx_sal_die + 1]
        else:
            df_prod_sal_die = df_prod_sal_die_raw.iloc[:20]

        df_prod_sal_die = df_prod_sal_die.dropna(axis=1, how='all')
        df_prod_sal_die = remove_decimals(df_prod_sal_die, skip_last=True)
        df_prod_sal_die = merge_extra_prod("Salina Cruz -Diesel", df_prod_sal_die)
        df_prod_sal_die_copy = df_prod_sal_die.copy()

        # --- 1.22 PROCESAR TURBOSINA SALINA CRUZ ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Salina Cruz -Turbosina", (73, 104), [11, 14], (73, 104), [92, 92], (20, 40), list(range(105, 107)))
        app.after(0, app.update_progress, 0.53, "Procesando Turbosina Salina Cruz...")
        df_data_sal_turb = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_sal_turb.columns = ["Turb Día", "Salina Cruz Turb"]
        df_data_sal_turb = df_data_sal_turb.dropna(how='all')
        df_data_sal_turb = remove_decimals(df_data_sal_turb)
        df_data_sal_turb = df_data_sal_turb.iloc[:num_dias_reales]

        df_snr_sal_turb = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_sal_turb.columns = ["CMP", "PODIM"]
        df_snr_sal_turb = df_snr_sal_turb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_sal_turb = remove_decimals(df_snr_sal_turb, skip_first=True)
        df_snr_sal_turb_copy = df_snr_sal_turb.copy()

        df_prod_sal_turb_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_sal_turb_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_sal_turb_raw = df_prod_sal_turb_raw.dropna(how='all')

        dic_idx_sal_turb = -1
        for idx, row in df_prod_sal_turb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_sal_turb = idx - 20
                break

        if dic_idx_sal_turb != -1:
            df_prod_sal_turb = df_prod_sal_turb_raw.iloc[:dic_idx_sal_turb + 1]
        else:
            df_prod_sal_turb = df_prod_sal_turb_raw.iloc[:20]

        df_prod_sal_turb = df_prod_sal_turb.dropna(axis=1, how='all')
        df_prod_sal_turb = remove_decimals(df_prod_sal_turb, skip_last=True)
        df_prod_sal_turb = merge_extra_prod("Salina Cruz -Turbosina", df_prod_sal_turb)
        df_prod_sal_turb_copy = df_prod_sal_turb.copy()

        # --- 1.23 PROCESAR COMBUSTOLEO SALINA CRUZ ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Salina Cruz -Combustoleo", (158, 189), [18, 23], (73, 104), [103, 103], (157, 179), list(range(45, 47)))
        app.after(0, app.update_progress, 0.54, "Procesando Combustoleo Salina Cruz...")
        df_data_sal_comb = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_sal_comb.columns = ["Comb Día", "Salina Cruz Comb"]
        df_data_sal_comb = df_data_sal_comb.dropna(how='all')
        df_data_sal_comb = remove_decimals(df_data_sal_comb)
        df_data_sal_comb = df_data_sal_comb.iloc[:num_dias_reales]

        df_snr_sal_comb = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_sal_comb.columns = ["CMP", "PODIM"]
        df_snr_sal_comb = df_snr_sal_comb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_sal_comb = remove_decimals(df_snr_sal_comb, skip_first=True)
        df_snr_sal_comb_copy = df_snr_sal_comb.copy()

        df_prod_sal_comb_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_sal_comb_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_sal_comb_raw = df_prod_sal_comb_raw.dropna(how='all')

        dic_idx_sal_comb = -1
        for idx, row in df_prod_sal_comb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_sal_comb = idx - 157
                break

        if dic_idx_sal_comb != -1:
            df_prod_sal_comb = df_prod_sal_comb_raw.iloc[:dic_idx_sal_comb + 1]
        else:
            df_prod_sal_comb = df_prod_sal_comb_raw.iloc[:22]

        df_prod_sal_comb = df_prod_sal_comb.dropna(axis=1, how='all')
        df_prod_sal_comb = remove_decimals(df_prod_sal_comb, skip_last=True)
        df_prod_sal_comb = merge_extra_prod("Salina Cruz -Combustoleo", df_prod_sal_comb)
        df_prod_sal_comb_copy = df_prod_sal_comb.copy()

        # --- 1.24 PROCESAR CRUDO TULA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Tula -Crudo", (20, 51), [0, 7], (73, 104), [65, 65], (20, 40), list(range(62, 64)))
        app.after(0, app.update_progress, 0.55, "Procesando Crudo Tula...")
        df_data_tula_crud = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_tula_crud.columns = ["Crudo Día", "Tula Crudo"]
        df_data_tula_crud = df_data_tula_crud.dropna(how='all')
        df_data_tula_crud = remove_decimals(df_data_tula_crud)
        df_data_tula_crud = df_data_tula_crud.iloc[:num_dias_reales]

        df_snr_tula_crud = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_tula_crud.columns = ["CMP", "PODIM"]
        df_snr_tula_crud = df_snr_tula_crud.dropna(how='all').dropna(axis=1, how='all')
        df_snr_tula_crud = remove_decimals(df_snr_tula_crud, skip_first=True)
        df_snr_tula_crud_copy = df_snr_tula_crud.copy()

        df_prod_tula_crud_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_tula_crud_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_tula_crud_raw = df_prod_tula_crud_raw.dropna(how='all')

        dic_idx_tula_crud = -1
        for idx, row in df_prod_tula_crud_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_tula_crud = idx - 20
                break

        if dic_idx_tula_crud != -1:
            df_prod_tula_crud = df_prod_tula_crud_raw.iloc[:dic_idx_tula_crud + 1]
        else:
            df_prod_tula_crud = df_prod_tula_crud_raw.iloc[:20]

        df_prod_tula_crud = df_prod_tula_crud.dropna(axis=1, how='all')
        df_prod_tula_crud = remove_decimals(df_prod_tula_crud, skip_last=True)
        df_prod_tula_crud = merge_extra_prod("Tula -Crudo", df_prod_tula_crud)
        df_prod_tula_crud_copy = df_prod_tula_crud.copy()

        # --- 1.25 PROCESAR GASOLINAS TULA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Tula -Gasolinas", (20, 51), [11, 17], (73, 104), [75, 75], (20, 40), list(range(80, 82)))
        app.after(0, app.update_progress, 0.56, "Procesando Gasolinas Tula...")
        df_data_tula_gas = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_tula_gas.columns = ["Gas Día", "Tula Gas"]
        df_data_tula_gas = df_data_tula_gas.dropna(how='all')
        df_data_tula_gas = remove_decimals(df_data_tula_gas)
        df_data_tula_gas = df_data_tula_gas.iloc[:num_dias_reales]

        df_snr_tula_gas = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_tula_gas.columns = ["CMP", "PODIM"]
        df_snr_tula_gas = df_snr_tula_gas.dropna(how='all').dropna(axis=1, how='all')
        df_snr_tula_gas = remove_decimals(df_snr_tula_gas, skip_first=True)
        df_snr_tula_gas_copy = df_snr_tula_gas.copy()

        df_prod_tula_gas_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_tula_gas_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_tula_gas_raw = df_prod_tula_gas_raw.dropna(how='all')

        dic_idx_tula_gas = -1
        for idx, row in df_prod_tula_gas_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_tula_gas = idx - 20
                break

        if dic_idx_tula_gas != -1:
            df_prod_tula_gas = df_prod_tula_gas_raw.iloc[:dic_idx_tula_gas + 1]
        else:
            df_prod_tula_gas = df_prod_tula_gas_raw.iloc[:20]

        df_prod_tula_gas = df_prod_tula_gas.dropna(axis=1, how='all')
        df_prod_tula_gas = remove_decimals(df_prod_tula_gas, skip_last=True)
        df_prod_tula_gas = merge_extra_prod("Tula -Gasolinas", df_prod_tula_gas)
        df_prod_tula_gas_copy = df_prod_tula_gas.copy()

        # --- 1.26 PROCESAR DIESEL TULA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Tula -Diesel", (73, 104), [0, 6], (73, 104), [85, 85], (20, 40), list(range(97, 99)))
        app.after(0, app.update_progress, 0.57, "Procesando Diesel Tula...")
        df_data_tula_die = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_tula_die.columns = ["Diesel Día", "Tula Die"]
        df_data_tula_die = df_data_tula_die.dropna(how='all')
        df_data_tula_die = remove_decimals(df_data_tula_die)
        df_data_tula_die = df_data_tula_die.iloc[:num_dias_reales]

        df_snr_tula_die = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_tula_die.columns = ["CMP", "PODIM"]
        df_snr_tula_die = df_snr_tula_die.dropna(how='all').dropna(axis=1, how='all')
        df_snr_tula_die = remove_decimals(df_snr_tula_die, skip_first=True)
        df_snr_tula_die_copy = df_snr_tula_die.copy()

        df_prod_tula_die_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_tula_die_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_tula_die_raw = df_prod_tula_die_raw.dropna(how='all')

        dic_idx_tula_die = -1
        for idx, row in df_prod_tula_die_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_tula_die = idx - 20
                break

        if dic_idx_tula_die != -1:
            df_prod_tula_die = df_prod_tula_die_raw.iloc[:dic_idx_tula_die + 1]
        else:
            df_prod_tula_die = df_prod_tula_die_raw.iloc[:20]

        df_prod_tula_die = df_prod_tula_die.dropna(axis=1, how='all')
        df_prod_tula_die = remove_decimals(df_prod_tula_die, skip_last=True)
        df_prod_tula_die = merge_extra_prod("Tula -Diesel", df_prod_tula_die)
        df_prod_tula_die_copy = df_prod_tula_die.copy()

        # --- 1.27 PROCESAR TURBOSINA TULA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Tula -Turbosina", (73, 104), [11, 15], (73, 104), [93, 93], (20, 40), list(range(107, 109)))
        app.after(0, app.update_progress, 0.58, "Procesando Turbosina Tula...")
        df_data_tula_turb = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_tula_turb.columns = ["Turb Día", "Tula Turb"]
        df_data_tula_turb = df_data_tula_turb.dropna(how='all')
        df_data_tula_turb = remove_decimals(df_data_tula_turb)
        df_data_tula_turb = df_data_tula_turb.iloc[:num_dias_reales]

        df_snr_tula_turb = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_tula_turb.columns = ["CMP", "PODIM"]
        df_snr_tula_turb = df_snr_tula_turb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_tula_turb = remove_decimals(df_snr_tula_turb, skip_first=True)
        df_snr_tula_turb_copy = df_snr_tula_turb.copy()

        df_prod_tula_turb_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_tula_turb_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_tula_turb_raw = df_prod_tula_turb_raw.dropna(how='all')

        dic_idx_tula_turb = -1
        for idx, row in df_prod_tula_turb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_tula_turb = idx - 20
                break

        if dic_idx_tula_turb != -1:
            df_prod_tula_turb = df_prod_tula_turb_raw.iloc[:dic_idx_tula_turb + 1]
        else:
            df_prod_tula_turb = df_prod_tula_turb_raw.iloc[:20]

        df_prod_tula_turb = df_prod_tula_turb.dropna(axis=1, how='all')
        df_prod_tula_turb = remove_decimals(df_prod_tula_turb, skip_last=True)
        df_prod_tula_turb = merge_extra_prod("Tula -Turbosina", df_prod_tula_turb)
        df_prod_tula_turb_copy = df_prod_tula_turb.copy()

        # --- 1.28 PROCESAR COMBUSTOLEO TULA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Tula -Combustoleo", (158, 189), [18, 24], (73, 104), [104, 104], (157, 179), list(range(48, 50)))
        app.after(0, app.update_progress, 0.59, "Procesando Combustoleo Tula...")
        df_data_tula_comb = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_tula_comb.columns = ["Comb Día", "Tula Comb"]
        df_data_tula_comb = df_data_tula_comb.dropna(how='all')
        df_data_tula_comb = remove_decimals(df_data_tula_comb)
        df_data_tula_comb = df_data_tula_comb.iloc[:num_dias_reales]

        df_snr_tula_comb = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_tula_comb.columns = ["CMP", "PODIM"]
        df_snr_tula_comb = df_snr_tula_comb.dropna(how='all').dropna(axis=1, how='all')
        df_snr_tula_comb = remove_decimals(df_snr_tula_comb, skip_first=True)
        df_snr_tula_comb_copy = df_snr_tula_comb.copy()

        df_prod_tula_comb_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_tula_comb_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_tula_comb_raw = df_prod_tula_comb_raw.dropna(how='all')

        dic_idx_tula_comb = -1
        for idx, row in df_prod_tula_comb_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_tula_comb = idx - 157
                break

        if dic_idx_tula_comb != -1:
            df_prod_tula_comb = df_prod_tula_comb_raw.iloc[:dic_idx_tula_comb + 1]
        else:
            df_prod_tula_comb = df_prod_tula_comb_raw.iloc[:22]

        df_prod_tula_comb = df_prod_tula_comb.dropna(axis=1, how='all')
        df_prod_tula_comb = remove_decimals(df_prod_tula_comb, skip_last=True)
        df_prod_tula_comb = merge_extra_prod("Tula -Combustoleo", df_prod_tula_comb)
        df_prod_tula_comb_copy = df_prod_tula_comb.copy()

        # --- 1.29 PROCESAR CRUDO OLMECA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Olmeca -Crudo", (157, 188), [2, 3], (158, 189), [13, 13], (191, 206), [2, 3])
        app.after(0, app.update_progress, 0.60, "Procesando Crudo Olmeca...")
        df_data_olme_crud = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_olme_crud.columns = ["Crudo Día", "Olmeca Crudo"]
        df_data_olme_crud = df_data_olme_crud.dropna(how='all')
        df_data_olme_crud = remove_decimals(df_data_olme_crud)
        df_data_olme_crud = df_data_olme_crud.iloc[:num_dias_reales]

        df_snr_olme_crud = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_olme_crud.columns = ["CMP", "PODIM"]
        df_snr_olme_crud = df_snr_olme_crud.dropna(how='all').dropna(axis=1, how='all')
        df_snr_olme_crud = remove_decimals(df_snr_olme_crud, skip_first=True)
        df_snr_olme_crud_copy = df_snr_olme_crud.copy()

        df_prod_olme_crud_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_olme_crud_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_olme_crud_raw = df_prod_olme_crud_raw.dropna(how='all')

        dic_idx_olme_crud = -1
        for idx, row in df_prod_olme_crud_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_olme_crud = idx - 191
                break

        if dic_idx_olme_crud != -1:
            df_prod_olme_crud = df_prod_olme_crud_raw.iloc[:dic_idx_olme_crud + 1]
        else:
            df_prod_olme_crud = df_prod_olme_crud_raw.iloc[:15]

        df_prod_olme_crud = df_prod_olme_crud.dropna(axis=1, how='all')
        df_prod_olme_crud = remove_decimals(df_prod_olme_crud, skip_last=True)
        df_prod_olme_crud = merge_extra_prod("Olmeca -Crudo", df_prod_olme_crud)
        df_prod_olme_crud_copy = df_prod_olme_crud.copy()

        # --- 1.30 PROCESAR GASOLINAS OLMECA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Olmeca -Gasolinas", (157, 188), [5, 6], (158, 189), [14, 14], (191, 206), [5, 6])
        app.after(0, app.update_progress, 0.61, "Procesando Gasolinas Olmeca...")
        df_data_olme_gas = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_olme_gas.columns = ["Gas Día", "Olmeca Gas"]
        df_data_olme_gas = df_data_olme_gas.dropna(how='all')
        df_data_olme_gas = remove_decimals(df_data_olme_gas)
        df_data_olme_gas = df_data_olme_gas.iloc[:num_dias_reales]

        df_snr_olme_gas = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_olme_gas.columns = ["CMP", "PODIM"]
        df_snr_olme_gas = df_snr_olme_gas.dropna(how='all').dropna(axis=1, how='all')
        df_snr_olme_gas = remove_decimals(df_snr_olme_gas, skip_first=True)
        df_snr_olme_gas_copy = df_snr_olme_gas.copy()

        df_prod_olme_gas_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_olme_gas_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_olme_gas_raw = df_prod_olme_gas_raw.dropna(how='all')

        dic_idx_olme_gas = -1
        for idx, row in df_prod_olme_gas_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_olme_gas = idx - 191
                break

        if dic_idx_olme_gas != -1:
            df_prod_olme_gas = df_prod_olme_gas_raw.iloc[:dic_idx_olme_gas + 1]
        else:
            df_prod_olme_gas = df_prod_olme_gas_raw.iloc[:15]

        df_prod_olme_gas = df_prod_olme_gas.dropna(axis=1, how='all')
        df_prod_olme_gas = remove_decimals(df_prod_olme_gas, skip_last=True)
        df_prod_olme_gas = merge_extra_prod("Olmeca -Gasolinas", df_prod_olme_gas)
        df_prod_olme_gas_copy = df_prod_olme_gas.copy()

        # --- 1.31 PROCESAR DIESEL OLMECA ---
        r_d, c_d, r_s, c_s, r_h, c_h = get_coords("Olmeca -Diesel", (157, 188), [8, 9], (158, 189), [15, 15], (191, 206), [8, 9])
        app.after(0, app.update_progress, 0.62, "Procesando Diesel Olmeca...")
        df_data_olme_die = df_sheet.iloc[r_d[0]:r_d[1], c_d].copy()
        df_data_olme_die.columns = ["Diesel Día", "Olmeca Die"]
        df_data_olme_die = df_data_olme_die.dropna(how='all')
        df_data_olme_die = remove_decimals(df_data_olme_die)
        df_data_olme_die = df_data_olme_die.iloc[:num_dias_reales]

        df_snr_olme_die = df_sheet.iloc[r_s[0]:r_s[1], c_s].copy()
        df_snr_olme_die.columns = ["CMP", "PODIM"]
        df_snr_olme_die = df_snr_olme_die.dropna(how='all').dropna(axis=1, how='all')
        df_snr_olme_die = remove_decimals(df_snr_olme_die, skip_first=True)
        df_snr_olme_die_copy = df_snr_olme_die.copy()

        df_prod_olme_die_raw = df_sheet.iloc[r_h[0]:r_h[1], c_h].copy()
        df_prod_olme_die_raw.columns = ["Año/Mes", "Produccion"]
        df_prod_olme_die_raw = df_prod_olme_die_raw.dropna(how='all')

        dic_idx_olme_die = -1
        for idx, row in df_prod_olme_die_raw.iterrows():
            val = str(row.iloc[0]).strip().lower()
            if "dic" in val or "diciembre" in val:
                dic_idx_olme_die = idx - 191
                break

        if dic_idx_olme_die != -1:
            df_prod_olme_die = df_prod_olme_die_raw.iloc[:dic_idx_olme_die + 1]
        else:
            df_prod_olme_die = df_prod_olme_die_raw.iloc[:15]

        df_prod_olme_die = df_prod_olme_die.dropna(axis=1, how='all')
        df_prod_olme_die = remove_decimals(df_prod_olme_die, skip_last=True)
        df_prod_olme_die = merge_extra_prod("Olmeca -Diesel", df_prod_olme_die)
        df_prod_olme_die_copy = df_prod_olme_die.copy()

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
        df_data_asfalto = df_asf.copy()

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
        
        # Determinar el año que representa la hoja de cálculo (sheet_year)
        sheet_year = 2026
        if df_prod is not None and not df_prod.empty:
            years_found = []
            for idx, row_data in df_prod.iterrows():
                val = str(row_data.iloc[0]).strip()
                if val.isdigit() and len(val) == 4:
                    years_found.append(int(val))
            if years_found:
                sheet_year = max(years_found) + 1

        system_year = now.year
        system_month = now.month
        system_day = now.day
        
        meses_nombres = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        meses_cortos = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
        
        # Días por mes dinámicos basados en la relación entre el año de la hoja y el año del sistema
        dias_por_mes = []
        for m in range(1, 13):
            if sheet_year < system_year:
                # El año de la hoja ya concluyó. Todos los meses completos.
                dias_por_mes.append(calendar.monthrange(sheet_year, m)[1])
            elif sheet_year > system_year:
                # El año de la hoja es futuro. Ningún mes transcurrido.
                dias_por_mes.append(0)
            else:
                # Es el año actual
                if m < system_month:
                    dias_por_mes.append(calendar.monthrange(sheet_year, m)[1])
                elif m == system_month:
                    dias_por_mes.append(max(0, system_day - 1))
                else:
                    dias_por_mes.append(0)

        days_passed = sum(dias_por_mes)

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

        # 18. Simulación Salamanca Crudo
        prod_dict_sala_crud = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_sala_crud.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_sala_crud[meses_nombres[i]] += p
                    break
        sim_data_sala_crud = []
        suma_total_sala_crud = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_sala_crud[mes]
            total = prod * dias
            suma_total_sala_crud += total
            sim_data_sala_crud.append([mes, int(prod), dias, int(total)])
        promedio_sala_crud = suma_total_sala_crud / days_passed if days_passed > 0 else 0
        sim_data_sala_crud.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_sala_crud)} | Prom: {promedio_sala_crud:.2f}"])
        df_sim_sala_crud = pd.DataFrame(sim_data_sala_crud, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 19. Simulación Salamanca Gasolinas
        prod_dict_sala_gas = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_sala_gas.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_sala_gas[meses_nombres[i]] += p
                    break
        sim_data_sala_gas = []
        suma_total_sala_gas = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_sala_gas[mes]
            total = prod * dias
            suma_total_sala_gas += total
            sim_data_sala_gas.append([mes, int(prod), dias, int(total)])
        promedio_sala_gas = suma_total_sala_gas / days_passed if days_passed > 0 else 0
        sim_data_sala_gas.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_sala_gas)} | Prom: {promedio_sala_gas:.2f}"])
        df_sim_sala_gas = pd.DataFrame(sim_data_sala_gas, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 20. Simulación Salamanca Diesel
        prod_dict_sala_die = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_sala_die.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_sala_die[meses_nombres[i]] += p
                    break
        sim_data_sala_die = []
        suma_total_sala_die = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_sala_die[mes]
            total = prod * dias
            suma_total_sala_die += total
            sim_data_sala_die.append([mes, int(prod), dias, int(total)])
        promedio_sala_die = suma_total_sala_die / days_passed if days_passed > 0 else 0
        sim_data_sala_die.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_sala_die)} | Prom: {promedio_sala_die:.2f}"])
        df_sim_sala_die = pd.DataFrame(sim_data_sala_die, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 21. Simulación Salamanca Turbosina
        prod_dict_sala_turb = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_sala_turb.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_sala_turb[meses_nombres[i]] += p
                    break
        sim_data_sala_turb = []
        suma_total_sala_turb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_sala_turb[mes]
            total = prod * dias
            suma_total_sala_turb += total
            sim_data_sala_turb.append([mes, int(prod), dias, int(total)])
        promedio_sala_turb = suma_total_sala_turb / days_passed if days_passed > 0 else 0
        sim_data_sala_turb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_sala_turb)} | Prom: {promedio_sala_turb:.2f}"])
        df_sim_sala_turb = pd.DataFrame(sim_data_sala_turb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 22. Simulación Salamanca Combustoleo
        prod_dict_sala_comb = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_sala_comb.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_sala_comb[meses_nombres[i]] += p
                    break
        sim_data_sala_comb = []
        suma_total_sala_comb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_sala_comb[mes]
            total = prod * dias
            suma_total_sala_comb += total
            sim_data_sala_comb.append([mes, int(prod), dias, int(total)])
        promedio_sala_comb = suma_total_sala_comb / days_passed if days_passed > 0 else 0
        sim_data_sala_comb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_sala_comb)} | Prom: {promedio_sala_comb:.2f}"])
        df_sim_sala_comb = pd.DataFrame(sim_data_sala_comb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 23. Simulación Salina Cruz Crudo
        prod_dict_sal_crud = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_sal_crud.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_sal_crud[meses_nombres[i]] += p
                    break
        sim_data_sal_crud = []
        suma_total_sal_crud = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_sal_crud[mes]
            total = prod * dias
            suma_total_sal_crud += total
            sim_data_sal_crud.append([mes, int(prod), dias, int(total)])
        promedio_sal_crud = suma_total_sal_crud / days_passed if days_passed > 0 else 0
        sim_data_sal_crud.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_sal_crud)} | Prom: {promedio_sal_crud:.2f}"])
        df_sim_sal_crud = pd.DataFrame(sim_data_sal_crud, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 24. Simulación Salina Cruz Gasolinas
        prod_dict_sal_gas = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_sal_gas.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_sal_gas[meses_nombres[i]] += p
                    break
        sim_data_sal_gas = []
        suma_total_sal_gas = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_sal_gas[mes]
            total = prod * dias
            suma_total_sal_gas += total
            sim_data_sal_gas.append([mes, int(prod), dias, int(total)])
        promedio_sal_gas = suma_total_sal_gas / days_passed if days_passed > 0 else 0
        sim_data_sal_gas.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_sal_gas)} | Prom: {promedio_sal_gas:.2f}"])
        df_sim_sal_gas = pd.DataFrame(sim_data_sal_gas, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 25. Simulación Salina Cruz Diesel
        prod_dict_sal_die = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_sal_die.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_sal_die[meses_nombres[i]] += p
                    break
        sim_data_sal_die = []
        suma_total_sal_die = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_sal_die[mes]
            total = prod * dias
            suma_total_sal_die += total
            sim_data_sal_die.append([mes, int(prod), dias, int(total)])
        promedio_sal_die = suma_total_sal_die / days_passed if days_passed > 0 else 0
        sim_data_sal_die.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_sal_die)} | Prom: {promedio_sal_die:.2f}"])
        df_sim_sal_die = pd.DataFrame(sim_data_sal_die, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 26. Simulación Salina Cruz Turbosina
        prod_dict_sal_turb = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_sal_turb.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_sal_turb[meses_nombres[i]] += p
                    break
        sim_data_sal_turb = []
        suma_total_sal_turb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_sal_turb[mes]
            total = prod * dias
            suma_total_sal_turb += total
            sim_data_sal_turb.append([mes, int(prod), dias, int(total)])
        promedio_sal_turb = suma_total_sal_turb / days_passed if days_passed > 0 else 0
        sim_data_sal_turb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_sal_turb)} | Prom: {promedio_sal_turb:.2f}"])
        df_sim_sal_turb = pd.DataFrame(sim_data_sal_turb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 27. Simulación Salina Cruz Combustoleo
        prod_dict_sal_comb = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_sal_comb.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_sal_comb[meses_nombres[i]] += p
                    break
        sim_data_sal_comb = []
        suma_total_sal_comb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_sal_comb[mes]
            total = prod * dias
            suma_total_sal_comb += total
            sim_data_sal_comb.append([mes, int(prod), dias, int(total)])
        promedio_sal_comb = suma_total_sal_comb / days_passed if days_passed > 0 else 0
        sim_data_sal_comb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_sal_comb)} | Prom: {promedio_sal_comb:.2f}"])
        df_sim_sal_comb = pd.DataFrame(sim_data_sal_comb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 28. Simulación Tula Crudo
        prod_dict_tula_crud = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_tula_crud.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_tula_crud[meses_nombres[i]] += p
                    break
        sim_data_tula_crud = []
        suma_total_tula_crud = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_tula_crud[mes]
            total = prod * dias
            suma_total_tula_crud += total
            sim_data_tula_crud.append([mes, int(prod), dias, int(total)])
        promedio_tula_crud = suma_total_tula_crud / days_passed if days_passed > 0 else 0
        sim_data_tula_crud.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_tula_crud)} | Prom: {promedio_tula_crud:.2f}"])
        df_sim_tula_crud = pd.DataFrame(sim_data_tula_crud, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 29. Simulación Tula Gasolinas
        prod_dict_tula_gas = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_tula_gas.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_tula_gas[meses_nombres[i]] += p
                    break
        sim_data_tula_gas = []
        suma_total_tula_gas = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_tula_gas[mes]
            total = prod * dias
            suma_total_tula_gas += total
            sim_data_tula_gas.append([mes, int(prod), dias, int(total)])
        promedio_tula_gas = suma_total_tula_gas / days_passed if days_passed > 0 else 0
        sim_data_tula_gas.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_tula_gas)} | Prom: {promedio_tula_gas:.2f}"])
        df_sim_tula_gas = pd.DataFrame(sim_data_tula_gas, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 30. Simulación Tula Diesel
        prod_dict_tula_die = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_tula_die.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_tula_die[meses_nombres[i]] += p
                    break
        sim_data_tula_die = []
        suma_total_tula_die = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_tula_die[mes]
            total = prod * dias
            suma_total_tula_die += total
            sim_data_tula_die.append([mes, int(prod), dias, int(total)])
        promedio_tula_die = suma_total_tula_die / days_passed if days_passed > 0 else 0
        sim_data_tula_die.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_tula_die)} | Prom: {promedio_tula_die:.2f}"])
        df_sim_tula_die = pd.DataFrame(sim_data_tula_die, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 31. Simulación Tula Turbosina
        prod_dict_tula_turb = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_tula_turb.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_tula_turb[meses_nombres[i]] += p
                    break
        sim_data_tula_turb = []
        suma_total_tula_turb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_tula_turb[mes]
            total = prod * dias
            suma_total_tula_turb += total
            sim_data_tula_turb.append([mes, int(prod), dias, int(total)])
        promedio_tula_turb = suma_total_tula_turb / days_passed if days_passed > 0 else 0
        sim_data_tula_turb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_tula_turb)} | Prom: {promedio_tula_turb:.2f}"])
        df_sim_tula_turb = pd.DataFrame(sim_data_tula_turb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 32. Simulación Tula Combustoleo
        prod_dict_tula_comb = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_tula_comb.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_tula_comb[meses_nombres[i]] += p
                    break
        sim_data_tula_comb = []
        suma_total_tula_comb = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_tula_comb[mes]
            total = prod * dias
            suma_total_tula_comb += total
            sim_data_tula_comb.append([mes, int(prod), dias, int(total)])
        promedio_tula_comb = suma_total_tula_comb / days_passed if days_passed > 0 else 0
        sim_data_tula_comb.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_tula_comb)} | Prom: {promedio_tula_comb:.2f}"])
        df_sim_tula_comb = pd.DataFrame(sim_data_tula_comb, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 33. Simulación Olmeca Crudo
        prod_dict_olme_crud = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_olme_crud.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_olme_crud[meses_nombres[i]] += p
                    break
        sim_data_olme_crud = []
        suma_total_olme_crud = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_olme_crud[mes]
            total = prod * dias
            suma_total_olme_crud += total
            sim_data_olme_crud.append([mes, int(prod), dias, int(total)])
        promedio_olme_crud = suma_total_olme_crud / days_passed if days_passed > 0 else 0
        sim_data_olme_crud.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_olme_crud)} | Prom: {promedio_olme_crud:.2f}"])
        df_sim_olme_crud = pd.DataFrame(sim_data_olme_crud, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 34. Simulación Olmeca Gasolinas
        prod_dict_olme_gas = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_olme_gas.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_olme_gas[meses_nombres[i]] += p
                    break
        sim_data_olme_gas = []
        suma_total_olme_gas = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_olme_gas[mes]
            total = prod * dias
            suma_total_olme_gas += total
            sim_data_olme_gas.append([mes, int(prod), dias, int(total)])
        promedio_olme_gas = suma_total_olme_gas / days_passed if days_passed > 0 else 0
        sim_data_olme_gas.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_olme_gas)} | Prom: {promedio_olme_gas:.2f}"])
        df_sim_olme_gas = pd.DataFrame(sim_data_olme_gas, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        # 35. Simulación Olmeca Diesel
        prod_dict_olme_die = {m: 0.0 for m in meses_nombres}
        for idx, row_data in df_prod_olme_die.iterrows():
            val_anio = str(row_data.iloc[0]).strip().lower()
            val_prod = row_data.iloc[1]
            try: p = float(val_prod)
            except: p = 0.0
            for i, (m_largo, m_corto) in enumerate(zip(meses_nombres, meses_cortos)):
                if m_largo.lower() in val_anio or m_corto.lower() in val_anio:
                    prod_dict_olme_die[meses_nombres[i]] += p
                    break
        sim_data_olme_die = []
        suma_total_olme_die = 0.0
        for i, mes in enumerate(meses_nombres):
            dias = dias_por_mes[i]
            prod = prod_dict_olme_die[mes]
            total = prod * dias
            suma_total_olme_die += total
            sim_data_olme_die.append([mes, int(prod), dias, int(total)])
        promedio_olme_die = suma_total_olme_die / days_passed if days_passed > 0 else 0
        sim_data_olme_die.append(["TOTALES", "---", f"Días pasados: {days_passed}", f"Suma: {int(suma_total_olme_die)} | Prom: {promedio_olme_die:.2f}"])
        df_sim_olme_die = pd.DataFrame(sim_data_olme_die, columns=["Mes", "Producción", "Días", "Total (Prod x Días)"])

        def post_process_sim_and_prod(proceso_name, df_sim_val, df_prod_val):
            if df_sim_val is None or df_sim_val.empty:
                return df_sim_val, df_prod_val
            import db_helper
            mods = db_helper.get_modificaciones(proceso_name)
            sim_mods = mods.get("simulacion", {})
            
            # Determinar el año que representa la hoja de cálculo (sheet_year)
            sheet_year = 2026
            if df_prod_val is not None and not df_prod_val.empty:
                years_found = []
                for r in df_prod_val.to_numpy().tolist():
                    val = str(r[0]).strip()
                    if val.isdigit() and len(val) == 4:
                        years_found.append(int(val))
                if years_found:
                    sheet_year = max(years_found) + 1
            
            rows_sim = df_sim_val.to_numpy().tolist()
            for idx in range(12):
                if idx >= len(rows_sim) - 1:
                    break
                row = rows_sim[idx]
                mes = str(row[0]).strip()
                
                # Claves específicas del año
                prod_key = f"Sim_{sheet_year}_{mes}_Prod"
                dias_key = f"Sim_{sheet_year}_{mes}_Dias"
                
                prod_val = row[1]
                dias_val = row[2]
                
                # Cargar producción (con fallback heredado de la versión anterior para 2026)
                if prod_key in sim_mods:
                    try: prod_val = int(round(float(sim_mods[prod_key])))
                    except: pass
                elif sheet_year == 2026 and f"Sim_{mes}_Prod" in sim_mods:
                    try: prod_val = int(round(float(sim_mods[f"Sim_{mes}_Prod"])))
                    except: pass
                    
                # Cargar días (con fallback heredado de la versión anterior para 2026)
                if dias_key in sim_mods:
                    try: dias_val = int(round(float(sim_mods[dias_key])))
                    except: pass
                elif sheet_year == 2026 and f"Sim_{mes}_Dias" in sim_mods:
                    try: dias_val = int(round(float(sim_mods[f"Sim_{mes}_Dias"])))
                    except: pass
                    
                try: total_val = int(prod_val * dias_val)
                except: total_val = 0
                rows_sim[idx][1] = prod_val
                rows_sim[idx][2] = dias_val
                rows_sim[idx][3] = total_val
                
            suma_total = 0.0
            total_dias = 0
            for idx in range(12):
                if idx >= len(rows_sim) - 1:
                    break
                try: suma_total += float(rows_sim[idx][3])
                except: pass
                try: total_dias += int(rows_sim[idx][2])
                except: pass
            promedio = suma_total / total_dias if total_dias > 0 else 0.0
            if len(rows_sim) > 0:
                rows_sim[-1][0] = "TOTALES"
                rows_sim[-1][1] = "---"
                rows_sim[-1][2] = f"Días pasados: {total_dias}"
                rows_sim[-1][3] = f"Suma: {int(suma_total)} | Prom: {promedio:.2f}"
            df_sim_new = pd.DataFrame(rows_sim, columns=df_sim_val.columns)
            
            # Agregar o actualizar la fila de sheet_year en df_prod
            df_prod_new = df_prod_val
            if df_prod_val is not None and not df_prod_val.empty:
                rows_prod = df_prod_val.to_numpy().tolist()
                year_str = str(sheet_year)
                year_idx = -1
                for idx, r in enumerate(rows_prod):
                    if str(r[0]).strip() == year_str:
                        year_idx = idx
                        break
                if year_idx != -1:
                    rows_prod[year_idx][1] = int(round(promedio))
                else:
                    # Encontrar posición de inserción: justo después de todos los años menores y antes del primer mes
                    insert_pos = 0
                    for idx, r in enumerate(rows_prod):
                        val0 = str(r[0]).strip()
                        if val0.isdigit() and len(val0) == 4:
                            if int(val0) < sheet_year:
                                insert_pos = idx + 1
                            else:
                                insert_pos = idx
                                break
                        else:
                            # Hemos llegado a las filas de meses, insertamos aquí
                            break
                    rows_prod.insert(insert_pos, [year_str, int(round(promedio))])
                df_prod_new = pd.DataFrame(rows_prod, columns=df_prod_val.columns)
            return df_sim_new, df_prod_new

        df_sim, df_prod_copy = post_process_sim_and_prod("Crudo", df_sim, df_prod_copy)
        df_sim_gasolinas, df_prod_gasolinas_copy = post_process_sim_and_prod("Gasolinas", df_sim_gasolinas, df_prod_gasolinas_copy)
        df_sim_diesel, df_prod_diesel_copy = post_process_sim_and_prod("Diesel", df_sim_diesel, df_prod_diesel_copy)
        df_sim_turbosina, df_prod_turbosina_copy = post_process_sim_and_prod("Turbosina", df_sim_turbosina, df_prod_turbosina_copy)
        df_sim_asfalto, df_prod_asfalto_copy = post_process_sim_and_prod("Asfalto", df_sim_asfalto, df_prod_asfalto_copy)
        df_sim_combustoleo, df_prod_combustoleo_copy = post_process_sim_and_prod("Combustoleo", df_sim_combustoleo, df_prod_combustoleo_copy)
        df_sim_cad_gas, df_prod_cad_gas_copy = post_process_sim_and_prod("Cadereyta -Gasolinas", df_sim_cad_gas, df_prod_cad_gas_copy)
        df_sim_cad_die, df_prod_cad_die_copy = post_process_sim_and_prod("Cadereyta -Diesel", df_sim_cad_die, df_prod_cad_die_copy)
        df_sim_cad, df_prod_cad_copy = post_process_sim_and_prod("Cadereyta -Crudo", df_sim_cad, df_prod_cad_copy)
        df_sim_cad_comb, df_prod_cad_comb_copy = post_process_sim_and_prod("Cadereyta -Combustoleo", df_sim_cad_comb, df_prod_cad_comb_copy)
        df_sim_mad_crud, df_prod_mad_crud_copy = post_process_sim_and_prod("Madero -Crudo", df_sim_mad_crud, df_prod_mad_crud_copy)
        df_sim_mad_gas, df_prod_mad_gas_copy = post_process_sim_and_prod("Madero -Gasolinas", df_sim_mad_gas, df_prod_mad_gas_copy)
        df_sim_mad_die, df_prod_mad_die_copy = post_process_sim_and_prod("Madero -Diesel", df_sim_mad_die, df_prod_mad_die_copy)
        df_sim_mad_turb, df_prod_mad_turb_copy = post_process_sim_and_prod("Madero -Turbosina", df_sim_mad_turb, df_prod_mad_turb_copy)
        df_sim_mad_comb, df_prod_mad_comb_copy = post_process_sim_and_prod("Madero -Combustoleo", df_sim_mad_comb, df_prod_mad_comb_copy)
        df_sim_mina_crud, df_prod_mina_crud_copy = post_process_sim_and_prod("Minatitlan -Crudo", df_sim_mina_crud, df_prod_mina_crud_copy)
        df_sim_mina_gas, df_prod_mina_gas_copy = post_process_sim_and_prod("Minatitlan -Gasolinas", df_sim_mina_gas, df_prod_mina_gas_copy)
        df_sim_mina_die, df_prod_mina_die_copy = post_process_sim_and_prod("Minatitlan -Diesel", df_sim_mina_die, df_prod_mina_die_copy)
        df_sim_mina_comb, df_prod_mina_comb_copy = post_process_sim_and_prod("Minatitlan -Combustoleo", df_sim_mina_comb, df_prod_mina_comb_copy)
        df_sim_sala_crud, df_prod_sala_crud_copy = post_process_sim_and_prod("Salamanca -Crudo", df_sim_sala_crud, df_prod_sala_crud_copy)
        df_sim_sala_gas, df_prod_sala_gas_copy = post_process_sim_and_prod("Salamanca -Gasolinas", df_sim_sala_gas, df_prod_sala_gas_copy)
        df_sim_sala_die, df_prod_sala_die_copy = post_process_sim_and_prod("Salamanca -Diesel", df_sim_sala_die, df_prod_sala_die_copy)
        df_sim_sala_turb, df_prod_sala_turb_copy = post_process_sim_and_prod("Salamanca -Turbosina", df_sim_sala_turb, df_prod_sala_turb_copy)
        df_sim_sala_comb, df_prod_sala_comb_copy = post_process_sim_and_prod("Salamanca -Combustoleo", df_sim_sala_comb, df_prod_sala_comb_copy)
        df_sim_sal_crud, df_prod_sal_crud_copy = post_process_sim_and_prod("Salina Cruz -Crudo", df_sim_sal_crud, df_prod_sal_crud_copy)
        df_sim_sal_gas, df_prod_sal_gas_copy = post_process_sim_and_prod("Salina Cruz -Gasolinas", df_sim_sal_gas, df_prod_sal_gas_copy)
        df_sim_sal_die, df_prod_sal_die_copy = post_process_sim_and_prod("Salina Cruz -Diesel", df_sim_sal_die, df_prod_sal_die_copy)
        df_sim_sal_turb, df_prod_sal_turb_copy = post_process_sim_and_prod("Salina Cruz -Turbosina", df_sim_sal_turb, df_prod_sal_turb_copy)
        df_sim_sal_comb, df_prod_sal_comb_copy = post_process_sim_and_prod("Salina Cruz -Combustoleo", df_sim_sal_comb, df_prod_sal_comb_copy)
        df_sim_tula_crud, df_prod_tula_crud_copy = post_process_sim_and_prod("Tula -Crudo", df_sim_tula_crud, df_prod_tula_crud_copy)
        df_sim_tula_gas, df_prod_tula_gas_copy = post_process_sim_and_prod("Tula -Gasolinas", df_sim_tula_gas, df_prod_tula_gas_copy)
        df_sim_tula_die, df_prod_tula_die_copy = post_process_sim_and_prod("Tula -Diesel", df_sim_tula_die, df_prod_tula_die_copy)
        df_sim_tula_turb, df_prod_tula_turb_copy = post_process_sim_and_prod("Tula -Turbosina", df_sim_tula_turb, df_prod_tula_turb_copy)
        df_sim_tula_comb, df_prod_tula_comb_copy = post_process_sim_and_prod("Tula -Combustoleo", df_sim_tula_comb, df_prod_tula_comb_copy)
        df_sim_olme_crud, df_prod_olme_crud_copy = post_process_sim_and_prod("Olmeca -Crudo", df_sim_olme_crud, df_prod_olme_crud_copy)
        df_sim_olme_gas, df_prod_olme_gas_copy = post_process_sim_and_prod("Olmeca -Gasolinas", df_sim_olme_gas, df_prod_olme_gas_copy)
        df_sim_olme_die, df_prod_olme_die_copy = post_process_sim_and_prod("Olmeca -Diesel", df_sim_olme_die, df_prod_olme_die_copy)

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
                   df_data_mina_comb, df_snr_mina_comb_copy, df_prod_mina_comb_copy, df_sim_mina_comb,
                   df_data_sala_crud, df_snr_sala_crud_copy, df_prod_sala_crud_copy, df_sim_sala_crud,
                   df_data_sala_gas, df_snr_sala_gas_copy, df_prod_sala_gas_copy, df_sim_sala_gas,
                   df_data_sala_die, df_snr_sala_die_copy, df_prod_sala_die_copy, df_sim_sala_die,
                   df_data_sala_turb, df_snr_sala_turb_copy, df_prod_sala_turb_copy, df_sim_sala_turb,
                   df_data_sala_comb, df_snr_sala_comb_copy, df_prod_sala_comb_copy, df_sim_sala_comb,
                   df_data_sal_crud, df_snr_sal_crud_copy, df_prod_sal_crud_copy, df_sim_sal_crud,
                   df_data_sal_gas, df_snr_sal_gas_copy, df_prod_sal_gas_copy, df_sim_sal_gas,
                   df_data_sal_die, df_snr_sal_die_copy, df_prod_sal_die_copy, df_sim_sal_die,
                   df_data_sal_turb, df_snr_sal_turb_copy, df_prod_sal_turb_copy, df_sim_sal_turb,
                   df_data_sal_comb, df_snr_sal_comb_copy, df_prod_sal_comb_copy, df_sim_sal_comb,
                   df_data_tula_crud, df_snr_tula_crud_copy, df_prod_tula_crud_copy, df_sim_tula_crud,
                   df_data_tula_gas, df_snr_tula_gas_copy, df_prod_tula_gas_copy, df_sim_tula_gas,
                   df_data_tula_die, df_snr_tula_die_copy, df_prod_tula_die_copy, df_sim_tula_die,
                   df_data_tula_turb, df_snr_tula_turb_copy, df_prod_tula_turb_copy, df_sim_tula_turb,
                   df_data_tula_comb, df_snr_tula_comb_copy, df_prod_tula_comb_copy, df_sim_tula_comb,
                   df_data_olme_crud, df_snr_olme_crud_copy, df_prod_olme_crud_copy, df_sim_olme_crud,
                   df_data_olme_gas, df_snr_olme_gas_copy, df_prod_olme_gas_copy, df_sim_olme_gas,
                   df_data_olme_die, df_snr_olme_die_copy, df_prod_olme_die_copy, df_sim_olme_die)

    except Exception as e:
        err_details = traceback.format_exc()
        app.after(0, app.on_load_error, str(e), err_details)
