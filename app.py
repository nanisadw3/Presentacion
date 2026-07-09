import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox
import os
from CTkTable import CTkTable
import threading
import traceback
from termcolor import colored

class ExcelViewerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Visor de Excel (Nativo de CustomTkinter)")
        self.geometry("1500x850")
        
        # Configuración del tema
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        # Contenedor principal
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Header Frame
        self.top_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        self.top_frame.pack(fill="x", pady=(0, 20))

        self.btn_buscar = ctk.CTkButton(self.top_frame, 
                                        text="Buscar Archivo Excel", 
                                        font=("Roboto", 14, "bold"),
                                        command=self.load_excel,
                                        height=40)
        self.btn_buscar.pack(pady=15, padx=20, side="left")

        self.lbl_file = ctk.CTkLabel(self.top_frame, 
                                     text="Ningún archivo seleccionado",
                                     font=("Roboto", 14))
        self.lbl_file.pack(pady=15, padx=20, side="left")

        # Botón para guardar las tablas en la Base de Datos SQLite
        self.btn_guardar = ctk.CTkButton(self.top_frame, 
                                        text="Guardar en Base de Datos", 
                                        font=("Roboto", 14, "bold"),
                                        command=self.save_to_database,
                                        height=40,
                                        fg_color="#28a745", hover_color="#218838")
        self.btn_guardar.pack(pady=15, padx=10, side="left")

        # Botón para mandar la información a PowerPoint
        self.btn_powerpoint = ctk.CTkButton(self.top_frame, 
                                            text="Mandar a PowerPoint", 
                                            font=("Roboto", 14, "bold"),
                                            command=self.send_to_powerpoint,
                                            height=40,
                                            fg_color="#8b5cf6", hover_color="#7c3aed")
        self.btn_powerpoint.pack(pady=15, padx=10, side="left")

        # ComboBox para alternar visualización de procesos
        self.lbl_proceso = ctk.CTkLabel(self.top_frame, text="Proceso:", font=("Roboto", 14, "bold"))
        self.lbl_proceso.pack(pady=15, padx=(20, 5), side="left")
        
        self.cb_proceso = ctk.CTkComboBox(self.top_frame, 
                                          values=["Crudo", "Gasolinas", "Diesel", "Turbosina"],
                                          font=("Roboto", 14),
                                          command=self.on_proceso_changed,
                                          state="readonly",
                                          width=150)
        self.cb_proceso.pack(pady=15, padx=5, side="left")
        self.cb_proceso.set("Crudo")

        # Scrollable Frame para contener la tabla
        self.scroll_frame = ctk.CTkScrollableFrame(self.main_frame, corner_radius=10)
        self.scroll_frame.pack(fill="both", expand=True)

        self.table = None
        self.df_data = None # Para la primera tabla
        self.df_snr = None  # Para la segunda tabla (AE-AF)
        self.df_prod = None # Para la tercera tabla (AE-AF, 21-40)
        self.df_sim = None  # Para la cuarta tabla (Simulación 12 meses)
        
        # Datos de Gasolinas
        self.df_data_gasolinas = None
        self.df_snr_gasolinas = None
        self.df_prod_gasolinas = None
        self.df_sim_gasolinas = None

        # Datos de Diesel
        self.df_data_diesel = None
        self.df_snr_diesel = None
        self.df_prod_diesel = None
        self.df_sim_diesel = None

        # Datos de Turbosina
        self.df_data_turbosina = None
        self.df_snr_turbosina = None
        self.df_prod_turbosina = None
        self.df_sim_turbosina = None

        # Definir directorio por defecto (con fallback si no existe el de descargas)
        downloads_path = "/mnt/c/Users/10900096799/Downloads"
        if os.path.exists(downloads_path):
            self.default_dir = downloads_path
        else:
            self.default_dir = os.path.dirname(os.path.abspath(__file__))

        # Barra de progreso (inicialmente oculta)
        self.progress_bar = ctk.CTkProgressBar(self.top_frame, width=200)
        self.progress_bar.set(0.0)

    def set_loading_state(self, is_loading, loading_text=""):
        if is_loading:
            self.lbl_file.configure(text=loading_text)
            self.btn_buscar.configure(state="disabled")
            self.btn_guardar.configure(state="disabled")
            self.btn_powerpoint.configure(state="disabled")
            self.progress_bar.pack(pady=15, padx=20, side="left")
            self.progress_bar.set(0.0)
        else:
            self.btn_buscar.configure(state="normal")
            self.btn_guardar.configure(state="normal")
            self.btn_powerpoint.configure(state="normal")
            self.progress_bar.pack_forget()

    def update_progress(self, val, text_msg=None):
        self.progress_bar.set(val)
        if text_msg is not None:
            self.lbl_file.configure(text=text_msg)

    def save_to_database(self):
        if self.df_data is None or self.df_snr is None or self.df_prod is None or self.df_sim is None:
            messagebox.showerror("Error", "Primero debes buscar un archivo Excel para extraer los datos.")
            return

        db_path = filedialog.asksaveasfilename(
            title="Crear o seleccionar Base de Datos SQLite",
            initialdir=self.default_dir,
            initialfile="datos_extraidos.db",
            defaultextension=".db",
            filetypes=[("SQLite Database", "*.db *.sqlite")]
        )

        if db_path:
            import sqlite3
            try:
                conn = sqlite3.connect(db_path)
                
                # Guardar datos de Crudo
                self.df_data.to_sql('crudo_tabla_principal', conn, if_exists='replace', index=False)
                self.df_snr.to_sql('crudo_programa_snr', conn, if_exists='replace', index=False)
                self.df_prod.to_sql('crudo_produccion', conn, if_exists='replace', index=False)
                self.df_sim.to_sql('crudo_simulacion_anual', conn, if_exists='replace', index=False)
                
                 # Guardar datos de Gasolinas
                if self.df_data_gasolinas is not None:
                    self.df_data_gasolinas.to_sql('gasolinas_tabla_principal', conn, if_exists='replace', index=False)
                    self.df_snr_gasolinas.to_sql('gasolinas_programa_snr', conn, if_exists='replace', index=False)
                    self.df_prod_gasolinas.to_sql('gasolinas_produccion', conn, if_exists='replace', index=False)
                    if hasattr(self, 'df_sim_gasolinas') and self.df_sim_gasolinas is not None:
                        self.df_sim_gasolinas.to_sql('gasolinas_simulacion_anual', conn, if_exists='replace', index=False)
                
                # Guardar datos de Diesel
                if self.df_data_diesel is not None:
                    self.df_data_diesel.to_sql('diesel_tabla_principal', conn, if_exists='replace', index=False)
                    self.df_snr_diesel.to_sql('diesel_programa_snr', conn, if_exists='replace', index=False)
                    self.df_prod_diesel.to_sql('diesel_produccion', conn, if_exists='replace', index=False)
                    if hasattr(self, 'df_sim_diesel') and self.df_sim_diesel is not None:
                        self.df_sim_diesel.to_sql('diesel_simulacion_anual', conn, if_exists='replace', index=False)

                # Guardar datos de Turbosina
                if self.df_data_turbosina is not None:
                    self.df_data_turbosina.to_sql('turbosina_tabla_principal', conn, if_exists='replace', index=False)
                    self.df_snr_turbosina.to_sql('turbosina_programa_snr', conn, if_exists='replace', index=False)
                    self.df_prod_turbosina.to_sql('turbosina_produccion', conn, if_exists='replace', index=False)
                    if hasattr(self, 'df_sim_turbosina') and self.df_sim_turbosina is not None:
                        self.df_sim_turbosina.to_sql('turbosina_simulacion_anual', conn, if_exists='replace', index=False)

                conn.close()
                messagebox.showinfo("Éxito", f"¡Los datos de todos los procesos han sido guardados en la base de datos!\n\nRuta:\n{db_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al guardar en SQLite:\n{str(e)}")

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            initialdir=self.default_dir,
            filetypes=[("Archivos de Excel", "*.xlsx *.xls *.xlsm")]
        )

        if file_path:
            self.set_loading_state(True, "Cargando datos del Excel...")
            # Correr carga en segundo plano
            threading.Thread(target=self.async_load_data, args=(file_path,), daemon=True).start()

    def async_load_data(self, file_path):
        try:
            import warnings
            warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

            # Función segura para quitar decimales sin causar errores de tipo
            def remove_decimals(df_to_clean):
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

            # --- 1. PROCESAR CRUDO ---
            # Leer fila 1 (index 0) para encabezados, Cols A:H (0:8)
            clean_headers = get_clean_headers(0, 0, 8)
            # Leer Tabla 1 (Rows 21-51 -> index 20:51), Cols A:H (0:8)
            df = df_sheet.iloc[20:51, 0:8].copy()
            df.columns = clean_headers
            df = df.dropna(how='all').dropna(axis=1, how='all')
            df = remove_decimals(df)
            df_data = df.copy()

            # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Cols AE:AF (30:32)
            df_snr = df_sheet.iloc[73:104, 30:32].copy()
            df_snr = df_snr.dropna(how='all').dropna(axis=1, how='all')
            df_snr = remove_decimals(df_snr)
            df_snr_copy = df_snr.copy()

            # --- Recortar filas de fin de mes si el mes tiene menos de 31 días ---
            snr_col = None
            for col in df_data.columns:
                if "SNR" in str(col).upper():
                    snr_col = col
                    break

            num_dias_reales = 31
            if snr_col:
                for i in range(30, -1, -1):
                    val_snr = 0
                    if i < len(df_data):
                        try:
                            raw_val = df_data[snr_col].iloc[i]
                            if raw_val != "":
                                val_snr = float(raw_val)
                        except:
                            val_snr = 0
                    if val_snr != 0:
                        num_dias_reales = i + 1
                        break

            df_data = df_data.iloc[:num_dias_reales]
            df_snr_copy = df_snr_copy.iloc[:num_dias_reales]

            # Leer Tabla 3 (Años y Producción, Rows 21-120 -> index 20:120), Cols AE:AF (30:32)
            df_prod_raw = df_sheet.iloc[20:120, 30:32].copy()
            df_prod_raw = df_prod_raw.dropna(how='all')
            
            dic_idx = -1
            for idx, row in df_prod_raw.iterrows():
                val = str(row.iloc[0]).strip().lower()
                if "dic" in val or "diciembre" in val:
                    dic_idx = idx - 20 # Ajustar al índice relativo de df_prod_raw
                    break
            
            if dic_idx != -1:
                df_prod = df_prod_raw.iloc[:dic_idx + 1]
            else:
                df_prod = df_prod_raw.iloc[:20]
                
            df_prod = df_prod.dropna(axis=1, how='all')
            df_prod = remove_decimals(df_prod)
            df_prod_copy = df_prod.copy()


            # --- 2. PROCESAR GASOLINAS ---
            # Leer fila 1 (index 0) para encabezados, Cols L:S (11:19)
            clean_headers_gas = get_clean_headers(0, 11, 19)
            # Leer Tabla 1 (Rows 21-51 -> index 20:51), Cols L:S (11:19)
            df_gas = df_sheet.iloc[20:51, 11:19].copy()
            df_gas.columns = clean_headers_gas
            df_gas = df_gas.dropna(how='all').dropna(axis=1, how='all')
            df_gas = remove_decimals(df_gas)
            df_data_gasolinas = df_gas.copy()

            # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Cols AK:AL (36:38)
            df_snr_gas = df_sheet.iloc[73:104, 36:38].copy()
            df_snr_gas = df_snr_gas.dropna(how='all').dropna(axis=1, how='all')
            df_snr_gas = remove_decimals(df_snr_gas)
            df_snr_gas_copy = df_snr_gas.copy()

            # Recortar filas de fin de mes
            df_data_gasolinas = df_data_gasolinas.iloc[:num_dias_reales]
            df_snr_gas_copy = df_snr_gas_copy.iloc[:num_dias_reales]

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
            df_prod_gas = remove_decimals(df_prod_gas)
            df_prod_gasolinas_copy = df_prod_gas.copy()


            # --- 3. PROCESAR DIESEL ---
            # Leer fila 54 (index 53) para encabezados, Cols A:H (0:8)
            clean_headers_die = get_clean_headers(53, 0, 8)
            # Leer Tabla 1 (Rows 74-104 -> index 73:104), Cols A:H (0:8)
            df_die = df_sheet.iloc[73:104, 0:8].copy()
            df_die.columns = clean_headers_die
            df_die = df_die.dropna(how='all').dropna(axis=1, how='all')
            df_die = remove_decimals(df_die)
            df_data_diesel = df_die.copy()

            # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Cols AQ:AR (42:44)
            df_snr_die = df_sheet.iloc[73:104, 42:44].copy()
            df_snr_die = df_snr_die.dropna(how='all').dropna(axis=1, how='all')
            df_snr_die = remove_decimals(df_snr_die)
            df_snr_die_copy = df_snr_die.copy()

            # Recortar filas de fin de mes
            df_data_diesel = df_data_diesel.iloc[:num_dias_reales]
            df_snr_die_copy = df_snr_die_copy.iloc[:num_dias_reales]

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
            df_prod_die = remove_decimals(df_prod_die)
            df_prod_diesel_copy = df_prod_die.copy()


            # --- 4. PROCESAR TURBOSINA ---
            # Leer fila 54 (index 53) para encabezados, Cols L:Q (11:17)
            clean_headers_turb = get_clean_headers(53, 11, 17)
            # Leer Tabla 1 (Rows 74-104 -> index 73:104), Cols L:Q (11:17)
            df_turb = df_sheet.iloc[73:104, 11:17].copy()
            df_turb.columns = clean_headers_turb
            df_turb = df_turb.dropna(how='all').dropna(axis=1, how='all')
            df_turb = remove_decimals(df_turb)
            df_data_turbosina = df_turb.copy()

            # Leer Tabla 2 (Programa, Rows 74-104 -> index 73:104), Cols AW:AX (48:50)
            df_snr_turb = df_sheet.iloc[73:104, 48:50].copy()
            df_snr_turb = df_snr_turb.dropna(how='all').dropna(axis=1, how='all')
            df_snr_turb = remove_decimals(df_snr_turb)
            df_snr_turb_copy = df_snr_turb.copy()

            # Recortar filas de fin de mes
            df_data_turbosina = df_data_turbosina.iloc[:num_dias_reales]
            df_snr_turb_copy = df_snr_turb_copy.iloc[:num_dias_reales]

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
            df_prod_turb = remove_decimals(df_prod_turb)
            df_prod_turbosina_copy = df_prod_turb.copy()


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

            # Pasar datos a la interfaz (main thread)
            self.after(0, self.on_load_success, file_path, df_data, df_snr_copy, df_prod_copy, df_sim,
                       df_data_gasolinas, df_snr_gas_copy, df_prod_gasolinas_copy, df_sim_gasolinas,
                       df_data_diesel, df_snr_die_copy, df_prod_diesel_copy, df_sim_diesel,
                       df_data_turbosina, df_snr_turb_copy, df_prod_turbosina_copy, df_sim_turbosina)

        except Exception as e:
            err_details = traceback.format_exc()
            self.after(0, self.on_load_error, str(e), err_details)

    def on_load_success(self, file_path, df_data, df_snr, df_prod, df_sim,
                        df_data_gasolinas=None, df_snr_gasolinas=None, df_prod_gasolinas=None, df_sim_gasolinas=None,
                        df_data_diesel=None, df_snr_diesel=None, df_prod_diesel=None, df_sim_diesel=None,
                        df_data_turbosina=None, df_snr_turbosina=None, df_prod_turbosina=None, df_sim_turbosina=None):
        self.df_data = df_data
        self.df_snr = df_snr
        self.df_prod = df_prod
        self.df_sim = df_sim
        
        self.df_data_gasolinas = df_data_gasolinas
        self.df_snr_gasolinas = df_snr_gasolinas
        self.df_prod_gasolinas = df_prod_gasolinas
        self.df_sim_gasolinas = df_sim_gasolinas

        self.df_data_diesel = df_data_diesel
        self.df_snr_diesel = df_snr_diesel
        self.df_prod_diesel = df_prod_diesel
        self.df_sim_diesel = df_sim_diesel

        self.df_data_turbosina = df_data_turbosina
        self.df_snr_turbosina = df_snr_turbosina
        self.df_prod_turbosina = df_prod_turbosina
        self.df_sim_turbosina = df_sim_turbosina

        # Mostrar las tablas correspondientes a la selección actual del ComboBox
        self.show_proceso_tables(self.cb_proceso.get())

        self.set_loading_state(False)
        self.lbl_file.configure(text=f"Archivo: {os.path.basename(file_path)}")

    def on_proceso_changed(self, selection):
        if self.df_data is None:
            return
        self.show_proceso_tables(selection)

    def show_proceso_tables(self, selection):
        # Limpiar tablas actuales
        if self.table is not None:
            self.table.destroy()
            self.table = None
        if hasattr(self, 'table2') and self.table2 is not None:
            self.table2.destroy()
            self.table2 = None
        if hasattr(self, 'lbl_table2') and self.lbl_table2 is not None:
            self.lbl_table2.destroy()
        if hasattr(self, 'table3') and self.table3 is not None:
            self.table3.destroy()
            self.table3 = None
        if hasattr(self, 'lbl_table3') and self.lbl_table3 is not None:
            self.lbl_table3.destroy()
        if hasattr(self, 'table4') and self.table4 is not None:
            self.table4.destroy()
            self.table4 = None
        if hasattr(self, 'lbl_table4') and self.lbl_table4 is not None:
            self.lbl_table4.destroy()

        if selection == "Crudo":
            df_data = self.df_data
            df_snr = self.df_snr
            df_prod = self.df_prod
            df_sim = self.df_sim
            lbl2_txt = "Programa del SNR (AE-AF, Filas 74-104)"
            lbl3_txt = "Fecha y Producción (AE-AF, Filas 21-40)"
        elif selection == "Gasolinas":
            df_data = self.df_data_gasolinas
            df_snr = self.df_snr_gasolinas
            df_prod = self.df_prod_gasolinas
            df_sim = self.df_sim_gasolinas
            lbl2_txt = "Programa de Gasolinas (AK-AL, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Gasolinas (AG-AH, Filas 21-40)"
        elif selection == "Diesel":
            df_data = self.df_data_diesel
            df_snr = self.df_snr_diesel
            df_prod = self.df_prod_diesel
            df_sim = self.df_sim_diesel
            lbl2_txt = "Programa de Diesel (AQ-AR, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Diesel (AK-AL, Filas 21-40)"
        else: # Turbosina
            df_data = self.df_data_turbosina
            df_snr = self.df_snr_turbosina
            df_prod = self.df_prod_turbosina
            df_sim = self.df_sim_turbosina
            lbl2_txt = "Programa de Turbosina (AW-AX, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Turbosina (AM-AN, Filas 21-40)"

        if df_data is None or df_snr is None or df_prod is None or df_sim is None:
            return

        # Dibujar Tabla 1
        headers = list(df_data.columns)
        rows = df_data.to_numpy().tolist()
        if len(rows) > 200: rows = rows[:200]
        table_values = [headers] + rows
        self.table = CTkTable(
            master=self.scroll_frame, 
            row=len(table_values), 
            column=len(table_values[0]), 
            values=table_values,
            header_color="#1f538d",
            colors=["#2a2a2a", "#242424"],
            hover_color="#3a3a3a",
        )
        self.table.pack(expand=True, fill="both", padx=10, pady=10)

        # Dibujar Tabla 2
        self.lbl_table2 = ctk.CTkLabel(self.scroll_frame, text=lbl2_txt, font=("Roboto", 16, "bold"), text_color="#3484F0")
        self.lbl_table2.pack(pady=(20, 5))
        headers2 = ["Columna 1", "Columna 2"]
        rows2 = df_snr.to_numpy().tolist()
        if len(rows2) > 200: rows2 = rows2[:200]
        table_values2 = [headers2] + rows2
        self.table2 = CTkTable(
            master=self.scroll_frame, 
            row=len(table_values2), 
            column=len(table_values2[0]), 
            values=table_values2,
            header_color="#1f538d",
            colors=["#2a2a2a", "#242424"],
            hover_color="#3a3a3a",
        )
        self.table2.pack(expand=True, fill="both", padx=10, pady=10)

        # Dibujar Tabla 3
        self.lbl_table3 = ctk.CTkLabel(self.scroll_frame, text=lbl3_txt, font=("Roboto", 16, "bold"), text_color="#3484F0")
        self.lbl_table3.pack(pady=(20, 5))
        headers3 = ["Año / Mes", "Producción"]
        rows3 = df_prod.to_numpy().tolist()
        if len(rows3) > 200: rows3 = rows3[:200]
        table_values3 = [headers3] + rows3
        self.table3 = CTkTable(
            master=self.scroll_frame, 
            row=len(table_values3), 
            column=len(table_values3[0]), 
            values=table_values3,
            header_color="#1f538d",
            colors=["#2a2a2a", "#242424"],
            hover_color="#3a3a3a",
        )
        self.table3.pack(expand=True, fill="both", padx=10, pady=10)

        # Dibujar Tabla 4
        self.lbl_table4 = ctk.CTkLabel(self.scroll_frame, text="Simulación de Producción Anual", font=("Roboto", 16, "bold"), text_color="#3484F0")
        self.lbl_table4.pack(pady=(30, 5))
        headers4 = list(df_sim.columns)
        rows4 = df_sim.to_numpy().tolist()
        table_values4 = [headers4] + rows4
        self.table4 = CTkTable(
            master=self.scroll_frame, 
            row=len(table_values4), 
            column=len(table_values4[0]), 
            values=table_values4,
            header_color="#1f538d",
            colors=["#2a2a2a", "#242424"],
            hover_color="#3a3a3a",
        )
        self.table4.pack(expand=True, fill="both", padx=10, pady=20)

    def on_load_error(self, err_msg, error_details):
        self.set_loading_state(False)
        self.lbl_file.configure(text="Error al cargar")
        
        try:
            print(colored(f"\n[!] ERROR EN HILO DE CARGA:\n{err_msg}\n\nTRACEBACK:\n{error_details}", "red"))
        except:
            print(f"\n[!] ERROR EN HILO DE CARGA:\n{err_msg}\n\nTRACEBACK:\n{error_details}")
            
        messagebox.showerror("Error Detallado", f"Ocurrió un error al leer el archivo:\n{err_msg}\n\nDetalles técnicos:\n{error_details}")

    def send_to_powerpoint(self):
        if self.df_data is None or self.df_snr is None or self.df_prod is None:
            messagebox.showerror("Error", "Primero debes buscar un archivo Excel para extraer los datos.")
            return

        file_path = filedialog.askopenfilename(
            title="Seleccionar plantilla PowerPoint",
            initialdir=self.default_dir,
            filetypes=[("Archivos PowerPoint", "*.pptx")]
        )

        if file_path:
            save_path = filedialog.asksaveasfilename(
                title="Guardar presentación actualizada",
                initialdir=self.default_dir,
                initialfile="Proceso y Producciones_Actualizado.pptx",
                defaultextension=".pptx",
                filetypes=[("Archivos PowerPoint", "*.pptx")]
            )

            if save_path:
                self.set_loading_state(True, "Inyectando datos a PowerPoint...")
                threading.Thread(target=self.async_send_to_pptx, args=(file_path, save_path), daemon=True).start()

    def update_slide_chart(self, chart, categories, proceso_vals, diario_vals, programa_vals, columna1_vals, wine_color, green_color):
        from pptx.chart.data import CategoryChartData
        from pptx.util import Pt

        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series('PROCESO', tuple(proceso_vals))
        chart_data.add_series('Diario', tuple(diario_vals))
        chart_data.add_series('Programa', tuple(programa_vals))
        chart_data.add_series('Columna1', tuple(columna1_vals))

        # Reemplazar los datos de la gráfica
        chart.replace_data(chart_data)

        # --- Restaurar y pintar los colores dinámicamente ---
        if len(chart.series) > 0:
            series = chart.series[0]
            
            # Buscar el último mes con producción real en la nueva lista de categorías
            last_month_idx = -1
            for i in range(17):
                if i < len(categories):
                    cat = categories[i]
                    val = proceso_vals[i]
                    # Si es un mes (contiene letras)
                    if any(c.isalpha() for c in cat):
                        if val is not None and val != 0:
                            last_month_idx = i

            # Eliminar cualquier formato específico (<c:dPt>) para los índices de años
            # Esto obliga a PowerPoint a usar el estilo/color por defecto de la serie (gris de la plantilla)
            try:
                ser_el = series._element
                ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
                dpts = ser_el.findall('c:dPt', ns)
                
                year_indices = set()
                for i, cat in enumerate(categories):
                    if not any(c.isalpha() for c in str(cat)):
                        year_indices.add(i)
                        
                for dpt in dpts:
                    idx_el = dpt.find('c:idx', ns)
                    if idx_el is not None:
                        idx = int(idx_el.attrib['val'])
                        if idx in year_indices:
                            ser_el.remove(dpt)
            except Exception:
                pass

            # Pintar únicamente las barras correspondientes a los meses
            for p_idx in range(min(17, len(series.points))):
                try:
                    cat = categories[p_idx]
                    # Si es un mes, aplicamos el color correspondiente (vino o verde)
                    if any(c.isalpha() for c in cat):
                        point = series.points[p_idx]
                        fill = point.format.fill
                        fill.solid()
                        if p_idx == last_month_idx:
                            fill.fore_color.rgb = wine_color
                            try:
                                point.data_label.font.size = Pt(14)
                                point.data_label.font.bold = True
                            except Exception:
                                pass
                        else:
                            fill.fore_color.rgb = green_color
                            
                        # Limpiar modificadores de color (lumMod/lumOff) heredados del punto original en la plantilla
                        try:
                            ns_a = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                            srgb_clr = point._element.find('.//a:srgbClr', ns_a)
                            if srgb_clr is not None:
                                for child in list(srgb_clr):
                                    if 'lumMod' in child.tag or 'lumOff' in child.tag:
                                        srgb_clr.remove(child)
                        except Exception:
                            pass
                except Exception:
                    pass

    def async_send_to_pptx(self, file_path, save_path):
        try:
            from pptx import Presentation
            from pptx.dml.color import RGBColor

            prs = Presentation(file_path)
            if len(prs.slides) < 5:
                raise ValueError("La presentación debe tener al menos 5 diapositivas (Crudo, Gasolinas, Diesel y Turbosina).")

            # --- 1. PROCESAR DIAPOSITIVA DE CRUDO (DIAPOSITIVA 2) ---
            slide = prs.slides[1]
            chart = None
            for shape in slide.shapes:
                if shape.has_chart:
                    chart = shape.chart
                    break

            if not chart:
                raise ValueError("No se encontró ninguna gráfica en la segunda diapositiva (Crudo).")

            snr_col = None
            for col in self.df_data.columns:
                if "SNR" in str(col).upper():
                    snr_col = col
                    break

            if not snr_col:
                raise ValueError("No se encontró la columna 'SNR' en la primera tabla.")

            categories = []
            proceso_vals = []
            diario_vals = []
            programa_vals = []
            columna1_vals = []

            # Filtrar df_prod para quedarnos con los años y los meses activos (con producción)
            prod_rows = []
            for idx, row in self.df_prod.iterrows():
                cat = str(row.iloc[0]).strip()
                val = row.iloc[1]
                
                if not cat:
                    continue
                    
                # Si es un año (no contiene letras)
                if not any(c.isalpha() for c in cat):
                    prod_rows.append((cat, val))
                else:
                    # Si es un mes, solo si tiene producción real (no vacío ni cero)
                    try:
                        p_val = float(val)
                    except:
                        p_val = 0
                    if p_val != 0:
                        prod_rows.append((cat, val))

            # Si hay más de 17 filas (años + meses activos), nos quedamos con las últimas 17 (los más recientes)
            if len(prod_rows) > 17:
                prod_rows = prod_rows[-17:]

            # Llenar filas de años y meses (sin celdas vacías al final)
            for i in range(len(prod_rows)):
                cat_val = prod_rows[i][0]
                try:
                    proc_val = float(prod_rows[i][1])
                except:
                    proc_val = None
                        
                categories.append(cat_val)
                proceso_vals.append(proc_val)
                diario_vals.append(None)
                programa_vals.append(None)
                columna1_vals.append(None)

            # Llenar filas diarias
            num_dias_reales = 31
            for i in range(30, -1, -1):
                val_snr = 0
                if i < len(self.df_data):
                    try:
                        raw_val = self.df_data[snr_col].iloc[i]
                        if raw_val != "":
                            val_snr = float(raw_val)
                        else:
                            val_snr = 0
                    except:
                        val_snr = 0
                if val_snr != 0:
                    num_dias_reales = i + 1
                    break

            for i in range(31):
                categories.append(str(i + 1))
                proceso_vals.append(None)
                
                if i >= num_dias_reales:
                    diario_vals.append(None)
                    programa_vals.append(None)
                    columna1_vals.append(None)
                    continue

                d_val = None
                if i < len(self.df_data):
                    try:
                        d_val = float(self.df_data[snr_col].iloc[i])
                    except:
                        d_val = None
                diario_vals.append(d_val)
                
                p_val = None
                if i < len(self.df_snr):
                    try:
                        p_val = float(self.df_snr.iloc[i, 0])
                    except:
                        p_val = None
                programa_vals.append(p_val)
                
                c_val = None
                if i < len(self.df_snr):
                    try:
                        c_val = float(self.df_snr.iloc[i, 1])
                    except:
                        c_val = None
                columna1_vals.append(c_val)

            # --- EXTRAER COLORES DE LA PLANTILLA DE FORMA DINÁMICA ---
            wine_color = None
            green_color = None

            if len(chart.series) > 0:
                series = chart.series[0]
                for p_idx in range(min(18, len(series.points))):
                    try:
                        point = series.points[p_idx]
                        fill = point.format.fill
                        if fill.type == 1: # SOLID
                            color_val = fill.fore_color.rgb
                            if p_idx == 15: # En la plantilla original, index 15 es Junio (vino)
                                wine_color = color_val
                            elif p_idx == 10: # En la plantilla original, index 10 es Enero (verde)
                                green_color = color_val
                    except Exception:
                        pass

            # Fallbacks seguros en caso de no encontrarse
            if not wine_color:
                wine_color = RGBColor(0x69, 0x19, 0x32)
            if not green_color:
                green_color = RGBColor(0x24, 0x5C, 0x4F)

            # Actualizar la gráfica de Crudo (Diapositiva 2)
            self.update_slide_chart(chart, categories, proceso_vals, diario_vals, programa_vals, columna1_vals, wine_color, green_color)


            # --- 2. PROCESAR DIAPOSITIVA DE GASOLINAS (DIAPOSITIVA 3) ---
            slide_gas = prs.slides[2]
            chart_gas = None
            for shape in slide_gas.shapes:
                if shape.has_chart:
                    chart_gas = shape.chart
                    break

            if not chart_gas:
                raise ValueError("No se encontró ninguna gráfica en la tercera diapositiva (Gasolinas).")

            snr_col_gas = None
            for col in self.df_data_gasolinas.columns:
                if "SNR" in str(col).upper():
                    snr_col_gas = col
                    break

            if not snr_col_gas:
                raise ValueError("No se encontró la columna 'SNR' en la tabla de Gasolinas.")

            categories_gas = []
            proceso_vals_gas = []
            diario_vals_gas = []
            programa_vals_gas = []
            columna1_vals_gas = []

            # Filtrar df_prod_gasolinas para quedarnos con los años y meses activos
            prod_rows_gas = []
            for idx, row in self.df_prod_gasolinas.iterrows():
                cat = str(row.iloc[0]).strip()
                val = row.iloc[1]
                
                if not cat:
                    continue
                    
                if not any(c.isalpha() for c in cat):
                    prod_rows_gas.append((cat, val))
                else:
                    try:
                        p_val = float(val)
                    except:
                        p_val = 0
                    if p_val != 0:
                        prod_rows_gas.append((cat, val))

            # Ajustar al límite de 17 categorías
            if len(prod_rows_gas) > 17:
                prod_rows_gas = prod_rows_gas[-17:]

            # Llenar filas de años y meses (sin celdas vacías al final)
            for i in range(len(prod_rows_gas)):
                cat_val = prod_rows_gas[i][0]
                try:
                    proc_val = float(prod_rows_gas[i][1])
                except:
                    proc_val = None
                        
                categories_gas.append(cat_val)
                proceso_vals_gas.append(proc_val)
                diario_vals_gas.append(None)
                programa_vals_gas.append(None)
                columna1_vals_gas.append(None)

            # Llenar filas diarias
            num_dias_reales_gas = 31
            for i in range(30, -1, -1):
                val_snr = 0
                if i < len(self.df_data_gasolinas):
                    try:
                        raw_val = self.df_data_gasolinas[snr_col_gas].iloc[i]
                        if raw_val != "":
                            val_snr = float(raw_val)
                        else:
                            val_snr = 0
                    except:
                        val_snr = 0
                if val_snr != 0:
                    num_dias_reales_gas = i + 1
                    break

            for i in range(31):
                categories_gas.append(str(i + 1))
                proceso_vals_gas.append(None)
                
                if i >= num_dias_reales_gas:
                    diario_vals_gas.append(None)
                    programa_vals_gas.append(None)
                    columna1_vals_gas.append(None)
                    continue

                d_val = None
                if i < len(self.df_data_gasolinas):
                    try:
                        d_val = float(self.df_data_gasolinas[snr_col_gas].iloc[i])
                    except:
                        d_val = None
                diario_vals_gas.append(d_val)
                
                p_val = None
                if i < len(self.df_snr_gasolinas):
                    try:
                        p_val = float(self.df_snr_gasolinas.iloc[i, 0])
                    except:
                        p_val = None
                programa_vals_gas.append(p_val)
                
                c_val = None
                if i < len(self.df_snr_gasolinas):
                    try:
                        c_val = float(self.df_snr_gasolinas.iloc[i, 1])
                    except:
                        c_val = None
                columna1_vals_gas.append(c_val)

            # Actualizar la gráfica de Gasolinas (Diapositiva 3)
            self.update_slide_chart(chart_gas, categories_gas, proceso_vals_gas, diario_vals_gas, programa_vals_gas, columna1_vals_gas, wine_color, green_color)


            # --- 3. PROCESAR DIAPOSITIVA DE DIESEL (DIAPOSITIVA 4) ---
            slide_die = prs.slides[3]
            chart_die = None
            for shape in slide_die.shapes:
                if shape.has_chart:
                    chart_die = shape.chart
                    break

            if not chart_die:
                raise ValueError("No se encontró ninguna gráfica en la cuarta diapositiva (Diesel).")

            snr_col_die = None
            for col in self.df_data_diesel.columns:
                if "SNR" in str(col).upper():
                    snr_col_die = col
                    break

            if not snr_col_die:
                raise ValueError("No se encontró la columna 'SNR' en la tabla de Diesel.")

            categories_die = []
            proceso_vals_die = []
            diario_vals_die = []
            programa_vals_die = []
            columna1_vals_die = []

            # Filtrar df_prod_diesel para quedarnos con los años y meses activos
            prod_rows_die = []
            for idx, row in self.df_prod_diesel.iterrows():
                cat = str(row.iloc[0]).strip()
                val = row.iloc[1]
                
                if not cat:
                    continue
                    
                if not any(c.isalpha() for c in cat):
                    prod_rows_die.append((cat, val))
                else:
                    try:
                        p_val = float(val)
                    except:
                        p_val = 0
                    if p_val != 0:
                        prod_rows_die.append((cat, val))

            # Ajustar al límite de 17 categorías
            if len(prod_rows_die) > 17:
                prod_rows_die = prod_rows_die[-17:]

            # Llenar filas de años y meses (sin celdas vacías al final)
            for i in range(len(prod_rows_die)):
                cat_val = prod_rows_die[i][0]
                try:
                    proc_val = float(prod_rows_die[i][1])
                except:
                    proc_val = None
                        
                categories_die.append(cat_val)
                proceso_vals_die.append(proc_val)
                diario_vals_die.append(None)
                programa_vals_die.append(None)
                columna1_vals_die.append(None)

            # Llenar filas diarias
            num_dias_reales_die = 31
            for i in range(30, -1, -1):
                val_snr = 0
                if i < len(self.df_data_diesel):
                    try:
                        raw_val = self.df_data_diesel[snr_col_die].iloc[i]
                        if raw_val != "":
                            val_snr = float(raw_val)
                        else:
                            val_snr = 0
                    except:
                        val_snr = 0
                if val_snr != 0:
                    num_dias_reales_die = i + 1
                    break

            for i in range(31):
                categories_die.append(str(i + 1))
                proceso_vals_die.append(None)
                
                if i >= num_dias_reales_die:
                    diario_vals_die.append(None)
                    programa_vals_die.append(None)
                    columna1_vals_die.append(None)
                    continue

                d_val = None
                if i < len(self.df_data_diesel):
                    try:
                        d_val = float(self.df_data_diesel[snr_col_die].iloc[i])
                    except:
                        d_val = None
                diario_vals_die.append(d_val)
                
                p_val = None
                if i < len(self.df_snr_diesel):
                    try:
                        p_val = float(self.df_snr_diesel.iloc[i, 0])
                    except:
                        p_val = None
                programa_vals_die.append(p_val)
                
                c_val = None
                if i < len(self.df_snr_diesel):
                    try:
                        c_val = float(self.df_snr_diesel.iloc[i, 1])
                    except:
                        c_val = None
                columna1_vals_die.append(c_val)

            # Actualizar la gráfica de Diesel (Diapositiva 4)
            self.update_slide_chart(chart_die, categories_die, proceso_vals_die, diario_vals_die, programa_vals_die, columna1_vals_die, wine_color, green_color)


            # --- 4. PROCESAR DIAPOSITIVA DE TURBOSINA (DIAPOSITIVA 5) ---
            slide_turb = prs.slides[4]
            chart_turb = None
            for shape in slide_turb.shapes:
                if shape.has_chart:
                    chart_turb = shape.chart
                    break

            if not chart_turb:
                raise ValueError("No se encontró ninguna gráfica en la quinta diapositiva (Turbosina).")

            snr_col_turb = None
            for col in self.df_data_turbosina.columns:
                if "SNR" in str(col).upper():
                    snr_col_turb = col
                    break

            if not snr_col_turb:
                raise ValueError("No se encontró la columna 'SNR' en la tabla de Turbosina.")

            categories_turb = []
            proceso_vals_turb = []
            diario_vals_turb = []
            programa_vals_turb = []
            columna1_vals_turb = []

            # Filtrar df_prod_turbosina para quedarnos con los años y meses activos
            prod_rows_turb = []
            for idx, row in self.df_prod_turbosina.iterrows():
                cat = str(row.iloc[0]).strip()
                val = row.iloc[1]
                
                if not cat:
                    continue
                    
                if not any(c.isalpha() for c in cat):
                    prod_rows_turb.append((cat, val))
                else:
                    try:
                        p_val = float(val)
                    except:
                        p_val = 0
                    if p_val != 0:
                        prod_rows_turb.append((cat, val))

            # Ajustar al límite de 17 categorías
            if len(prod_rows_turb) > 17:
                prod_rows_turb = prod_rows_turb[-17:]

            # Llenar filas de años y meses (sin celdas vacías al final)
            for i in range(len(prod_rows_turb)):
                cat_val = prod_rows_turb[i][0]
                try:
                    proc_val = float(prod_rows_turb[i][1])
                except:
                    proc_val = None
                        
                categories_turb.append(cat_val)
                proceso_vals_turb.append(proc_val)
                diario_vals_turb.append(None)
                programa_vals_turb.append(None)
                columna1_vals_turb.append(None)

            # Llenar filas diarias
            num_dias_reales_turb = 31
            for i in range(30, -1, -1):
                val_snr = 0
                if i < len(self.df_data_turbosina):
                    try:
                        raw_val = self.df_data_turbosina[snr_col_turb].iloc[i]
                        if raw_val != "":
                            val_snr = float(raw_val)
                        else:
                            val_snr = 0
                    except:
                        val_snr = 0
                if val_snr != 0:
                    num_dias_reales_turb = i + 1
                    break

            for i in range(31):
                categories_turb.append(str(i + 1))
                proceso_vals_turb.append(None)
                
                if i >= num_dias_reales_turb:
                    diario_vals_turb.append(None)
                    programa_vals_turb.append(None)
                    columna1_vals_turb.append(None)
                    continue

                d_val = None
                if i < len(self.df_data_turbosina):
                    try:
                        d_val = float(self.df_data_turbosina[snr_col_turb].iloc[i])
                    except:
                        d_val = None
                diario_vals_turb.append(d_val)
                
                p_val = None
                if i < len(self.df_snr_turbosina):
                    try:
                        p_val = float(self.df_snr_turbosina.iloc[i, 0])
                    except:
                        p_val = None
                programa_vals_turb.append(p_val)
                
                c_val = None
                if i < len(self.df_snr_turbosina):
                    try:
                        c_val = float(self.df_snr_turbosina.iloc[i, 1])
                    except:
                        c_val = None
                columna1_vals_turb.append(c_val)

            # Actualizar la gráfica de Turbosina (Diapositiva 5)
            self.update_slide_chart(chart_turb, categories_turb, proceso_vals_turb, diario_vals_turb, programa_vals_turb, columna1_vals_turb, wine_color, green_color)


            prs.save(save_path)

            self.after(0, self.on_pptx_success, save_path)

        except Exception as e:
            err_details = traceback.format_exc()
            self.after(0, self.on_pptx_error, str(e), err_details)

    def on_pptx_success(self, save_path):
        self.set_loading_state(False)
        self.lbl_file.configure(text="Inyección a PowerPoint lista")
        messagebox.showinfo("Éxito", f"¡Datos de la gráfica actualizados con éxito!\n\nGuardado en:\n{save_path}")

    def on_pptx_error(self, err_msg, error_details):
        self.set_loading_state(False)
        self.lbl_file.configure(text="Error PowerPoint")
        
        try:
            print(colored(f"\n[!] ERROR EN HILO DE POWERPOINT:\n{err_msg}\n\nTRACEBACK:\n{error_details}", "red"))
        except:
            print(f"\n[!] ERROR EN HILO DE POWERPOINT:\n{err_msg}\n\nTRACEBACK:\n{error_details}")
            
        messagebox.showerror("Error", f"Ocurrió un error al procesar la presentación:\n{err_msg}")

if __name__ == "__main__":
    import signal
    import sys

    def sigint_handler(sig, frame):
        print("\nEjecución cancelada por el usuario (Ctrl+C). Saliendo...")
        sys.exit(0)

    signal.signal(signal.SIGINT, sigint_handler)

    app = ExcelViewerApp()

    def check_signals():
        app.after(500, check_signals)
        
    app.after(500, check_signals)
    app.mainloop()
