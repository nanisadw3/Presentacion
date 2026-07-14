import os
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox
from CTkTable import CTkTable
import threading
import traceback
from termcolor import colored
import db_helper

class ExcelViewerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Visor de Excel (Nativo de CustomTkinter)")
        
        # Center main window
        width, height = 1500, 850
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = int((screen_w/2) - (width/2))
        y = int((screen_h/2) - (height/2))
        self.geometry(f"{width}x{height}+{x}+{y}")
        
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

        # Botón para agregar años manualmente a la BD
        self.btn_add_year = ctk.CTkButton(self.top_frame, 
                                          text="Agregar Año Extra", 
                                          font=("Roboto", 14, "bold"),
                                          command=self.open_add_year_dialog,
                                          height=40,
                                          fg_color="#007bff", hover_color="#0056b3")
        self.btn_add_year.pack(pady=15, padx=10, side="left")

        # ComboBox para alternar visualización de procesos
        self.lbl_proceso = ctk.CTkLabel(self.top_frame, text="Proceso:", font=("Roboto", 14, "bold"))
        self.lbl_proceso.pack(pady=15, padx=(20, 5), side="left")
        
        self.cb_proceso = ctk.CTkComboBox(self.top_frame, 
                                            values=["Crudo", "Gasolinas", "Diesel", "Turbosina", "Asfalto", "Combustoleo", "Cadereyta -Crudo", "Cadereyta -Gasolinas", "Cadereyta -Diesel", "Cadereyta -Combustoleo", "Madero -Crudo", "Madero -Gasolinas", "Madero -Diesel", "Madero -Turbosina", "Madero -Combustoleo", "Minatitlan -Crudo", "Minatitlan -Gasolinas", "Minatitlan -Diesel", "Minatitlan -Combustoleo"],
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
 
        # Datos de Cadereyta -Gasolinas
        self.df_data_cad_gas = None
        self.df_snr_cad_gas = None
        self.df_prod_cad_gas = None
        self.df_sim_cad_gas = None

        # Datos de Cadereyta -Diesel
        self.df_data_cad_die = None
        self.df_snr_cad_die = None
        self.df_prod_cad_die = None
        self.df_sim_cad_die = None
        self.df_data_cad_comb = None
        self.df_snr_cad_comb = None
        self.df_prod_cad_comb = None
        self.df_sim_cad_comb = None
        # Datos de Cadereyta -Crudo
        self.df_data_cad = None
        self.df_snr_cad = None
        self.df_prod_cad = None
        self.df_sim_cad = None

        # Datos de Madero -Crudo
        self.df_data_mad_crud = None
        self.df_snr_mad_crud = None
        self.df_prod_mad_crud = None
        self.df_sim_mad_crud = None

        # Datos de Madero -Gasolinas
        self.df_data_mad_gas = None
        self.df_snr_mad_gas = None
        self.df_prod_mad_gas = None
        self.df_sim_mad_gas = None

        # Datos de Madero -Diesel
        self.df_data_mad_die = None
        self.df_snr_mad_die = None
        self.df_prod_mad_die = None
        self.df_sim_mad_die = None

        # Datos de Madero -Turbosina
        self.df_data_mad_turb = None
        self.df_snr_mad_turb = None
        self.df_prod_mad_turb = None
        self.df_sim_mad_turb = None

        # Datos de Madero -Combustoleo
        self.df_data_mad_comb = None
        self.df_snr_mad_comb = None
        self.df_prod_mad_comb = None
        self.df_sim_mad_comb = None

        # Datos de Minatitlan -Crudo
        self.df_data_mina_crud = None
        self.df_snr_mina_crud = None
        self.df_prod_mina_crud = None
        self.df_sim_mina_crud = None

        # Datos de Minatitlan -Gasolinas
        self.df_data_mina_gas = None
        self.df_snr_mina_gas = None
        self.df_prod_mina_gas = None
        self.df_sim_mina_gas = None

        # Datos de Minatitlan -Diesel
        self.df_data_mina_die = None
        self.df_snr_mina_die = None
        self.df_prod_mina_die = None
        self.df_sim_mina_die = None

        # Datos de Minatitlan -Combustoleo
        self.df_data_mina_comb = None
        self.df_snr_mina_comb = None
        self.df_prod_mina_comb = None
        self.df_sim_mina_comb = None
 
        # Datos de Turbosina
        self.df_data_turbosina = None

        self.df_snr_turbosina = None
        self.df_prod_turbosina = None
        self.df_sim_turbosina = None

        # Definir directorio por defecto (con fallback si no existe el de descargas)
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        wsl_path = "/mnt/c/Users/10900096799/Downloads"
        if os.path.exists(downloads_path):
            self.default_dir = downloads_path
        elif os.path.exists(wsl_path):
            self.default_dir = wsl_path
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

    def open_add_year_dialog(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Agregar/Editar Producción")
        
        width, height = 450, 500
        screen_w = dialog.winfo_screenwidth()
        screen_h = dialog.winfo_screenheight()
        x = int((screen_w/2) - (width/2))
        y = int((screen_h/2) - (height/2))
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        dialog.transient(self)
        dialog.grab_set()

        lbl_proceso = ctk.CTkLabel(dialog, text="Proceso:", font=("Roboto", 14, "bold"))
        lbl_proceso.pack(pady=(10, 0))
        opciones_procesos = ["Crudo", "Gasolinas", "Diesel", "Turbosina", "Asfalto", "Combustoleo", "Cadereyta -Crudo", "Cadereyta -Gasolinas", "Cadereyta -Diesel", "Cadereyta -Combustoleo", "Madero -Crudo", "Madero -Gasolinas", "Madero -Diesel", "Madero -Turbosina", "Madero -Combustoleo", "Minatitlan -Crudo", "Minatitlan -Gasolinas", "Minatitlan -Diesel", "Minatitlan -Combustoleo"]
        combo_proceso = ctk.CTkComboBox(dialog, values=opciones_procesos, width=250)
        combo_proceso.pack(pady=(5, 10))

        tipo_var = ctk.StringVar(value="mes")
        frame_radios = ctk.CTkFrame(dialog, fg_color="transparent")
        frame_radios.pack(pady=10)
        
        radio_mes = ctk.CTkRadioButton(frame_radios, text="Agregar un Mes", variable=tipo_var, value="mes")
        radio_mes.pack(side="left", padx=10)
        radio_anio = ctk.CTkRadioButton(frame_radios, text="Agregar un Año (Total)", variable=tipo_var, value="anio")
        radio_anio.pack(side="left", padx=10)

        lbl_anio = ctk.CTkLabel(dialog, text="Año (ej. 2024):", font=("Roboto", 14, "bold"))
        lbl_anio.pack(pady=(10, 0))
        entry_anio = ctk.CTkEntry(dialog, width=250)
        entry_anio.pack(pady=(5, 10))

        frame_mes_input = ctk.CTkFrame(dialog, fg_color="transparent")
        frame_mes_input.pack(pady=5)
        lbl_mes = ctk.CTkLabel(frame_mes_input, text="Mes:", font=("Roboto", 14, "bold"))
        lbl_mes.pack(side="left", padx=5)
        meses_opciones = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
        combo_mes = ctk.CTkComboBox(frame_mes_input, values=meses_opciones, width=150)
        combo_mes.pack(side="left", padx=5)

        def on_radio_change(*args):
            if tipo_var.get() == "anio":
                combo_mes.configure(state="disabled")
            else:
                combo_mes.configure(state="normal")
        tipo_var.trace_add("write", on_radio_change)

        lbl_prod = ctk.CTkLabel(dialog, text="Producción:", font=("Roboto", 14, "bold"))
        lbl_prod.pack(pady=(10, 0))
        entry_prod = ctk.CTkEntry(dialog, width=250)
        entry_prod.pack(pady=(5, 20))

        def on_save():
            proceso = combo_proceso.get()
            anio = entry_anio.get().strip()
            if not anio.isdigit():
                messagebox.showerror("Error", "El año debe ser un número válido.")
                return
            
            try:
                prod = float(entry_prod.get().strip())
            except:
                messagebox.showerror("Error", "La producción debe ser numérica.")
                return
            
            import db_helper
            mes_val = combo_mes.get() if tipo_var.get() == "mes" else "AÑO"
            
            db_helper.save_extra_prod(proceso, anio, mes_val, prod)
            
            tipo_txt = "mes de " + mes_val if mes_val != "AÑO" else "año completo"
            messagebox.showinfo("Éxito", f"Producción del {tipo_txt} para '{proceso}' en {anio} guardada correctamente.")
            dialog.destroy()
            
            if hasattr(self, 'current_file_path') and self.current_file_path:
                self.set_loading_state(True, "Recargando datos del Excel...")
                threading.Thread(target=self.async_load_data, args=(self.current_file_path,), daemon=True).start()

        def on_clear():
            if messagebox.askyesno("Confirmar", "¿Estás seguro de que quieres borrar TODOS los años y meses extra de la Base de Datos? (Esto no afecta al Excel original)"):
                import db_helper
                db_helper.clear_db()
                messagebox.showinfo("Limpieza Completa", "Se ha borrado toda la información extra de la base de datos local.")
                dialog.destroy()
                
                if hasattr(self, 'current_file_path') and self.current_file_path:
                    self.set_loading_state(True, "Recargando datos del Excel...")
                    threading.Thread(target=self.async_load_data, args=(self.current_file_path,), daemon=True).start()

        btn_save = ctk.CTkButton(dialog, text="Guardar", command=on_save, fg_color="#28a745", hover_color="#218838")
        btn_save.pack(pady=(15, 5))

        btn_clear = ctk.CTkButton(dialog, text="Limpiar Base de Datos", command=on_clear, fg_color="#dc3545", hover_color="#c82333")
        btn_clear.pack(pady=5)

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
            import math
            try:
                conn = sqlite3.connect(db_path)

                # Redondear todos los valores numéricos a enteros para la BD
                # (la pantalla conserva los decimales de MCP y Producción).
                def to_int_val(v):
                    try:
                        if v == "" or v is None:
                            return v
                        if isinstance(v, str):
                            return v
                        f = float(v)
                        if math.isinf(f) or math.isnan(f):
                            return ""
                        return int(round(f))
                    except (ValueError, TypeError):
                        return v

                def redondear_df(df):
                    out = df.copy()
                    
                    # Renombrar columnas vacías o duplicadas para que SQLite no arroje error
                    nuevas_columnas = []
                    for i, c in enumerate(out.columns):
                        if str(c).strip() == "" or pd.isna(c):
                            nuevas_columnas.append(f"Col_Vacia_{i+1}")
                        else:
                            nuevas_columnas.append(str(c))
                            
                    vistos = set()
                    for i in range(len(nuevas_columnas)):
                        orig = nuevas_columnas[i]
                        if orig in vistos:
                            j = 1
                            while f"{orig}_{j}" in vistos:
                                j += 1
                            nuevas_columnas[i] = f"{orig}_{j}"
                        vistos.add(nuevas_columnas[i])
                        
                    out.columns = nuevas_columnas

                    for c in out.columns:
                        out[c] = out[c].map(to_int_val)
                    return out

                def save_df(df, table_name):
                    if df is not None and len(df.columns) > 0:
                        redondear_df(df).to_sql(table_name, conn, if_exists='replace', index=False)

                # Guardar datos de Crudo
                save_df(self.df_data, 'crudo_tabla_principal')
                save_df(self.df_snr, 'crudo_programa_snr')
                save_df(self.df_prod, 'crudo_produccion')
                save_df(self.df_sim, 'crudo_simulacion_anual')

                # Guardar datos de Crudo Cadereyta
                save_df(getattr(self, 'df_data_cad', None), 'cadereyta_crudo_tabla_principal')
                save_df(getattr(self, 'df_snr_cad', None), 'cadereyta_crudo_programa_snr')
                save_df(getattr(self, 'df_prod_cad', None), 'cadereyta_crudo_produccion')
                save_df(getattr(self, 'df_sim_cad', None), 'cadereyta_crudo_simulacion_anual')

                 # Guardar datos de Gasolinas
                save_df(self.df_data_gasolinas, 'gasolinas_tabla_principal')
                save_df(self.df_snr_gasolinas, 'gasolinas_programa_snr')
                save_df(self.df_prod_gasolinas, 'gasolinas_produccion')
                save_df(getattr(self, 'df_sim_gasolinas', None), 'gasolinas_simulacion_anual')

                # Guardar datos de Diesel
                save_df(self.df_data_diesel, 'diesel_tabla_principal')
                save_df(self.df_snr_diesel, 'diesel_programa_snr')
                save_df(self.df_prod_diesel, 'diesel_produccion')
                save_df(getattr(self, 'df_sim_diesel', None), 'diesel_simulacion_anual')

                # Guardar datos de Turbosina
                save_df(self.df_data_turbosina, 'turbosina_tabla_principal')
                save_df(self.df_snr_turbosina, 'turbosina_programa_snr')
                save_df(self.df_prod_turbosina, 'turbosina_produccion')
                save_df(getattr(self, 'df_sim_turbosina', None), 'turbosina_simulacion_anual')

                # Guardar datos de Asfalto
                save_df(getattr(self, 'df_data_asfalto', None), 'asfalto_tabla_principal')
                save_df(getattr(self, 'df_snr_asfalto', None), 'asfalto_programa_snr')
                save_df(getattr(self, 'df_prod_asfalto', None), 'asfalto_produccion')
                save_df(getattr(self, 'df_sim_asfalto', None), 'asfalto_simulacion_anual')

                # Guardar datos de Combustoleo
                save_df(getattr(self, 'df_data_combustoleo', None), 'combustoleo_tabla_principal')
                save_df(getattr(self, 'df_snr_combustoleo', None), 'combustoleo_programa_snr')
                save_df(getattr(self, 'df_prod_combustoleo', None), 'combustoleo_produccion')
                save_df(getattr(self, 'df_sim_combustoleo', None), 'combustoleo_simulacion_anual')

                # Guardar datos de Cadereyta - Gasolinas
                save_df(getattr(self, 'df_data_cad_gas', None), 'cadereyta_gasolinas_tabla_principal')
                save_df(getattr(self, 'df_snr_cad_gas', None), 'cadereyta_gasolinas_programa_snr')
                save_df(getattr(self, 'df_prod_cad_gas', None), 'cadereyta_gasolinas_produccion')
                save_df(getattr(self, 'df_sim_cad_gas', None), 'cadereyta_gasolinas_simulacion_anual')

                # Guardar datos de Cadereyta - Diesel
                save_df(getattr(self, 'df_data_cad_die', None), 'cadereyta_diesel_tabla_principal')
                save_df(getattr(self, 'df_snr_cad_die', None), 'cadereyta_diesel_programa_snr')
                save_df(getattr(self, 'df_prod_cad_die', None), 'cadereyta_diesel_produccion')
                save_df(getattr(self, 'df_sim_cad_die', None), 'cadereyta_diesel_simulacion_anual')
                save_df(getattr(self, 'df_data_cad_comb', None), 'cadereyta_combustoleo_tabla_principal')
                save_df(getattr(self, 'df_snr_cad_comb', None), 'cadereyta_combustoleo_programa_snr')
                save_df(getattr(self, 'df_prod_cad_comb', None), 'cadereyta_combustoleo_produccion')
                save_df(getattr(self, 'df_sim_cad_comb', None), 'cadereyta_combustoleo_simulacion_anual')

                # Guardar datos de Madero - Crudo
                save_df(getattr(self, 'df_data_mad_crud', None), 'madero_crudo_tabla_principal')
                save_df(getattr(self, 'df_snr_mad_crud', None), 'madero_crudo_programa_snr')
                save_df(getattr(self, 'df_prod_mad_crud', None), 'madero_crudo_produccion')
                save_df(getattr(self, 'df_sim_mad_crud', None), 'madero_crudo_simulacion_anual')

                # Guardar datos de Madero - Gasolinas
                save_df(getattr(self, 'df_data_mad_gas', None), 'madero_gasolinas_tabla_principal')
                save_df(getattr(self, 'df_snr_mad_gas', None), 'madero_gasolinas_programa_snr')
                save_df(getattr(self, 'df_prod_mad_gas', None), 'madero_gasolinas_produccion')
                save_df(getattr(self, 'df_sim_mad_gas', None), 'madero_gasolinas_simulacion_anual')

                # Guardar datos de Madero - Diesel
                save_df(getattr(self, 'df_data_mad_die', None), 'madero_diesel_tabla_principal')
                save_df(getattr(self, 'df_snr_mad_die', None), 'madero_diesel_programa_snr')
                save_df(getattr(self, 'df_prod_mad_die', None), 'madero_diesel_produccion')
                save_df(getattr(self, 'df_sim_mad_die', None), 'madero_diesel_simulacion_anual')

                # Guardar datos de Madero - Turbosina
                save_df(getattr(self, 'df_data_mad_turb', None), 'madero_turbosina_tabla_principal')
                save_df(getattr(self, 'df_snr_mad_turb', None), 'madero_turbosina_programa_snr')
                save_df(getattr(self, 'df_prod_mad_turb', None), 'madero_turbosina_produccion')
                save_df(getattr(self, 'df_sim_mad_turb', None), 'madero_turbosina_simulacion_anual')

                # Guardar datos de Madero - Combustoleo
                save_df(getattr(self, 'df_data_mad_comb', None), 'madero_combustoleo_tabla_principal')
                save_df(getattr(self, 'df_snr_mad_comb', None), 'madero_combustoleo_programa_snr')
                save_df(getattr(self, 'df_prod_mad_comb', None), 'madero_combustoleo_produccion')
                save_df(getattr(self, 'df_sim_mad_comb', None), 'madero_combustoleo_simulacion_anual')

                # Guardar datos de Minatitlan - Crudo
                save_df(getattr(self, 'df_data_mina_crud', None), 'minatitlan_crudo_tabla_principal')
                save_df(getattr(self, 'df_snr_mina_crud', None), 'minatitlan_crudo_programa_snr')
                save_df(getattr(self, 'df_prod_mina_crud', None), 'minatitlan_crudo_produccion')
                save_df(getattr(self, 'df_sim_mina_crud', None), 'minatitlan_crudo_simulacion_anual')

                # Guardar datos de Minatitlan - Gasolinas
                save_df(getattr(self, 'df_data_mina_gas', None), 'minatitlan_gasolinas_tabla_principal')
                save_df(getattr(self, 'df_snr_mina_gas', None), 'minatitlan_gasolinas_programa_snr')
                save_df(getattr(self, 'df_prod_mina_gas', None), 'minatitlan_gasolinas_produccion')
                save_df(getattr(self, 'df_sim_mina_gas', None), 'minatitlan_gasolinas_simulacion_anual')

                # Guardar datos de Minatitlan - Diesel
                save_df(getattr(self, 'df_data_mina_die', None), 'minatitlan_diesel_tabla_principal')
                save_df(getattr(self, 'df_snr_mina_die', None), 'minatitlan_diesel_programa_snr')
                save_df(getattr(self, 'df_prod_mina_die', None), 'minatitlan_diesel_produccion')
                save_df(getattr(self, 'df_sim_mina_die', None), 'minatitlan_diesel_simulacion_anual')

                # Guardar datos de Minatitlan - Combustoleo
                save_df(getattr(self, 'df_data_mina_comb', None), 'minatitlan_combustoleo_tabla_principal')
                save_df(getattr(self, 'df_snr_mina_comb', None), 'minatitlan_combustoleo_programa_snr')
                save_df(getattr(self, 'df_prod_mina_comb', None), 'minatitlan_combustoleo_produccion')
                save_df(getattr(self, 'df_sim_mina_comb', None), 'minatitlan_combustoleo_simulacion_anual')

                conn.commit()
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
        import excel_parser
        excel_parser.load_data(self, file_path)
        
    def on_load_success(self, file_path, df_data, df_snr, df_prod, df_sim,
                            df_data_gasolinas=None, df_snr_gasolinas=None, df_prod_gasolinas=None, df_sim_gasolinas=None,
                            df_data_diesel=None, df_snr_diesel=None, df_prod_diesel=None, df_sim_diesel=None,
                            df_data_turbosina=None, df_snr_turbosina=None, df_prod_turbosina=None, df_sim_turbosina=None,
                            df_data_asfalto=None, df_snr_asfalto=None, df_prod_asfalto=None, df_sim_asfalto=None,
                            df_data_combustoleo=None, df_snr_combustoleo=None, df_prod_combustoleo=None, df_sim_combustoleo=None,
                            df_data_cad_gas=None, df_snr_cad_gas=None, df_prod_cad_gas=None, df_sim_cad_gas=None,
                            df_data_cad_die=None, df_snr_cad_die=None, df_prod_cad_die=None, df_sim_cad_die=None,
                            df_data_cad=None, df_snr_cad=None, df_prod_cad=None, df_sim_cad=None,
                            df_data_cad_comb=None, df_snr_cad_comb=None, df_prod_cad_comb=None, df_sim_cad_comb=None,
                            df_data_mad_crud=None, df_snr_mad_crud=None, df_prod_mad_crud=None, df_sim_mad_crud=None,
                            df_data_mad_gas=None, df_snr_mad_gas=None, df_prod_mad_gas=None, df_sim_mad_gas=None,
                            df_data_mad_die=None, df_snr_mad_die=None, df_prod_mad_die=None, df_sim_mad_die=None,
                            df_data_mad_turb=None, df_snr_mad_turb=None, df_prod_mad_turb=None, df_sim_mad_turb=None,
                            df_data_mad_comb=None, df_snr_mad_comb=None, df_prod_mad_comb=None, df_sim_mad_comb=None,
                            df_data_mina_crud=None, df_snr_mina_crud=None, df_prod_mina_crud=None, df_sim_mina_crud=None,
                            df_data_mina_gas=None, df_snr_mina_gas=None, df_prod_mina_gas=None, df_sim_mina_gas=None,
                            df_data_mina_die=None, df_snr_mina_die=None, df_prod_mina_die=None, df_sim_mina_die=None,
                            df_data_mina_comb=None, df_snr_mina_comb=None, df_prod_mina_comb=None, df_sim_mina_comb=None):
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
 
        self.df_data_cad_gas = df_data_cad_gas
        self.df_snr_cad_gas = df_snr_cad_gas
        self.df_prod_cad_gas = df_prod_cad_gas
        self.df_sim_cad_gas = df_sim_cad_gas

        self.df_data_cad_die = df_data_cad_die
        self.df_snr_cad_die = df_snr_cad_die
        self.df_prod_cad_die = df_prod_cad_die
        self.df_sim_cad_die = df_sim_cad_die
 
        self.df_data_cad = df_data_cad
        self.df_snr_cad = df_snr_cad
        self.df_prod_cad = df_prod_cad
        self.df_sim_cad = df_sim_cad

        self.df_data_cad_comb = df_data_cad_comb
        self.df_snr_cad_comb = df_snr_cad_comb
        self.df_prod_cad_comb = df_prod_cad_comb
        self.df_sim_cad_comb = df_sim_cad_comb
 
        self.df_data_turbosina = df_data_turbosina

        self.df_snr_turbosina = df_snr_turbosina
        self.df_prod_turbosina = df_prod_turbosina
        self.df_sim_turbosina = df_sim_turbosina

        self.df_data_asfalto = df_data_asfalto
        self.df_snr_asfalto = df_snr_asfalto
        self.df_prod_asfalto = df_prod_asfalto
        self.df_sim_asfalto = df_sim_asfalto

        self.df_data_combustoleo = df_data_combustoleo
        self.df_snr_combustoleo = df_snr_combustoleo
        self.df_prod_combustoleo = df_prod_combustoleo
        self.df_sim_combustoleo = df_sim_combustoleo

        self.df_data_mad_crud = df_data_mad_crud
        self.df_snr_mad_crud = df_snr_mad_crud
        self.df_prod_mad_crud = df_prod_mad_crud
        self.df_sim_mad_crud = df_sim_mad_crud

        self.df_data_mad_gas = df_data_mad_gas
        self.df_snr_mad_gas = df_snr_mad_gas
        self.df_prod_mad_gas = df_prod_mad_gas
        self.df_sim_mad_gas = df_sim_mad_gas

        self.df_data_mad_die = df_data_mad_die
        self.df_snr_mad_die = df_snr_mad_die
        self.df_prod_mad_die = df_prod_mad_die
        self.df_sim_mad_die = df_sim_mad_die

        self.df_data_mad_turb = df_data_mad_turb
        self.df_snr_mad_turb = df_snr_mad_turb
        self.df_prod_mad_turb = df_prod_mad_turb
        self.df_sim_mad_turb = df_sim_mad_turb

        self.df_data_mad_comb = df_data_mad_comb
        self.df_snr_mad_comb = df_snr_mad_comb
        self.df_prod_mad_comb = df_prod_mad_comb
        self.df_sim_mad_comb = df_sim_mad_comb

        self.df_data_mina_crud = df_data_mina_crud
        self.df_snr_mina_crud = df_snr_mina_crud
        self.df_prod_mina_crud = df_prod_mina_crud
        self.df_sim_mina_crud = df_sim_mina_crud

        self.df_data_mina_gas = df_data_mina_gas
        self.df_snr_mina_gas = df_snr_mina_gas
        self.df_prod_mina_gas = df_prod_mina_gas
        self.df_sim_mina_gas = df_sim_mina_gas

        self.df_data_mina_die = df_data_mina_die
        self.df_snr_mina_die = df_snr_mina_die
        self.df_prod_mina_die = df_prod_mina_die
        self.df_sim_mina_die = df_sim_mina_die

        self.df_data_mina_comb = df_data_mina_comb
        self.df_snr_mina_comb = df_snr_mina_comb
        self.df_prod_mina_comb = df_prod_mina_comb
        self.df_sim_mina_comb = df_sim_mina_comb

        # Mostrar las tablas correspondientes a la selección actual del ComboBox
        self.show_proceso_tables(self.cb_proceso.get())

        self.set_loading_state(False)
        self.current_file_path = file_path
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
        if hasattr(self, 'lbl_table1') and self.lbl_table1 is not None:
            self.lbl_table1.destroy()
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
            lbl2_txt = "Programa (CMP, PODIM)"
            lbl3_txt = "Fecha y Producción (Año/Mes, Producción)"
        elif selection == "Cadereyta -Crudo":
            df_data = self.df_data_cad
            df_snr = self.df_snr_cad
            df_prod = self.df_prod_cad
            df_sim = self.df_sim_cad
            lbl2_txt = "Programa Crudo Cadereyta (Col BI x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción Crudo Cadereyta (AW-AX, Filas 21-40)"
        elif selection == "Cadereyta -Gasolinas":
            df_data = self.df_data_cad_gas
            df_snr = self.df_snr_cad_gas
            df_prod = self.df_prod_cad_gas
            df_sim = self.df_sim_cad_gas
            lbl2_txt = "Programa de Gasolinas (Col BS x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Gasolinas (BO-BP, Filas 21-40)"
        elif selection == "Cadereyta -Diesel":
            df_data = self.df_data_cad_die
            df_snr = self.df_snr_cad_die
            df_prod = self.df_prod_cad_die
            df_sim = self.df_sim_cad_die
            lbl2_txt = "Programa de Diesel (Col CC x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Diesel (CG-CH, Filas 21-40)"
        elif selection == "Cadereyta -Combustoleo":
            df_data = self.df_data_cad_comb
            df_snr = self.df_snr_cad_comb
            df_prod = self.df_prod_cad_comb
            df_sim = self.df_sim_cad_comb
            lbl2_txt = "Programa de Combustoleo (Col CU x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Combustoleo (AE-AF, Filas 159-189)"
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
        elif selection == "Turbosina":
            df_data = self.df_data_turbosina
            df_snr = self.df_snr_turbosina
            df_prod = self.df_prod_turbosina
            df_sim = self.df_sim_turbosina
            lbl2_txt = "Programa de Turbosina (AW-AX, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Turbosina (AM-AN, Filas 21-40)"
        elif selection == "Asfalto":
            df_data = self.df_data_asfalto
            df_snr = self.df_snr_asfalto
            df_prod = self.df_prod_asfalto
            df_sim = self.df_sim_asfalto
            lbl2_txt = "Programa de Asfalto (AK-AL, Filas 122-152)"
            lbl3_txt = "Fecha y Producción de Asfalto (AQ-AR, Filas 121-140)"
        elif selection == "Combustoleo":
            df_data = self.df_data_combustoleo
            df_snr = self.df_snr_combustoleo
            df_prod = self.df_prod_combustoleo
            df_sim = self.df_sim_combustoleo
            lbl2_txt = "Programa de Combustoleo (BC-BD, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Combustoleo (AQ-AR, Filas 21-40)"
        elif selection == "Madero -Crudo":
            df_data = self.df_data_mad_crud
            df_snr = self.df_snr_mad_crud
            df_prod = self.df_prod_mad_crud
            df_sim = self.df_sim_mad_crud
            lbl2_txt = "Programa de Crudo Madero (Col BJ x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Crudo Madero (AY-AZ, Filas 21-40)"
        elif selection == "Madero -Gasolinas":
            df_data = self.df_data_mad_gas
            df_snr = self.df_snr_mad_gas
            df_prod = self.df_prod_mad_gas
            df_sim = self.df_sim_mad_gas
            lbl2_txt = "Programa de Gasolinas Madero (Col BT x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Gasolinas Madero (BS-BT, Filas 21-40)"
        elif selection == "Madero -Diesel":
            df_data = self.df_data_mad_die
            df_snr = self.df_snr_mad_die
            df_prod = self.df_prod_mad_die
            df_sim = self.df_sim_mad_die
            lbl2_txt = "Programa de Diesel Madero (Col CD x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Diesel Madero (CI-CJ, Filas 21-40)"
        elif selection == "Madero -Turbosina":
            df_data = self.df_data_mad_turb
            df_snr = self.df_snr_mad_turb
            df_prod = self.df_prod_mad_turb
            df_sim = self.df_sim_mad_turb
            lbl2_txt = "Programa de Turbosina Madero (Col CM x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Turbosina Madero (CX-CY, Filas 21-40)"
        elif selection == "Madero -Combustoleo":
            df_data = self.df_data_mad_comb
            df_snr = self.df_snr_mad_comb
            df_prod = self.df_prod_mad_comb
            df_sim = self.df_sim_mad_comb
            lbl2_txt = "Programa de Combustoleo Madero (Col CV x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Combustoleo Madero (AK-AL, Filas 158-179)"
        elif selection == "Minatitlan -Crudo":
            df_data = self.df_data_mina_crud
            df_snr = self.df_snr_mina_crud
            df_prod = self.df_prod_mina_crud
            df_sim = self.df_sim_mina_crud
            lbl2_txt = "Programa de Crudo Minatitlan (Col BK x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Crudo Minatitlan (BC-BD, Filas 21-40)"
        elif selection == "Minatitlan -Gasolinas":
            df_data = self.df_data_mina_gas
            df_snr = self.df_snr_mina_gas
            df_prod = self.df_prod_mina_gas
            df_sim = self.df_sim_mina_gas
            lbl2_txt = "Programa de Gasolinas Minatitlan (Col BU x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Gasolinas Minatitlan (BU-BV, Filas 21-40)"
        elif selection == "Minatitlan -Diesel":
            df_data = self.df_data_mina_die
            df_snr = self.df_snr_mina_die
            df_prod = self.df_prod_mina_die
            df_sim = self.df_sim_mina_die
            lbl2_txt = "Programa de Diesel Minatitlan (Col CE x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Diesel Minatitlan (CM-CN, Filas 21-40)"
        elif selection == "Minatitlan -Combustoleo":
            df_data = self.df_data_mina_comb
            df_snr = self.df_snr_mina_comb
            df_prod = self.df_prod_mina_comb
            df_sim = self.df_sim_mina_comb
            lbl2_txt = "Programa de Combustoleo Minatitlan (Col CW x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Combustoleo Minatitlan (AN-AO, Filas 158-179)"

        if df_data is None or df_snr is None or df_prod is None:
            return

        # Mostrar valores redondeados a enteros en la vista previa.
        # (los cálculos de simulación siguen usando los datos originales sin redondear)
        def _red(val):
            if val == "" or val is None:
                return val
            try:
                f = float(val)
                import math
                if math.isinf(f) or math.isnan(f):
                    return ""
                return int(round(f))
            except (ValueError, TypeError):
                return val

        def _red_df(d):
            out = d.copy()
            for c in out.columns:
                out[c] = out[c].map(_red)
            return out

        df_data_v = _red_df(df_data)
        df_snr_v = _red_df(df_snr)
        df_prod_v = _red_df(df_prod)

        # Dibujar Tabla 1
        self.lbl_table1 = ctk.CTkLabel(self.scroll_frame, text=selection, font=("Roboto", 16, "bold"), text_color="#3484F0")
        self.lbl_table1.pack(pady=(20, 5))
        headers = list(df_data_v.columns)
        rows = df_data_v.to_numpy().tolist()
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
        self.lbl_table2 = ctk.CTkLabel(self.scroll_frame, text=f"Programa de {selection}", font=("Roboto", 16, "bold"), text_color="#3484F0")
        self.lbl_table2.pack(pady=(20, 5))
        headers2 = ["MCP", "PODIM"]
        rows2 = df_snr_v.to_numpy().tolist()
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
        rows3 = df_prod_v.to_numpy().tolist()
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
        if df_sim is not None:
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

    def async_send_to_pptx(self, file_path, save_path):
        import pptx_exporter
        pptx_exporter.export_to_pptx(self, file_path, save_path)
        
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
