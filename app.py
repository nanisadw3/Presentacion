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

    def report_callback_exception(self, exc, val, tb):
        import traceback
        from tkinter import messagebox
        err_details = "".join(traceback.format_exception(exc, val, tb))
        print(f"Error detectado en callback:\n{err_details}")
        messagebox.showerror("Error Inesperado", f"Ocurrió un error inesperado en la aplicación:\n\n{val}\n\nDetalles:\n{err_details}")

        self.title("Sistema de Proyección y Reportes de Refinerías")
        
        # Centrar ventana al 80% de la pantalla principal
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = int(screen_w * 0.8)
        height = int(screen_h * 0.8)
        x = int((screen_w - width) / 2)
        y = int((screen_h - height) / 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # Configuración del tema
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        # ═══ Contenedor principal ═══
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=15, pady=(10, 5))

        # ═══ FILA 1: Archivo + Estado ═══
        self.row1_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        self.row1_frame.pack(fill="x", pady=(0, 8))

        self.btn_buscar = ctk.CTkButton(self.row1_frame, 
                                        text="📂  Cargar Excel", 
                                        font=("Roboto", 13, "bold"),
                                        command=self.load_excel,
                                        height=36, width=150,
                                        corner_radius=8)
        self.btn_buscar.pack(pady=10, padx=(15, 8), side="left")

        self.lbl_file = ctk.CTkLabel(self.row1_frame, 
                                     text="Sin archivo cargado",
                                     font=("Roboto", 12),
                                     text_color="#888888")
        self.lbl_file.pack(pady=10, padx=10, side="left", fill="x", expand=True)

        # Separador visual
        sep1 = ctk.CTkFrame(self.row1_frame, width=2, height=26, fg_color="#444444")
        sep1.pack(side="left", padx=8, pady=10)

        # Selector de proceso
        self.lbl_proceso = ctk.CTkLabel(self.row1_frame, text="Proceso:", font=("Roboto", 12, "bold"))
        self.lbl_proceso.pack(pady=10, padx=(8, 4), side="left")
        
        self.cb_proceso = ctk.CTkComboBox(self.row1_frame, 
                                            values=["Crudo", "Gasolinas", "Diesel", "Turbosina", "Asfalto", "Combustoleo",
                                                    "Cadereyta -Crudo", "Cadereyta -Gasolinas", "Cadereyta -Diesel", "Cadereyta -Combustoleo",
                                                    "Madero -Crudo", "Madero -Gasolinas", "Madero -Diesel", "Madero -Turbosina", "Madero -Combustoleo",
                                                    "Minatitlan -Crudo", "Minatitlan -Gasolinas", "Minatitlan -Diesel", "Minatitlan -Combustoleo",
                                                    "Salamanca -Crudo", "Salamanca -Gasolinas", "Salamanca -Diesel", "Salamanca -Turbosina", "Salamanca -Combustoleo",
                                                    "Salina Cruz -Crudo", "Salina Cruz -Gasolinas", "Salina Cruz -Diesel", "Salina Cruz -Turbosina", "Salina Cruz -Combustoleo",
                                                    "Tula -Crudo", "Tula -Gasolinas", "Tula -Diesel", "Tula -Turbosina", "Tula -Combustoleo",
                                                    "Olmeca -Crudo", "Olmeca -Gasolinas", "Olmeca -Diesel"],
                                            font=("Roboto", 12),
                                            command=self.on_proceso_changed,
                                            state="readonly",
                                            width=200,
                                            corner_radius=8)
        self.cb_proceso.pack(pady=10, padx=(4, 15), side="left")
        self.cb_proceso.set("Crudo")

        # ═══ FILA 2: Herramientas ═══
        self.row2_frame = ctk.CTkFrame(self.main_frame, corner_radius=10, fg_color="#1a1a2e")
        self.row2_frame.pack(fill="x", pady=(0, 10))

        lbl_tools = ctk.CTkLabel(self.row2_frame, text="Herramientas:", font=("Roboto", 11, "bold"), text_color="#777777")
        lbl_tools.pack(pady=8, padx=15, side="left")

        self.btn_powerpoint = ctk.CTkButton(self.row2_frame, 
                                            text="📊  Exportar a PowerPoint", 
                                            font=("Roboto", 12, "bold"),
                                            command=self.send_to_powerpoint,
                                            height=32, width=195,
                                            corner_radius=8,
                                            fg_color="#8b5cf6", hover_color="#7c3aed")
        self.btn_powerpoint.pack(pady=8, padx=6, side="left")

        self.btn_guardar = ctk.CTkButton(self.row2_frame, 
                                        text="💾  Guardar en BD", 
                                        font=("Roboto", 12, "bold"),
                                        command=self.save_to_database,
                                        height=32, width=160,
                                        corner_radius=8,
                                        fg_color="#28a745", hover_color="#218838")
        self.btn_guardar.pack(pady=8, padx=6, side="left")

        self.btn_add_year = ctk.CTkButton(self.row2_frame, 
                                          text="📅  Agregar Año Extra", 
                                          font=("Roboto", 12, "bold"),
                                          command=self.open_add_year_dialog,
                                          height=32, width=175,
                                          corner_radius=8,
                                          fg_color="#0d6efd", hover_color="#0b5ed7")
        self.btn_add_year.pack(pady=8, padx=6, side="left")

        self.btn_config_coords = ctk.CTkButton(self.row2_frame, 
                                              text="⚙  Coordenadas Excel", 
                                              font=("Roboto", 12, "bold"),
                                              command=self.open_config_coords_dialog,
                                              height=32, width=185,
                                              corner_radius=8,
                                              fg_color="#6c757d", hover_color="#5a6268")
        self.btn_config_coords.pack(pady=8, padx=6, side="left")


        # Scrollable Frame para contener la tabla
        self.scroll_frame = ctk.CTkScrollableFrame(self.main_frame, corner_radius=10)
        self.scroll_frame.pack(fill="both", expand=True)

        # Función para enlazar y propagar el scroll del mousewheel/trackpad en laptops y ratones
        import platform
        _is_macos = platform.system() == "Darwin"
        
        def _on_mousewheel(event):
            try:
                canvas = self.scroll_frame._parent_canvas
                if _is_macos:
                    # macOS: delta ya viene en la dirección correcta
                    # En algunos mouses/trackpads macOS, delta puede ser flotante o pequeño, por lo que usamos su signo
                    if event.delta > 0:
                        canvas.yview_scroll(-2, "units")
                    elif event.delta < 0:
                        canvas.yview_scroll(2, "units")
                else:
                    # Windows: delta suele ser +/- 120 o valores pequeños en trackpads de laptops
                    if event.delta > 0:
                        canvas.yview_scroll(-2, "units")
                    elif event.delta < 0:
                        canvas.yview_scroll(2, "units")
            except Exception:
                pass
        
        def _on_linux_scroll(event):
            try:
                canvas = self.scroll_frame._parent_canvas
                if event.num == 4:
                    canvas.yview_scroll(-2, "units")
                elif event.num == 5:
                    canvas.yview_scroll(2, "units")
            except Exception:
                pass

        # Vincular eventos de scroll globales para que funcionen sobre cualquier widget interno
        self.bind_all("<MouseWheel>", _on_mousewheel)
        # Linux usa Button-4/5 para scroll
        self.bind_all("<Button-4>", _on_linux_scroll)
        self.bind_all("<Button-5>", _on_linux_scroll)

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

        # Datos de Salamanca -Crudo
        self.df_data_sala_crud = None
        self.df_snr_sala_crud = None
        self.df_prod_sala_crud = None
        self.df_sim_sala_crud = None

        # Datos de Salamanca -Gasolinas
        self.df_data_sala_gas = None
        self.df_snr_sala_gas = None
        self.df_prod_sala_gas = None
        self.df_sim_sala_gas = None

        # Datos de Salamanca -Diesel
        self.df_data_sala_die = None
        self.df_snr_sala_die = None
        self.df_prod_sala_die = None
        self.df_sim_sala_die = None

        # Datos de Salamanca -Turbosina
        self.df_data_sala_turb = None
        self.df_snr_sala_turb = None
        self.df_prod_sala_turb = None
        self.df_sim_sala_turb = None

        # Datos de Salamanca -Combustoleo
        self.df_data_sala_comb = None
        self.df_snr_sala_comb = None
        self.df_prod_sala_comb = None
        self.df_sim_sala_comb = None

        # Datos de Salina Cruz -Crudo
        self.df_data_sal_crud = None
        self.df_snr_sal_crud = None
        self.df_prod_sal_crud = None
        self.df_sim_sal_crud = None

        # Datos de Salina Cruz -Gasolinas
        self.df_data_sal_gas = None
        self.df_snr_sal_gas = None
        self.df_prod_sal_gas = None
        self.df_sim_sal_gas = None

        # Datos de Salina Cruz -Diesel
        self.df_data_sal_die = None
        self.df_snr_sal_die = None
        self.df_prod_sal_die = None
        self.df_sim_sal_die = None

        # Datos de Salina Cruz -Turbosina
        self.df_data_sal_turb = None
        self.df_snr_sal_turb = None
        self.df_prod_sal_turb = None
        self.df_sim_sal_turb = None

        # Datos de Salina Cruz -Combustoleo
        self.df_data_sal_comb = None
        self.df_snr_sal_comb = None
        self.df_prod_sal_comb = None
        self.df_sim_sal_comb = None

        # Datos de Tula -Crudo
        self.df_data_tula_crud = None
        self.df_snr_tula_crud = None
        self.df_prod_tula_crud = None
        self.df_sim_tula_crud = None

        # Datos de Tula -Gasolinas
        self.df_data_tula_gas = None
        self.df_snr_tula_gas = None
        self.df_prod_tula_gas = None
        self.df_sim_tula_gas = None

        # Datos de Tula -Diesel
        self.df_data_tula_die = None
        self.df_snr_tula_die = None
        self.df_prod_tula_die = None
        self.df_sim_tula_die = None

        # Datos de Tula -Turbosina
        self.df_data_tula_turb = None
        self.df_snr_tula_turb = None
        self.df_prod_tula_turb = None
        self.df_sim_tula_turb = None

        # Datos de Tula -Combustoleo
        self.df_data_tula_comb = None
        self.df_snr_tula_comb = None
        self.df_prod_tula_comb = None
        self.df_sim_tula_comb = None

        # Datos de Olmeca -Crudo
        self.df_data_olme_crud = None
        self.df_snr_olme_crud = None
        self.df_prod_olme_crud = None
        self.df_sim_olme_crud = None

        # Datos de Olmeca -Gasolinas
        self.df_data_olme_gas = None
        self.df_snr_olme_gas = None
        self.df_prod_olme_gas = None
        self.df_sim_olme_gas = None

        # Datos de Olmeca -Diesel
        self.df_data_olme_die = None
        self.df_snr_olme_die = None
        self.df_prod_olme_die = None
        self.df_sim_olme_die = None
 
        # Datos de Turbosina
        self.df_data_turbosina = None

        self.df_snr_turbosina = None
        self.df_prod_turbosina = None
        self.df_sim_turbosina = None

        # Definir directorios por defecto con sus respectivas rutas preferidas y fallback
        path_excel_preferida = "/mnt/d/Datos_Perfil/400131/OneDrive - PETROLEOS MEXICANOS/Disco D/Margen 2026"
        path_pptx_preferida = "/mnt/d/Datos_Perfil/400131/Downloads"
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        app_dir = os.path.dirname(os.path.abspath(__file__))

        def check_and_get_dir(preferred_path):
            # 1. Probar ruta preferida tal cual (funciona en Linux/WSL/Docker)
            if os.path.exists(preferred_path):
                return preferred_path
            # 2. Si estamos en Windows, traducir formato /mnt/d/... a D:\...
            if os.name == 'nt' and preferred_path.startswith('/mnt/'):
                parts = preferred_path.split('/')
                if len(parts) >= 3:
                    drive = parts[2].upper() + ':'
                    remaining = '\\'.join(parts[3:])
                    win_path = drive + '\\' + remaining
                    if os.path.exists(win_path):
                        return win_path
            # 3. Fallback: carpeta de Downloads del usuario
            if os.path.exists(downloads_path):
                return downloads_path
            # 4. Fallback final: directorio donde corre la app
            return app_dir

        self.default_excel_dir = check_and_get_dir(path_excel_preferida)
        self.default_pptx_dir = check_and_get_dir(path_pptx_preferida)

        # Barra de progreso (inicialmente oculta)
        self.progress_bar = ctk.CTkProgressBar(self.row1_frame, width=200)
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
        opciones_procesos = ["Crudo", "Gasolinas", "Diesel", "Turbosina", "Asfalto", "Combustoleo", "Cadereyta -Crudo", "Cadereyta -Gasolinas", "Cadereyta -Diesel", "Cadereyta -Combustoleo", "Madero -Crudo", "Madero -Gasolinas", "Madero -Diesel", "Madero -Turbosina", "Madero -Combustoleo", "Minatitlan -Crudo", "Minatitlan -Gasolinas", "Minatitlan -Diesel", "Minatitlan -Combustoleo", "Salamanca -Crudo", "Salamanca -Gasolinas", "Salamanca -Diesel", "Salamanca -Turbosina", "Salamanca -Combustoleo", "Salina Cruz -Crudo", "Salina Cruz -Gasolinas", "Salina Cruz -Diesel", "Salina Cruz -Turbosina", "Salina Cruz -Combustoleo", "Tula -Crudo", "Tula -Gasolinas", "Tula -Diesel", "Tula -Turbosina", "Tula -Combustoleo", "Olmeca -Crudo", "Olmeca -Gasolinas", "Olmeca -Diesel"]
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
            initialdir=self.default_excel_dir,
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

                # Guardar datos de Salamanca - Crudo
                save_df(getattr(self, 'df_data_sala_crud', None), 'salamanca_crudo_tabla_principal')
                save_df(getattr(self, 'df_snr_sala_crud', None), 'salamanca_crudo_programa_snr')
                save_df(getattr(self, 'df_prod_sala_crud', None), 'salamanca_crudo_produccion')
                save_df(getattr(self, 'df_sim_sala_crud', None), 'salamanca_crudo_simulacion_anual')

                # Guardar datos de Salamanca - Gasolinas
                save_df(getattr(self, 'df_data_sala_gas', None), 'salamanca_gasolinas_tabla_principal')
                save_df(getattr(self, 'df_snr_sala_gas', None), 'salamanca_gasolinas_programa_snr')
                save_df(getattr(self, 'df_prod_sala_gas', None), 'salamanca_gasolinas_produccion')
                save_df(getattr(self, 'df_sim_sala_gas', None), 'salamanca_gasolinas_simulacion_anual')

                # Guardar datos de Salamanca - Diesel
                save_df(getattr(self, 'df_data_sala_die', None), 'salamanca_diesel_tabla_principal')
                save_df(getattr(self, 'df_snr_sala_die', None), 'salamanca_diesel_programa_snr')
                save_df(getattr(self, 'df_prod_sala_die', None), 'salamanca_diesel_produccion')
                save_df(getattr(self, 'df_sim_sala_die', None), 'salamanca_diesel_simulacion_anual')

                # Guardar datos de Salamanca - Turbosina
                save_df(getattr(self, 'df_data_sala_turb', None), 'salamanca_turbosina_tabla_principal')
                save_df(getattr(self, 'df_snr_sala_turb', None), 'salamanca_turbosina_programa_snr')
                save_df(getattr(self, 'df_prod_sala_turb', None), 'salamanca_turbosina_produccion')
                save_df(getattr(self, 'df_sim_sala_turb', None), 'salamanca_turbosina_simulacion_anual')

                # Guardar datos de Salamanca - Combustoleo
                save_df(getattr(self, 'df_data_sala_comb', None), 'salamanca_combustoleo_tabla_principal')
                save_df(getattr(self, 'df_snr_sala_comb', None), 'salamanca_combustoleo_programa_snr')
                save_df(getattr(self, 'df_prod_sala_comb', None), 'salamanca_combustoleo_produccion')
                save_df(getattr(self, 'df_sim_sala_comb', None), 'salamanca_combustoleo_simulacion_anual')

                # Guardar datos de Salina Cruz - Crudo
                save_df(getattr(self, 'df_data_sal_crud', None), 'salina_cruz_crudo_tabla_principal')
                save_df(getattr(self, 'df_snr_sal_crud', None), 'salina_cruz_crudo_programa_snr')
                save_df(getattr(self, 'df_prod_sal_crud', None), 'salina_cruz_crudo_produccion')
                save_df(getattr(self, 'df_sim_sal_crud', None), 'salina_cruz_crudo_simulacion_anual')

                # Guardar datos de Salina Cruz - Gasolinas
                save_df(getattr(self, 'df_data_sal_gas', None), 'salina_cruz_gasolinas_tabla_principal')
                save_df(getattr(self, 'df_snr_sal_gas', None), 'salina_cruz_gasolinas_programa_snr')
                save_df(getattr(self, 'df_prod_sal_gas', None), 'salina_cruz_gasolinas_produccion')
                save_df(getattr(self, 'df_sim_sal_gas', None), 'salina_cruz_gasolinas_simulacion_anual')

                # Guardar datos de Salina Cruz - Diesel
                save_df(getattr(self, 'df_data_sal_die', None), 'salina_cruz_diesel_tabla_principal')
                save_df(getattr(self, 'df_snr_sal_die', None), 'salina_cruz_diesel_programa_snr')
                save_df(getattr(self, 'df_prod_sal_die', None), 'salina_cruz_diesel_produccion')
                save_df(getattr(self, 'df_sim_sal_die', None), 'salina_cruz_diesel_simulacion_anual')

                # Guardar datos de Salina Cruz - Turbosina
                save_df(getattr(self, 'df_data_sal_turb', None), 'salina_cruz_turbosina_tabla_principal')
                save_df(getattr(self, 'df_snr_sal_turb', None), 'salina_cruz_turbosina_programa_snr')
                save_df(getattr(self, 'df_prod_sal_turb', None), 'salina_cruz_turbosina_produccion')
                save_df(getattr(self, 'df_sim_sal_turb', None), 'salina_cruz_turbosina_simulacion_anual')

                # Guardar datos de Salina Cruz - Combustoleo
                save_df(getattr(self, 'df_data_sal_comb', None), 'salina_cruz_combustoleo_tabla_principal')
                save_df(getattr(self, 'df_snr_sal_comb', None), 'salina_cruz_combustoleo_programa_snr')
                save_df(getattr(self, 'df_prod_sal_comb', None), 'salina_cruz_combustoleo_produccion')
                save_df(getattr(self, 'df_sim_sal_comb', None), 'salina_cruz_combustoleo_simulacion_anual')

                # Guardar datos de Tula - Crudo
                save_df(getattr(self, 'df_data_tula_crud', None), 'tula_crudo_tabla_principal')
                save_df(getattr(self, 'df_snr_tula_crud', None), 'tula_crudo_programa_snr')
                save_df(getattr(self, 'df_prod_tula_crud', None), 'tula_crudo_produccion')
                save_df(getattr(self, 'df_sim_tula_crud', None), 'tula_crudo_simulacion_anual')

                # Guardar datos de Tula - Gasolinas
                save_df(getattr(self, 'df_data_tula_gas', None), 'tula_gasolinas_tabla_principal')
                save_df(getattr(self, 'df_snr_tula_gas', None), 'tula_gasolinas_programa_snr')
                save_df(getattr(self, 'df_prod_tula_gas', None), 'tula_gasolinas_produccion')
                save_df(getattr(self, 'df_sim_tula_gas', None), 'tula_gasolinas_simulacion_anual')

                # Guardar datos de Tula - Diesel
                save_df(getattr(self, 'df_data_tula_die', None), 'tula_diesel_tabla_principal')
                save_df(getattr(self, 'df_snr_tula_die', None), 'tula_diesel_programa_snr')
                save_df(getattr(self, 'df_prod_tula_die', None), 'tula_diesel_produccion')
                save_df(getattr(self, 'df_sim_tula_die', None), 'tula_diesel_simulacion_anual')

                # Guardar datos de Tula - Turbosina
                save_df(getattr(self, 'df_data_tula_turb', None), 'tula_turbosina_tabla_principal')
                save_df(getattr(self, 'df_snr_tula_turb', None), 'tula_turbosina_programa_snr')
                save_df(getattr(self, 'df_prod_tula_turb', None), 'tula_turbosina_produccion')
                save_df(getattr(self, 'df_sim_tula_turb', None), 'tula_turbosina_simulacion_anual')

                # Guardar datos de Tula - Combustoleo
                save_df(getattr(self, 'df_data_tula_comb', None), 'tula_combustoleo_tabla_principal')
                save_df(getattr(self, 'df_snr_tula_comb', None), 'tula_combustoleo_programa_snr')
                save_df(getattr(self, 'df_prod_tula_comb', None), 'tula_combustoleo_produccion')
                save_df(getattr(self, 'df_sim_tula_comb', None), 'tula_combustoleo_simulacion_anual')

                # Guardar datos de Olmeca - Crudo
                save_df(getattr(self, 'df_data_olme_crud', None), 'olmeca_crudo_tabla_principal')
                save_df(getattr(self, 'df_snr_olme_crud', None), 'olmeca_crudo_programa_snr')
                save_df(getattr(self, 'df_prod_olme_crud', None), 'olmeca_crudo_produccion')
                save_df(getattr(self, 'df_sim_olme_crud', None), 'olmeca_crudo_simulacion_anual')

                # Guardar datos de Olmeca - Gasolinas
                save_df(getattr(self, 'df_data_olme_gas', None), 'olmeca_gasolinas_tabla_principal')
                save_df(getattr(self, 'df_snr_olme_gas', None), 'olmeca_gasolinas_programa_snr')
                save_df(getattr(self, 'df_prod_olme_gas', None), 'olmeca_gasolinas_produccion')
                save_df(getattr(self, 'df_sim_olme_gas', None), 'olmeca_gasolinas_simulacion_anual')

                # Guardar datos de Olmeca - Diesel
                save_df(getattr(self, 'df_data_olme_die', None), 'olmeca_diesel_tabla_principal')
                save_df(getattr(self, 'df_snr_olme_die', None), 'olmeca_diesel_programa_snr')
                save_df(getattr(self, 'df_prod_olme_die', None), 'olmeca_diesel_produccion')
                save_df(getattr(self, 'df_sim_olme_die', None), 'olmeca_diesel_simulacion_anual')

                conn.commit()
                conn.close()
                messagebox.showinfo("Éxito", f"¡Los datos de todos los procesos han sido guardados en la base de datos!\n\nRuta:\n{db_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al guardar en SQLite:\n{str(e)}")

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            initialdir=self.default_excel_dir,
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
                            df_data_mina_comb=None, df_snr_mina_comb=None, df_prod_mina_comb=None, df_sim_mina_comb=None,
                            df_data_sala_crud=None, df_snr_sala_crud=None, df_prod_sala_crud=None, df_sim_sala_crud=None,
                            df_data_sala_gas=None, df_snr_sala_gas=None, df_prod_sala_gas=None, df_sim_sala_gas=None,
                            df_data_sala_die=None, df_snr_sala_die=None, df_prod_sala_die=None, df_sim_sala_die=None,
                            df_data_sala_turb=None, df_snr_sala_turb=None, df_prod_sala_turb=None, df_sim_sala_turb=None,
                            df_data_sala_comb=None, df_snr_sala_comb=None, df_prod_sala_comb=None, df_sim_sala_comb=None,
                            df_data_sal_crud=None, df_snr_sal_crud=None, df_prod_sal_crud=None, df_sim_sal_crud=None,
                            df_data_sal_gas=None, df_snr_sal_gas=None, df_prod_sal_gas=None, df_sim_sal_gas=None,
                            df_data_sal_die=None, df_snr_sal_die=None, df_prod_sal_die=None, df_sim_sal_die=None,
                            df_data_sal_turb=None, df_snr_sal_turb=None, df_prod_sal_turb=None, df_sim_sal_turb=None,
                            df_data_sal_comb=None, df_snr_sal_comb=None, df_prod_sal_comb=None, df_sim_sal_comb=None,
                            df_data_tula_crud=None, df_snr_tula_crud=None, df_prod_tula_crud=None, df_sim_tula_crud=None,
                            df_data_tula_gas=None, df_snr_tula_gas=None, df_prod_tula_gas=None, df_sim_tula_gas=None,
                            df_data_tula_die=None, df_snr_tula_die=None, df_prod_tula_die=None, df_sim_tula_die=None,
                            df_data_tula_turb=None, df_snr_tula_turb=None, df_prod_tula_turb=None, df_sim_tula_turb=None,
                            df_data_tula_comb=None, df_snr_tula_comb=None, df_prod_tula_comb=None, df_sim_tula_comb=None,
                            df_data_olme_crud=None, df_snr_olme_crud=None, df_prod_olme_crud=None, df_sim_olme_crud=None,
                            df_data_olme_gas=None, df_snr_olme_gas=None, df_prod_olme_gas=None, df_sim_olme_gas=None,
                            df_data_olme_die=None, df_snr_olme_die=None, df_prod_olme_die=None, df_sim_olme_die=None):
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

        self.df_data_sala_crud = df_data_sala_crud
        self.df_snr_sala_crud = df_snr_sala_crud
        self.df_prod_sala_crud = df_prod_sala_crud
        self.df_sim_sala_crud = df_sim_sala_crud

        self.df_data_sala_gas = df_data_sala_gas
        self.df_snr_sala_gas = df_snr_sala_gas
        self.df_prod_sala_gas = df_prod_sala_gas
        self.df_sim_sala_gas = df_sim_sala_gas

        self.df_data_sala_die = df_data_sala_die
        self.df_snr_sala_die = df_snr_sala_die
        self.df_prod_sala_die = df_prod_sala_die
        self.df_sim_sala_die = df_sim_sala_die

        self.df_data_sala_turb = df_data_sala_turb
        self.df_snr_sala_turb = df_snr_sala_turb
        self.df_prod_sala_turb = df_prod_sala_turb
        self.df_sim_sala_turb = df_sim_sala_turb

        self.df_data_sala_comb = df_data_sala_comb
        self.df_snr_sala_comb = df_snr_sala_comb
        self.df_prod_sala_comb = df_prod_sala_comb
        self.df_sim_sala_comb = df_sim_sala_comb

        self.df_data_sal_crud = df_data_sal_crud
        self.df_snr_sal_crud = df_snr_sal_crud
        self.df_prod_sal_crud = df_prod_sal_crud
        self.df_sim_sal_crud = df_sim_sal_crud

        self.df_data_sal_gas = df_data_sal_gas
        self.df_snr_sal_gas = df_snr_sal_gas
        self.df_prod_sal_gas = df_prod_sal_gas
        self.df_sim_sal_gas = df_sim_sal_gas

        self.df_data_sal_die = df_data_sal_die
        self.df_snr_sal_die = df_snr_sal_die
        self.df_prod_sal_die = df_prod_sal_die
        self.df_sim_sal_die = df_sim_sal_die

        self.df_data_sal_turb = df_data_sal_turb
        self.df_snr_sal_turb = df_snr_sal_turb
        self.df_prod_sal_turb = df_prod_sal_turb
        self.df_sim_sal_turb = df_sim_sal_turb

        self.df_data_sal_comb = df_data_sal_comb
        self.df_snr_sal_comb = df_snr_sal_comb
        self.df_prod_sal_comb = df_prod_sal_comb
        self.df_sim_sal_comb = df_sim_sal_comb

        self.df_data_tula_crud = df_data_tula_crud
        self.df_snr_tula_crud = df_snr_tula_crud
        self.df_prod_tula_crud = df_prod_tula_crud
        self.df_sim_tula_crud = df_sim_tula_crud

        self.df_data_tula_gas = df_data_tula_gas
        self.df_snr_tula_gas = df_snr_tula_gas
        self.df_prod_tula_gas = df_prod_tula_gas
        self.df_sim_tula_gas = df_sim_tula_gas

        self.df_data_tula_die = df_data_tula_die
        self.df_snr_tula_die = df_snr_tula_die
        self.df_prod_tula_die = df_prod_tula_die
        self.df_sim_tula_die = df_sim_tula_die

        self.df_data_tula_turb = df_data_tula_turb
        self.df_snr_tula_turb = df_snr_tula_turb
        self.df_prod_tula_turb = df_prod_tula_turb
        self.df_sim_tula_turb = df_sim_tula_turb

        self.df_data_tula_comb = df_data_tula_comb
        self.df_snr_tula_comb = df_snr_tula_comb
        self.df_prod_tula_comb = df_prod_tula_comb
        self.df_sim_tula_comb = df_sim_tula_comb

        self.df_data_olme_crud = df_data_olme_crud
        self.df_snr_olme_crud = df_snr_olme_crud
        self.df_prod_olme_crud = df_prod_olme_crud
        self.df_sim_olme_crud = df_sim_olme_crud

        self.df_data_olme_gas = df_data_olme_gas
        self.df_snr_olme_gas = df_snr_olme_gas
        self.df_prod_olme_gas = df_prod_olme_gas
        self.df_sim_olme_gas = df_sim_olme_gas

        self.df_data_olme_die = df_data_olme_die
        self.df_snr_olme_die = df_snr_olme_die
        self.df_prod_olme_die = df_prod_olme_die
        self.df_sim_olme_die = df_sim_olme_die

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
        elif selection == "Salamanca -Crudo":
            df_data = self.df_data_sala_crud
            df_snr = self.df_snr_sala_crud
            df_prod = self.df_prod_sala_crud
            df_sim = self.df_sim_sala_crud
            lbl2_txt = "Programa de Crudo Salamanca (Col BL x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Crudo Salamanca (BE-BF, Filas 21-40)"
        elif selection == "Salamanca -Gasolinas":
            df_data = self.df_data_sala_gas
            df_snr = self.df_snr_sala_gas
            df_prod = self.df_prod_sala_gas
            df_sim = self.df_sim_sala_gas
            lbl2_txt = "Programa de Gasolinas Salamanca (Col BV x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Gasolinas Salamanca (BW-BX, Filas 21-40)"
        elif selection == "Salamanca -Diesel":
            df_data = self.df_data_sala_die
            df_snr = self.df_snr_sala_die
            df_prod = self.df_prod_sala_die
            df_sim = self.df_sim_sala_die
            lbl2_txt = "Programa de Diesel Salamanca (Col CF x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Diesel Salamanca (CO-CP, Filas 21-40)"
        elif selection == "Salamanca -Turbosina":
            df_data = self.df_data_sala_turb
            df_snr = self.df_snr_sala_turb
            df_prod = self.df_prod_sala_turb
            df_sim = self.df_sim_sala_turb
            lbl2_txt = "Programa de Turbosina Salamanca (Col CN x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Turbosina Salamanca (CZ-DA, Filas 21-40)"
        elif selection == "Salamanca -Combustoleo":
            df_data = self.df_data_sala_comb
            df_snr = self.df_snr_sala_comb
            df_prod = self.df_prod_sala_comb
            df_sim = self.df_sim_sala_comb
            lbl2_txt = "Programa de Combustoleo Salamanca (Col CY x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Combustoleo Salamanca (AQ-AR, Filas 158-179)"
        elif selection == "Salina Cruz -Crudo":
            df_data = self.df_data_sal_crud
            df_snr = self.df_snr_sal_crud
            df_prod = self.df_prod_sal_crud
            df_sim = self.df_sim_sal_crud
            lbl2_txt = "Programa de Crudo Salina Cruz (Col BM x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Crudo Salina Cruz (BI-BJ, Filas 21-40)"
        elif selection == "Salina Cruz -Gasolinas":
            df_data = self.df_data_sal_gas
            df_snr = self.df_snr_sal_gas
            df_prod = self.df_prod_sal_gas
            df_sim = self.df_sim_sal_gas
            lbl2_txt = "Programa de Gasolinas Salina Cruz (Col BW x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Gasolinas Salina Cruz (BY-BZ, Filas 21-40)"
        elif selection == "Salina Cruz -Diesel":
            df_data = self.df_data_sal_die
            df_snr = self.df_snr_sal_die
            df_prod = self.df_prod_sal_die
            df_sim = self.df_sim_sal_die
            lbl2_txt = "Programa de Diesel Salina Cruz (Col CG x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Diesel Salina Cruz (CQ-CR, Filas 21-40)"
        elif selection == "Salina Cruz -Turbosina":
            df_data = self.df_data_sal_turb
            df_snr = self.df_snr_sal_turb
            df_prod = self.df_prod_sal_turb
            df_sim = self.df_sim_sal_turb
            lbl2_txt = "Programa de Turbosina Salina Cruz (Col CO x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Turbosina Salina Cruz (DB-DC, Filas 21-40)"
        elif selection == "Salina Cruz -Combustoleo":
            df_data = self.df_data_sal_comb
            df_snr = self.df_snr_sal_comb
            df_prod = self.df_prod_sal_comb
            df_sim = self.df_sim_sal_comb
            lbl2_txt = "Programa de Combustoleo Salina Cruz (Col CZ x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Combustoleo Salina Cruz (AT-AU, Filas 158-179)"
        elif selection == "Tula -Crudo":
            df_data = self.df_data_tula_crud
            df_snr = self.df_snr_tula_crud
            df_prod = self.df_prod_tula_crud
            df_sim = self.df_sim_tula_crud
            lbl2_txt = "Programa de Crudo Tula (Col BN x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Crudo Tula (BK-BL, Filas 21-40)"
        elif selection == "Tula -Gasolinas":
            df_data = self.df_data_tula_gas
            df_snr = self.df_snr_tula_gas
            df_prod = self.df_prod_tula_gas
            df_sim = self.df_sim_tula_gas
            lbl2_txt = "Programa de Gasolinas Tula (Col BX x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Gasolinas Tula (CC-CD, Filas 21-40)"
        elif selection == "Tula -Diesel":
            df_data = self.df_data_tula_die
            df_snr = self.df_snr_tula_die
            df_prod = self.df_prod_tula_die
            df_sim = self.df_sim_tula_die
            lbl2_txt = "Programa de Diesel Tula (Col CH x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Diesel Tula (CT-CU, Filas 21-40)"
        elif selection == "Tula -Turbosina":
            df_data = self.df_data_tula_turb
            df_snr = self.df_snr_tula_turb
            df_prod = self.df_prod_tula_turb
            df_sim = self.df_sim_tula_turb
            lbl2_txt = "Programa de Turbosina Tula (Col CP x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Turbosina Tula (DD-DE, Filas 21-40)"
        elif selection == "Tula -Combustoleo":
            df_data = self.df_data_tula_comb
            df_snr = self.df_snr_tula_comb
            df_prod = self.df_prod_tula_comb
            df_sim = self.df_sim_tula_comb
            lbl2_txt = "Programa de Combustoleo Tula (Col DA x2, Filas 74-104)"
            lbl3_txt = "Fecha y Producción de Combustoleo Tula (AW-AX, Filas 158-179)"
        elif selection == "Olmeca -Crudo":
            df_data = self.df_data_olme_crud
            df_snr = self.df_snr_olme_crud
            df_prod = self.df_prod_olme_crud
            df_sim = self.df_sim_olme_crud
            lbl2_txt = "Programa de Crudo Olmeca (Col N x2, Filas 159-189)"
            lbl3_txt = "Fecha y Producción de Crudo Olmeca (C-D, Filas 192-206)"
        elif selection == "Olmeca -Gasolinas":
            df_data = self.df_data_olme_gas
            df_snr = self.df_snr_olme_gas
            df_prod = self.df_prod_olme_gas
            df_sim = self.df_sim_olme_gas
            lbl2_txt = "Programa de Gasolinas Olmeca (Col O x2, Filas 159-189)"
            lbl3_txt = "Fecha y Producción de Gasolinas Olmeca (F-G, Filas 192-206)"
        elif selection == "Olmeca -Diesel":
            df_data = self.df_data_olme_die
            df_snr = self.df_snr_olme_die
            df_prod = self.df_prod_olme_die
            df_sim = self.df_sim_olme_die
            lbl2_txt = "Programa de Diesel Olmeca (Col P x2, Filas 159-189)"
            lbl3_txt = "Fecha y Producción de Diesel Olmeca (I-J, Filas 192-206)"

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
            command=lambda cell: self.on_table_clicked("diaria", cell),
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
            command=lambda cell: self.on_table_clicked("programa", cell),
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
            command=lambda cell: self.on_table_clicked("historica", cell),
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
            initialdir=self.default_pptx_dir,
            filetypes=[("Archivos PowerPoint", "*.pptx")]
        )

        if file_path:
            save_path = filedialog.asksaveasfilename(
                title="Guardar presentación actualizada",
                initialdir=self.default_pptx_dir,
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

    def on_table_clicked(self, table_name, cell_data):
        row = cell_data.get('row')
        col = cell_data.get('column')
        if row == 0 or row is None:
            return
            
        selection = self.cb_proceso.get()
        if table_name == "diaria":
            headers = self.table.values[0]
            col_name = headers[col]
            row_data = self.table.values[row]
            dia_key = str(row_data[0])
            valor_actual = str(row_data[col])
            
            target_proceso = self.resolve_proceso_name(selection, col_name)
            if target_proceso:
                self.show_edit_delete_dialog(target_proceso, "diaria", dia_key, valor_actual)
                
        elif table_name == "programa":
            headers = self.table2.values[0]
            col_name = headers[col]
            row_data = self.table2.values[row]
            key = f"Row_{row - 1}_{col_name}"
            valor_actual = str(row_data[col])
            
            self.show_edit_delete_dialog(selection, "programa", key, valor_actual)
            
        elif table_name == "historica":
            row_data = self.table3.values[row]
            period_key = str(row_data[0])
            valor_actual = str(row_data[1])
            
            self.show_edit_delete_dialog(selection, "historica", period_key, valor_actual)

    def resolve_proceso_name(self, selection, col_name):
        if " -" in selection:
            return selection
        
        refinery = col_name.split()[0].strip()
        valid_refineries = ["Cadereyta", "Madero", "Minatitlan", "Salamanca", "Salina Cruz", "Tula", "Olmeca"]
        
        def norm(s):
            return s.lower().replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").replace("ñ","n")
            
        ref_norm = norm(refinery)
        matched_ref = None
        for r in valid_refineries:
            if norm(r) in ref_norm or ref_norm in norm(r):
                matched_ref = r
                break
        
        if matched_ref:
            return f"{matched_ref} -{selection}"
        return None

    def show_edit_delete_dialog(self, proceso, tabla, clave, valor):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Editar / Eliminar Valor")
        dialog.geometry("420x260")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        x = self.winfo_x() + (self.winfo_width() - w) // 2
        y = self.winfo_y() + (self.winfo_height() - h) // 2
        dialog.geometry(f"+{x}+{y}")

        type_lbl = ""
        if tabla == "diaria": type_lbl = "Producción Diaria"
        elif tabla == "programa": type_lbl = "Programa de Planificación"
        elif tabla == "historica": type_lbl = "Fecha y Producción Histórica"

        lbl = ctk.CTkLabel(dialog, text=f"Proceso: {proceso}\nTipo: {type_lbl}\nClave: {clave}", font=("Roboto", 13, "bold"))
        lbl.pack(pady=15)

        lbl_val = ctk.CTkLabel(dialog, text="Ingresa el valor deseado:", font=("Roboto", 12))
        lbl_val.pack(pady=5)
        
        entry_val = ctk.CTkEntry(dialog, placeholder_text="Valor numérico", width=200)
        entry_val.insert(0, valor)
        entry_val.pack(pady=5)

        def on_update():
            val_str = entry_val.get().strip()
            if not val_str:
                messagebox.showerror("Error", "Debes ingresar un valor válido.", parent=dialog)
                return
            try:
                nuevo_val = float(val_str)
            except ValueError:
                messagebox.showerror("Error", "El valor debe ser un número.", parent=dialog)
                return
            
            import db_helper
            if tabla == "historica":
                parts = clave.split()
                if len(parts) == 2:
                    mes = parts[0].strip()
                    anio = parts[1].strip()
                else:
                    mes = "AÑO"
                    anio = clave.strip()
                db_helper.save_extra_prod(proceso, anio, mes, nuevo_val)
            else:
                db_helper.save_modificacion(proceso, tabla, clave, nuevo_val)
                
            messagebox.showinfo("Éxito", "Valor actualizado correctamente.", parent=dialog)
            dialog.destroy()
            
            if self.current_file_path:
                self.set_loading_state(True, "Recargando datos...")
                threading.Thread(target=self.async_load_data, args=(self.current_file_path,), daemon=True).start()

        def on_restore():
            if not messagebox.askyesno("Confirmar", "¿Restaurar al valor original del archivo Excel?", parent=dialog):
                return
            
            import db_helper
            if tabla == "historica":
                parts = clave.split()
                if len(parts) == 2:
                    mes = parts[0].strip()
                    anio = parts[1].strip()
                else:
                    mes = "AÑO"
                    anio = clave.strip()
                import sqlite3
                conn = sqlite3.connect(db_helper.DB_PATH)
                c = conn.cursor()
                c.execute('DELETE FROM produccion_extra WHERE proceso=? AND anio=? AND mes=?', (proceso, anio, mes))
                conn.commit()
                conn.close()
            else:
                db_helper.delete_modificacion(proceso, tabla, clave)
                
            messagebox.showinfo("Éxito", "Valor restaurado al original.", parent=dialog)
            dialog.destroy()
            
            if self.current_file_path:
                self.set_loading_state(True, "Recargando datos...")
                threading.Thread(target=self.async_load_data, args=(self.current_file_path,), daemon=True).start()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=20)

        btn_act = ctk.CTkButton(btn_frame, text="Actualizar", command=on_update, width=120, fg_color="#24a0ed", hover_color="#0080ff")
        btn_act.pack(side="left", padx=10)

        btn_rst = ctk.CTkButton(btn_frame, text="Restaurar Original", command=on_restore, width=130, fg_color="#e0a800", hover_color="#c69500", text_color="black")
        btn_rst.pack(side="left", padx=10)

    def open_config_coords_dialog(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Configuración Visual de Coordenadas (Excel)")
        dialog.geometry("780x620")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        x = self.winfo_x() + (self.winfo_width() - w) // 2
        y = self.winfo_y() + (self.winfo_height() - h) // 2
        dialog.geometry(f"+{x}+{y}")

        # Título superior
        title_lbl = ctk.CTkLabel(dialog, text="Mapeador de Celdas y Rangos de Excel", font=("Roboto", 16, "bold"), text_color="#3484F0")
        title_lbl.pack(pady=(15, 10))

        # Selector de proceso
        proc_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        proc_frame.pack(fill="x", padx=30, pady=5)
        
        proc_lbl = ctk.CTkLabel(proc_frame, text="Proceso:", font=("Roboto", 12, "bold"))
        proc_lbl.pack(side="left", padx=(0, 10))
        
        specific_processes = [v for v in self.cb_proceso.cget("values") if " -" in v]
        
        cb_proc = ctk.CTkComboBox(proc_frame, values=specific_processes, font=("Roboto", 12), state="readonly", width=300)
        cb_proc.pack(side="left")
        
        current_sel = self.cb_proceso.get()
        if current_sel in specific_processes:
            cb_proc.set(current_sel)
        else:
            cb_proc.set(specific_processes[0])

        # Contenedor principal de pestañas (inicialmente sin command para evitar UnboundLocalError)
        tabview = ctk.CTkTabview(dialog, width=720, height=200)
        tabview.pack(padx=30, pady=10, fill="x")

        tab_t1 = tabview.add("📊 T1 (Diaria)")
        tab_t2 = tabview.add("📅 T2 (Programa)")
        tab_t3 = tabview.add("📜 T3 (Histórica)")

        entries = {}

        # ═══ Maquetación de la Pestaña 1 (Diaria) ═══
        t1_info = ctk.CTkLabel(tab_t1, text="Configura las coordenadas para los datos de Producción Diaria.", font=("Roboto", 11, "italic"), text_color="#aaaaaa")
        t1_info.pack(pady=(5, 5))
        
        t1_grid = ctk.CTkFrame(tab_t1, fg_color="transparent")
        t1_grid.pack(pady=5)
        
        ctk.CTkLabel(t1_grid, text="Letras de Columna (e.g. A-H o L,M):", font=("Roboto", 12, "bold")).grid(row=0, column=0, sticky="e", padx=10, pady=5)
        entries["d_cols"] = ctk.CTkEntry(t1_grid, width=120, placeholder_text="Ej. L,M")
        entries["d_cols"].grid(row=0, column=1, padx=10, pady=5)
        
        ctk.CTkLabel(t1_grid, text="Rango de Filas (e.g. 21-51):", font=("Roboto", 12, "bold")).grid(row=0, column=2, sticky="e", padx=10, pady=5)
        entries["d_filas"] = ctk.CTkEntry(t1_grid, width=120, placeholder_text="Ej. 21-51")
        entries["d_filas"].grid(row=0, column=3, padx=10, pady=5)

        # ═══ Maquetación de la Pestaña 2 (Programa) ═══
        t2_info = ctk.CTkLabel(tab_t2, text="Configura las coordenadas para las metas de planeación (CMP/PODIM).", font=("Roboto", 11, "italic"), text_color="#aaaaaa")
        t2_info.pack(pady=(5, 5))
        
        t2_grid = ctk.CTkFrame(tab_t2, fg_color="transparent")
        t2_grid.pack(pady=5)
        
        ctk.CTkLabel(t2_grid, text="Columna Programa (e.g. AE o BI):", font=("Roboto", 12, "bold")).grid(row=0, column=0, sticky="e", padx=10, pady=5)
        entries["p_cols"] = ctk.CTkEntry(t2_grid, width=120, placeholder_text="Ej. BI")
        entries["p_cols"].grid(row=0, column=1, padx=10, pady=5)
        
        ctk.CTkLabel(t2_grid, text="Rango de Filas (e.g. 74-104):", font=("Roboto", 12, "bold")).grid(row=0, column=2, sticky="e", padx=10, pady=5)
        entries["p_filas"] = ctk.CTkEntry(t2_grid, width=120, placeholder_text="Ej. 74-104")
        entries["p_filas"].grid(row=0, column=3, padx=10, pady=5)

        # ═══ Maquetación de la Pestaña 3 (Histórica) ═══
        t3_info = ctk.CTkLabel(tab_t3, text="Configura las coordenadas para las fechas e históricos mensuales.", font=("Roboto", 11, "italic"), text_color="#aaaaaa")
        t3_info.pack(pady=(5, 5))
        
        t3_grid = ctk.CTkFrame(tab_t3, fg_color="transparent")
        t3_grid.pack(pady=5)
        
        ctk.CTkLabel(t3_grid, text="Rango Columnas (e.g. AW-AX):", font=("Roboto", 12, "bold")).grid(row=0, column=0, sticky="e", padx=10, pady=5)
        entries["h_cols"] = ctk.CTkEntry(t3_grid, width=120, placeholder_text="Ej. AW-AX")
        entries["h_cols"].grid(row=0, column=1, padx=10, pady=5)
        
        ctk.CTkLabel(t3_grid, text="Rango de Filas (e.g. 21-40):", font=("Roboto", 12, "bold")).grid(row=0, column=2, sticky="e", padx=10, pady=5)
        entries["h_filas"] = ctk.CTkEntry(t3_grid, width=120, placeholder_text="Ej. 21-40")
        entries["h_filas"].grid(row=0, column=3, padx=10, pady=5)

        # ═══ Vista Previa de Datos ═══
        preview_frame = ctk.CTkFrame(dialog)
        preview_frame.pack(fill="both", expand=True, padx=30, pady=10)
        
        preview_header = ctk.CTkFrame(preview_frame, fg_color="transparent")
        preview_header.pack(fill="x", padx=10, pady=(5, 2))
        
        preview_title = ctk.CTkLabel(preview_header, text="🔍 Vista Previa del Excel (Primeras 5 Filas)", font=("Roboto", 11, "bold"), text_color="#3484F0")
        preview_title.pack(side="left")
        
        preview_label = ctk.CTkLabel(preview_frame, text="Presiona 'Generar Vista Previa' para ver los datos del rango seleccionado.", font=("Courier New", 11), justify="left", anchor="nw")
        preview_label.pack(fill="both", expand=True, padx=15, pady=5)

        # Cargar valores actuales en las entradas al cambiar el proceso
        def load_proc_coords(proceso_name):
            for e in entries.values():
                e.delete(0, "end")
                
            import db_helper
            coords = db_helper.get_coordenadas_override(proceso_name)
            if coords:
                entries["d_filas"].insert(0, coords.get("diaria_filas") or "")
                entries["d_cols"].insert(0, coords.get("diaria_cols") or "")
                entries["p_filas"].insert(0, coords.get("programa_filas") or "")
                entries["p_cols"].insert(0, coords.get("programa_cols") or "")
                entries["h_filas"].insert(0, coords.get("historica_filas") or "")
                entries["h_cols"].insert(0, coords.get("historica_cols") or "")
            else:
                entries["d_filas"].configure(placeholder_text="Defecto Excel")
                entries["d_cols"].configure(placeholder_text="Defecto Excel")
                entries["p_filas"].configure(placeholder_text="Defecto Excel")
                entries["p_cols"].configure(placeholder_text="Defecto Excel")
                entries["h_filas"].configure(placeholder_text="Defecto Excel")
                entries["h_cols"].configure(placeholder_text="Defecto Excel")
            
            preview_label.configure(text="Presiona 'Generar Vista Previa' para ver los datos del rango seleccionado.")

        # Función para actualizar la vista previa de los datos (solo bajo demanda)
        def update_preview():
            if not self.current_file_path:
                preview_label.configure(text="Sube un archivo de Excel primero para ver una vista previa.")
                return

            preview_label.configure(text="Cargando vista previa...")
            dialog.update_idletasks()

            try:
                import pandas as pd
                xls = pd.ExcelFile(self.current_file_path)
                sheet_to_use = None
                for sheet in xls.sheet_names:
                    normalizado = sheet.lower().replace(" ", "").replace("_", "")
                    if "enviocalculopromedio" in normalizado or "enviocalculo" in normalizado:
                        sheet_to_use = sheet
                        break
                if not sheet_to_use:
                    sheet_to_use = xls.sheet_names[0]
                
                df_sheet = pd.read_excel(self.current_file_path, sheet_name=sheet_to_use, header=None)

                active_tab = tabview.get()
                prefix = "d_"
                if "Programa" in active_tab:
                    prefix = "p_"
                elif "Histórica" in active_tab:
                    prefix = "h_"
                
                rows_input = entries[f"{prefix}filas"].get().strip()
                cols_input = entries[f"{prefix}cols"].get().strip()
                
                if not rows_input or not cols_input:
                    import db_helper
                    override = db_helper.get_coordenadas_override(cb_proc.get())
                    if override:
                        rows_input = rows_input or override.get(f"{prefix}filas")
                        cols_input = cols_input or override.get(f"{prefix}cols")

                if not rows_input or not cols_input:
                    preview_label.configure(text="Usando coordenadas preestablecidas de Excel.\nEscribe columnas y filas personalizadas para previsualizar aquí.")
                    return

                parts = rows_input.split("-")
                r_start = int(parts[0]) - 1
                r_end = int(parts[1])
                
                cols_str = cols_input.replace(" ", "")
                def col_to_num(col_str):
                    col_str = col_str.upper().strip()
                    num = 0
                    for c in col_str:
                        if 'A' <= c <= 'Z':
                            num = num * 26 + (ord(c) - ord('A') + 1)
                    return num - 1
                
                if "-" in cols_str:
                    col_parts = cols_str.split("-")
                    c_indices = list(range(col_to_num(col_parts[0]), col_to_num(col_parts[1]) + 1))
                elif "," in cols_str:
                    c_indices = [col_to_num(x) for x in cols_str.split(",")]
                else:
                    c_indices = [col_to_num(cols_str)]
                
                df_slice = df_sheet.iloc[r_start:r_end, c_indices].head(5)
                text_out = df_slice.to_string(index=True, header=False)
                preview_label.configure(text=f"Pestaña activa: {active_tab}  |  Cols: {cols_input}  |  Filas: {rows_input}\n\n{text_out}")
            except Exception as e:
                preview_label.configure(text=f"No se pudo generar la vista previa.\nRevisa la sintaxis de las coordenadas.\n\nDetalle: {str(e)}")

        btn_preview = ctk.CTkButton(preview_header, text="Generar Vista Previa", command=update_preview, width=170, height=28, font=("Roboto", 11, "bold"), fg_color="#17a2b8", hover_color="#138496")
        btn_preview.pack(side="right")

        cb_proc.configure(command=load_proc_coords)
        load_proc_coords(cb_proc.get())

        def on_save():
            proceso_sel = cb_proc.get()
            d_f = entries["d_filas"].get().strip()
            d_c = entries["d_cols"].get().strip()
            p_f = entries["p_filas"].get().strip()
            p_c = entries["p_cols"].get().strip()
            h_f = entries["h_filas"].get().strip()
            h_c = entries["h_cols"].get().strip()

            def is_valid_range(r_str):
                if not r_str: return True
                return "-" in r_str and len(r_str.split("-")) == 2 and all(x.isdigit() for x in r_str.replace(" ", "").split("-"))

            if not all(is_valid_range(r) for r in [d_f, p_f, h_f]):
                messagebox.showerror("Error de Formato", "Los rangos de filas deben tener el formato 'Inicio-Fin' (ej. 21-51).", parent=dialog)
                return

            import db_helper
            db_helper.save_coordenadas_override(proceso_sel, d_f, d_c, p_f, p_c, h_f, h_c)
            messagebox.showinfo("Éxito", f"Coordenadas para '{proceso_sel}' guardadas en base de datos.", parent=dialog)
            dialog.destroy()

            if self.current_file_path:
                self.set_loading_state(True, "Recargando datos del Excel con nuevas coordenadas...")
                threading.Thread(target=self.async_load_data, args=(self.current_file_path,), daemon=True).start()

        def on_restore_default():
            proceso_sel = cb_proc.get()
            if not messagebox.askyesno("Confirmar", f"¿Deseas eliminar la configuración personalizada de '{proceso_sel}' y volver a las coordenadas por defecto del Excel?", parent=dialog):
                return
            
            import db_helper
            db_helper.delete_coordenadas_override(proceso_sel)
            messagebox.showinfo("Éxito", f"Configuración de '{proceso_sel}' restaurada a valores por defecto.", parent=dialog)
            dialog.destroy()

            if self.current_file_path:
                self.set_loading_state(True, "Recargando datos del Excel con coordenadas por defecto...")
                threading.Thread(target=self.async_load_data, args=(self.current_file_path,), daemon=True).start()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=30, pady=15)

        btn_save = ctk.CTkButton(btn_frame, text="Guardar Cambios", command=on_save, width=160, fg_color="#28a745", hover_color="#218838")
        btn_save.pack(side="left", padx=10)

        btn_restore = ctk.CTkButton(btn_frame, text="Restaurar Defectos", command=on_restore_default, width=160, fg_color="#6c757d", hover_color="#5a6268")
        btn_restore.pack(side="left", padx=10)

        btn_cancel = ctk.CTkButton(btn_frame, text="Cancelar", command=dialog.destroy, width=100, fg_color="#dc3545", hover_color="#c82333")
        btn_cancel.pack(side="right", padx=10)

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
