import os
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import traceback
import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor

master_current_month = None

import re as _re
import unicodedata

_PRODUCTOS = [("crudo", "crudo"), ("gasolina", "gasolinas"), ("diesel", "diesel"),
              ("turbosina", "turbosina"), ("asfalto", "asfalto"), ("combust", "combustoleo")]
_REFINERIAS = [("salina cruz", "salina cruz"), ("cadereyta", "cadereyta"), ("madero", "madero"),
               ("minatit", "minatitlan"), ("salamanca", "salamanca"), ("tula", "tula"),
               ("olmeca", "olmeca"), ("snr", "snr")]


def _normalizar(texto):
    texto = unicodedata.normalize("NFD", str(texto))
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return _re.sub(r"\s+", " ", texto).strip().lower()


_MESES_ABR = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def days_in_current_month(month_label, year):
    """Días del mes en curso (28/29/30/31) según la etiqueta del último mes con datos ('Sep', 'Feb'...).

    Si la etiqueta no se reconoce, usa el mes del sistema.
    """
    import calendar
    from datetime import datetime
    now = datetime.now()
    abr = str(month_label or "").strip().lower()[:3]
    month = _MESES_ABR.index(abr) + 1 if abr in _MESES_ABR else now.month
    return calendar.monthrange(int(year or now.year), month)[1]


def classify_slide_title(text):
    """Devuelve (refinería, producto) si el texto es un título de proceso/producción, o None.

    Ej.: 'Producción de Gasolinas Salina Cruz Programa: 93.0 Mbd' -> ('salina cruz', 'gasolinas').
    Ignora 'Rendimientos ...' y cualquier texto que no empiece con 'Proceso de'/'Producción de'.
    """
    t = _normalizar(text)
    if not (t.startswith("proceso de") or t.startswith("produccion de")):
        return None
    producto = next((v for k, v in _PRODUCTOS if k in t), None)
    refineria = next((v for k, v in _REFINERIAS if k in t), None)
    if producto and refineria:
        return refineria, producto
    return None


def build_slides_map(prs):
    """Mapa {(refinería, producto): diapositiva} leyendo el título de cada diapositiva."""
    slides_map, duplicated = {}, []
    for slide in prs.slides:
        key = None
        for shape in slide.shapes:
            if shape.has_text_frame:
                key = classify_slide_title(shape.text_frame.text)
                if key:
                    break
        if key:
            if key in slides_map:
                duplicated.append(key)
            else:
                slides_map[key] = slide
    return slides_map, duplicated


class _MissingSlide:
    """Diapositiva ausente: sin formas, así los bloques que buscan una gráfica la omiten."""
    shapes = []


def find_slide(slides_map, refineria, producto):
    return slides_map.get((refineria, producto)) or _MissingSlide()


def slide_exists(slides_map, refineria, producto):
    return (refineria, producto) in slides_map


def update_slide_chart(chart, categories, proceso_vals, diario_vals, programa_vals, columna1_vals, wine_color, green_color, current_month_name=None):
    global master_current_month
    if not current_month_name:
        current_month_name = master_current_month
    from pptx.chart.data import CategoryChartData
    from pptx.util import Pt

    # Enviar a la gráfica valores ya redondeados a enteros.
    # (los cálculos internos más abajo siguen usando los datos originales)
    def _round(v):
        if v is None:
            return None
        try:
            return int(round(float(v)))
        except:
            return v

    proceso_r = [_round(v) for v in proceso_vals]
    diario_r = [_round(v) for v in diario_vals]
    programa_r = [_round(v) for v in programa_vals]
    columna1_r = [_round(v) for v in columna1_vals]

    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series('PROCESO', tuple(proceso_r))
    chart_data.add_series('Diario', tuple(diario_r))
    chart_data.add_series('Programa', tuple(programa_r))
    chart_data.add_series('Columna1', tuple(columna1_r))

    # Reemplazar los datos de la gráfica
    chart.replace_data(chart_data)

    # --- Restaurar y pintar los colores dinámicamente ---
    if len(chart.series) > 0:
        series = chart.series[0]
        
        # Buscar el último mes con producción real en la nueva lista de categorías
        # Determinar cuál es el último mes con datos para pintarlo de vino
        last_month_idx = -1
        for i, val in enumerate(proceso_vals):
            if i < len(categories) and any(c.isalpha() for c in categories[i]):
                if current_month_name and categories[i].strip().lower()[:3] == current_month_name.strip().lower()[:3]:
                    last_month_idx = i
                elif val is not None and val != 0:
                    last_month_idx = i

        # Eliminar TODOS los formatos específicos de puntos (<c:dPt>)
        # Esto obliga a que todas las barras hereden el color gris por defecto de la serie
        try:
            ser_el = series._element
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            
            for _ in range(2):
                dpts = ser_el.findall('c:dPt', ns)
                for dpt in dpts:
                    ser_el.remove(dpt)
        except Exception:
            pass

        # Definir color gris para los años
        from pptx.dml.color import RGBColor
        gray_color = RGBColor(192, 192, 192)

        # Pintar todas las barras y formatear sus etiquetas de datos
        for p_idx in range(len(series.points)):
            if p_idx >= len(categories):
                break
            cat = categories[p_idx]
            if cat.isdigit() and len(cat) <= 2:
                break
            try:
                point = series.points[p_idx]
                fill = point.format.fill
                fill.solid()
                
                # Si es un mes, aplicamos vino o verde
                if any(c.isalpha() for c in cat):
                    if p_idx == last_month_idx:
                        fill.fore_color.rgb = wine_color
                        try:
                            point.data_label.font.size = Pt(14)
                            point.data_label.font.bold = True
                        except Exception:
                            pass
                    else:
                        fill.fore_color.rgb = green_color
                        try:
                            point.data_label.font.size = Pt(9)
                            point.data_label.font.bold = False
                        except Exception:
                            pass
                        
                    # Limpiar modificadores de color
                    try:
                        ns_a = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                        srgb_clr = point._element.find('.//a:srgbClr', ns_a)
                        if srgb_clr is not None:
                            for child in list(srgb_clr):
                                if 'lumMod' in child.tag or 'lumOff' in child.tag:
                                    srgb_clr.remove(child)
                    except Exception:
                        pass
                else:
                    # Si es un año, aplicamos el gris explícitamente y reseteamos fuente
                    fill.fore_color.rgb = gray_color
                    try:
                        point.data_label.font.size = Pt(9)
                        point.data_label.font.bold = False
                    except Exception:
                        pass
                    
            except Exception:
                pass


def export_to_pptx(app, file_path, save_path):
    try:
        from pptx import Presentation
        from pptx.dml.color import RGBColor

        prs = Presentation(file_path)
        # Las gráficas se asignan leyendo el título de cada diapositiva, no su posición.
        slides_map, duplicated = build_slides_map(prs)
        app.pptx_slides_found = sorted(f"{r} - {p}" for r, p in slides_map)
        for ref, prod in duplicated:
            print(f"[PPTX] Aviso: hay más de una diapositiva '{prod} {ref}'; se usa la primera.")
        if not slides_map:
            raise ValueError("No se encontró ninguna diapositiva con título de proceso o producción "
                             "(p. ej. 'Proceso de crudo Madero' o 'Producción de Gasolinas SNR').")
 
        # --- 1. PROCESAR DIAPOSITIVA DE CRUDO (DIAPOSITIVA 2) ---
        slide = find_slide(slides_map, "snr", "crudo")
        chart = None

        for shape in slide.shapes:
            if shape.has_chart:
                chart = shape.chart
            elif shape.has_text_frame and "CMP:" in shape.text:
                cmp_val = getattr(app, 'cmp_value', "1234.8")
                import re
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if "CMP:" in run.text:
                            run.text = re.sub(r'CMP:\s*[\d\.]+', f'CMP: {cmp_val}', run.text)

        if not chart:
            raise ValueError("No se encontró una diapositiva con gráfica para Crudo (el título debe ser p. ej. 'Producción de Crudo SNR').")

        # Columnas específicas de la tabla 1 para Crudo: 'Crudo' (0), 'Cadereyta' (1)
        # En la tabla, "Cadereyta" es la que va en la columna de proceso.
        # Los datos de producción diaria están en la segunda columna (df_data.iloc[:, 1])
        proceso_col = app.df_data.columns[1]

        categories = []
        proceso_vals = []
        diario_vals = []
        programa_vals = []
        columna1_vals = []

        # Filtrar df_prod para quedarnos con los años y los meses activos
        prod_rows = []
        for idx, row in app.df_prod.iterrows():
            cat = str(row.iloc[0]).strip()
            val = row.iloc[1]
            if not cat: continue
            if not any(c.isalpha() for c in cat):
                prod_rows.append((cat, val))
            else:
                if val != "" and val is not None:
                    try:
                        float(val)
                        prod_rows.append((cat, val))
                    except: pass
        
        if len(prod_rows) > 30: prod_rows = prod_rows[-30:]

        global master_current_month
        master_current_month = None
        for cat, val in prod_rows:
            if any(c.isalpha() for c in cat):
                master_current_month = cat

        # Las gráficas diarias llevan solo los días que tiene el mes en curso (no siempre 31)
        dias_mes = days_in_current_month(master_current_month, getattr(app, "sheet_year", None))

        for i in range(len(prod_rows)):
            categories.append(prod_rows[i][0])
            try: proceso_vals.append(float(prod_rows[i][1]))
            except: proceso_vals.append(None)
            diario_vals.append(None)
            programa_vals.append(None)
            columna1_vals.append(None)

        # Llenar datos diarios (días del mes en curso)
        for i in range(dias_mes):
            categories.append(str(i + 1))
            proceso_vals.append(None)
            
            try: diario_vals.append(float(app.df_data[proceso_col].iloc[i]))
            except: diario_vals.append(None)
            
            try: programa_vals.append(float(app.df_snr.iloc[i, 0]))
            except: programa_vals.append(None)
            
            try: columna1_vals.append(float(app.df_snr.iloc[i, 1]))
            except: columna1_vals.append(None)

        # --- EXTRAER COLORES DE LA PLANTILLA ---
        wine_color = RGBColor(0x69, 0x19, 0x32)
        green_color = RGBColor(0x24, 0x5C, 0x4F)

        # Actualizar gráfica
        update_slide_chart(chart, categories, proceso_vals, diario_vals, programa_vals, columna1_vals, wine_color, green_color)


        # --- NUEVA SECCIÓN: PROCESAR DIAPOSITIVA 10 (Crudo Cadereyta) ---
        slide_cad = find_slide(slides_map, "cadereyta", "crudo")
        chart_cad = None
        for shape in slide_cad.shapes:
            if shape.has_chart:
                chart_cad = shape.chart
                break
        
        if chart_cad:
            categories_cad = []
            proceso_vals_cad = []
            diario_vals_cad = []
            programa_vals_cad = []
            columna1_vals_cad = []

            # Filtrar df_prod_cad
            prod_rows_cad = []
            for idx, row in app.df_prod_cad.iterrows():
                cat = str(row.iloc[0]).strip()
                val = row.iloc[1]
                if not cat: continue
                if not any(c.isalpha() for c in cat):
                    prod_rows_cad.append((cat, val))
                else:
                    if val != "" and val is not None:
                        try:
                            float(val)
                            prod_rows_cad.append((cat, val))
                        except: pass
            
            if len(prod_rows_cad) > 30: prod_rows_cad = prod_rows_cad[-30:]
            
            for i in range(len(prod_rows_cad)):
                categories_cad.append(prod_rows_cad[i][0])
                try: proceso_vals_cad.append(float(prod_rows_cad[i][1]))
                except: proceso_vals_cad.append(None)
                diario_vals_cad.append(None)
                programa_vals_cad.append(None)
                columna1_vals_cad.append(None)

            proceso_col_cad = app.df_data_cad.columns[1]

            for i in range(dias_mes):
                categories_cad.append(str(i + 1))
                proceso_vals_cad.append(None)
                
                try: diario_vals_cad.append(float(app.df_data_cad[proceso_col_cad].iloc[i]))
                except: diario_vals_cad.append(None)
                
                try: programa_vals_cad.append(float(app.df_snr_cad.iloc[i, 0]))
                except: programa_vals_cad.append(None)
                
                try: columna1_vals_cad.append(float(app.df_snr_cad.iloc[i, 1]))
                except: columna1_vals_cad.append(None)

            update_slide_chart(chart_cad, categories_cad, proceso_vals_cad, diario_vals_cad, programa_vals_cad, columna1_vals_cad, wine_color, green_color)


        # --- 2. PROCESAR DIAPOSITIVA DE GASOLINAS (DIAPOSITIVA 3) ---
        slide_gas = find_slide(slides_map, "snr", "gasolinas")
        chart_gas = None
        for shape in slide_gas.shapes:
            if shape.has_chart:
                chart_gas = shape.chart
            elif shape.has_text_frame and "CMP" in shape.text and "Demanda" in shape.text:
                cmp_val = getattr(app, 'cmp_gasolinas', "513")
                import re
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if "CMP" in run.text:
                            run.text = re.sub(r'CMP\s*:\s*[\d\.]+', f'CMP : {cmp_val}', run.text)

        if not chart_gas:
            raise ValueError("No se encontró una diapositiva con gráfica para Gasolinas (el título debe ser p. ej. 'Producción de Gasolinas SNR').")

        snr_col_gas = None
        for col in app.df_data_gasolinas.columns:
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
        for idx, row in app.df_prod_gasolinas.iterrows():
            cat = str(row.iloc[0]).strip()
            val = row.iloc[1]
            
            if not cat:
                continue
                
            if not any(c.isalpha() for c in cat):
                prod_rows_gas.append((cat, val))
            else:
                if val != "" and val is not None:
                    try:
                        float(val)
                        prod_rows_gas.append((cat, val))
                    except: pass

        # Ajustar al límite de 30 categorías
        if len(prod_rows_gas) > 30:
            prod_rows_gas = prod_rows_gas[-30:]

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
        for i in range(dias_mes):
            categories_gas.append(str(i + 1))
            proceso_vals_gas.append(None)
            
            d_val = None
            if i < len(app.df_data_gasolinas):
                try:
                    d_val = float(app.df_data_gasolinas[snr_col_gas].iloc[i])
                except:
                    d_val = None
            diario_vals_gas.append(d_val)
            
            p_val = None
            if i < len(app.df_snr_gasolinas):
                try:
                    p_val = float(app.df_snr_gasolinas.iloc[i, 0])
                except:
                    p_val = None
            programa_vals_gas.append(p_val)
            
            c_val = None
            if i < len(app.df_snr_gasolinas):
                try:
                    c_val = float(app.df_snr_gasolinas.iloc[i, 1])
                except:
                    c_val = None
            columna1_vals_gas.append(c_val)

        # Actualizar la gráfica de Gasolinas (Diapositiva 3)
        update_slide_chart(chart_gas, categories_gas, proceso_vals_gas, diario_vals_gas, programa_vals_gas, columna1_vals_gas, wine_color, green_color)


        # --- 3. PROCESAR DIAPOSITIVA DE DIESEL (DIAPOSITIVA 4) ---
        slide_die = find_slide(slides_map, "snr", "diesel")
        chart_die = None
        for shape in slide_die.shapes:
            if shape.has_chart:
                chart_die = shape.chart
            elif shape.has_text_frame and "CMP" in shape.text and "Demanda" in shape.text:
                cmp_val = getattr(app, 'cmp_diesel', "386.9")
                import re
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if "CMP" in run.text:
                            run.text = re.sub(r'CMP\s*:\s*[\d\.]+', f'CMP : {cmp_val}', run.text)

        if not chart_die:
            raise ValueError("No se encontró una diapositiva con gráfica para Diesel (el título debe ser p. ej. 'Producción de Diesel SNR').")

        snr_col_die = None
        for col in app.df_data_diesel.columns:
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
        for idx, row in app.df_prod_diesel.iterrows():
            cat = str(row.iloc[0]).strip()
            val = row.iloc[1]
            
            if not cat:
                continue
                
            if not any(c.isalpha() for c in cat):
                prod_rows_die.append((cat, val))
            else:
                if val != "" and val is not None:
                    try:
                        float(val)
                        prod_rows_die.append((cat, val))
                    except: pass

        # Ajustar al límite de 30 categorías
        if len(prod_rows_die) > 30:
            prod_rows_die = prod_rows_die[-30:]

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
        for i in range(dias_mes):
            categories_die.append(str(i + 1))
            proceso_vals_die.append(None)
            
            d_val = None
            if i < len(app.df_data_diesel):
                try:
                    d_val = float(app.df_data_diesel[snr_col_die].iloc[i])
                except:
                    d_val = None
            diario_vals_die.append(d_val)
            
            p_val = None
            if i < len(app.df_snr_diesel):
                try:
                    p_val = float(app.df_snr_diesel.iloc[i, 0])
                except:
                    p_val = None
            programa_vals_die.append(p_val)
            
            c_val = None
            if i < len(app.df_snr_diesel):
                try:
                    c_val = float(app.df_snr_diesel.iloc[i, 1])
                except:
                    c_val = None
            columna1_vals_die.append(c_val)

        # Actualizar la gráfica de Diesel (Diapositiva 4)
        update_slide_chart(chart_die, categories_die, proceso_vals_die, diario_vals_die, programa_vals_die, columna1_vals_die, wine_color, green_color)


        # --- 4. PROCESAR DIAPOSITIVA DE TURBOSINA (DIAPOSITIVA 5) ---
        slide_turb = find_slide(slides_map, "snr", "turbosina")
        chart_turb = None
        for shape in slide_turb.shapes:
            if shape.has_chart:
                chart_turb = shape.chart
            elif shape.has_text_frame and "CMP" in shape.text and "Demanda" in shape.text:
                cmp_val = getattr(app, 'cmp_turbosina', "312.4")
                import re
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if "CMP" in run.text:
                            run.text = re.sub(r'CMP\s*:\s*[\d\.]+', f'CMP : {cmp_val}', run.text)

        if not chart_turb:
            raise ValueError("No se encontró una diapositiva con gráfica para Turbosina (el título debe ser p. ej. 'Producción de Turbosina SNR').")

        snr_col_turb = None
        for col in app.df_data_turbosina.columns:
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
        for idx, row in app.df_prod_turbosina.iterrows():
            cat = str(row.iloc[0]).strip()
            val = row.iloc[1]
            
            if not cat:
                continue
                
            if not any(c.isalpha() for c in cat):
                prod_rows_turb.append((cat, val))
            else:
                if val != "" and val is not None:
                    try:
                        float(val)
                        prod_rows_turb.append((cat, val))
                    except: pass

        # Ajustar al límite de 30 categorías
        if len(prod_rows_turb) > 30:
            prod_rows_turb = prod_rows_turb[-30:]

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
        for i in range(dias_mes):
            categories_turb.append(str(i + 1))
            proceso_vals_turb.append(None)
            
            d_val = None
            if i < len(app.df_data_turbosina):
                try:
                    d_val = float(app.df_data_turbosina[snr_col_turb].iloc[i])
                except:
                    d_val = None
            diario_vals_turb.append(d_val)
            
            p_val = None
            if i < len(app.df_snr_turbosina):
                try:
                    p_val = float(app.df_snr_turbosina.iloc[i, 0])
                except:
                    p_val = None
            programa_vals_turb.append(p_val)
            
            c_val = None
            if i < len(app.df_snr_turbosina):
                try:
                    c_val = float(app.df_snr_turbosina.iloc[i, 1])
                except:
                    c_val = None
            columna1_vals_turb.append(c_val)

        # Actualizar la gráfica de Turbosina (Diapositiva 5)
        update_slide_chart(chart_turb, categories_turb, proceso_vals_turb, diario_vals_turb, programa_vals_turb, columna1_vals_turb, wine_color, green_color)


        # --- 5. PROCESAR DIAPOSITIVA DE ASFALTO (DIAPOSITIVA 6) ---
        if app.df_data_asfalto is not None and app.df_snr_asfalto is not None and app.df_prod_asfalto is not None:
            slide_asf = find_slide(slides_map, "snr", "asfalto")
            chart_asf = None
            for shape in slide_asf.shapes:
                if shape.has_chart:
                    chart_asf = shape.chart
                elif shape.has_text_frame and "CMP" in shape.text and "Demanda" in shape.text:
                    cmp_val = getattr(app, 'cmp_asfalto', "21.4")
                    import re
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "CMP" in run.text:
                                run.text = re.sub(r'CMP\s*:\s*[\d\.]+', f'CMP : {cmp_val}', run.text)

            if not chart_asf:
                raise ValueError("No se encontró una diapositiva con gráfica para Asfalto (el título debe ser p. ej. 'Producción de Asfalto SNR').")

            snr_col_asf = None
            for col in app.df_data_asfalto.columns:
                if "SNR" in str(col).upper() or "REAL" in str(col).upper():
                    snr_col_asf = col
                    break
            if not snr_col_asf and len(app.df_data_asfalto.columns) > 1:
                snr_col_asf = app.df_data_asfalto.columns[1]

            if not snr_col_asf:
                raise ValueError("No se encontró la columna de producción diaria ('REAL') en la tabla de Asfalto.")

            categories_asf = []
            proceso_vals_asf = []
            diario_vals_asf = []
            programa_vals_asf = []
            columna1_vals_asf = []

            # Filtrar df_prod_asfalto para quedarnos con los años y meses activos
            prod_rows_asf = []
            for idx, row in app.df_prod_asfalto.iterrows():
                cat = str(row.iloc[0]).strip()
                val = row.iloc[1]

                if not cat:
                    continue

                if not any(c.isalpha() for c in cat):
                    prod_rows_asf.append((cat, val))
                else:
                    if val != "" and val is not None:
                        try:
                            float(val)
                            prod_rows_asf.append((cat, val))
                        except: pass

            # Ajustar al límite de 30 categorías
            if len(prod_rows_asf) > 30:
                prod_rows_asf = prod_rows_asf[-30:]

            # Llenar filas de años y meses (sin celdas vacías al final)
            for i in range(len(prod_rows_asf)):
                cat_val = prod_rows_asf[i][0]
                try:
                    proc_val = float(prod_rows_asf[i][1])
                except:
                    proc_val = None

                categories_asf.append(cat_val)
                proceso_vals_asf.append(proc_val)
                diario_vals_asf.append(None)
                programa_vals_asf.append(None)
                columna1_vals_asf.append(None)

            # Llenar filas diarias
            for i in range(dias_mes):
                categories_asf.append(str(i + 1))
                proceso_vals_asf.append(None)

                d_val = None
                if i < len(app.df_data_asfalto):
                    try:
                        d_val = float(app.df_data_asfalto[snr_col_asf].iloc[i])
                    except:
                        d_val = None
                diario_vals_asf.append(d_val)

                p_val = None
                if i < len(app.df_snr_asfalto):
                    try:
                        p_val = float(app.df_snr_asfalto.iloc[i, 0])
                    except:
                        p_val = None
                programa_vals_asf.append(p_val)

                c_val = None
                if i < len(app.df_snr_asfalto):
                    try:
                        c_val = float(app.df_snr_asfalto.iloc[i, 1])
                    except:
                        c_val = None
                columna1_vals_asf.append(c_val)

            # Actualizar la gráfica de Asfalto (Diapositiva 6)
            update_slide_chart(chart_asf, categories_asf, proceso_vals_asf, diario_vals_asf, programa_vals_asf, columna1_vals_asf, wine_color, green_color, current_month_name=master_current_month)


        # --- 6. PROCESAR DIAPOSITIVA DE COMBUSTOLEO (DIAPOSITIVA 7) ---
        if app.df_data_combustoleo is not None and app.df_snr_combustoleo is not None and app.df_prod_combustoleo is not None:
            slide_comb = find_slide(slides_map, "snr", "combustoleo")
            chart_comb = None
            for shape in slide_comb.shapes:
                if shape.has_chart:
                    chart_comb = shape.chart
                elif shape.has_text_frame and "CMP" in shape.text and "Demanda" in shape.text:
                    cmp_val = getattr(app, 'cmp_combustoleo', "250.9")
                    import re
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "CMP" in run.text:
                                run.text = re.sub(r'CMP\s*:\s*[\d\.]+', f'CMP : {cmp_val}', run.text)

            if not chart_comb:
                raise ValueError("No se encontró una diapositiva con gráfica para Combustoleo (el título debe ser p. ej. 'Producción de Combustoleo SNR').")

            # La columna de producción diaria real es "Combustoleo" (no la de SNR)
            diario_col_comb = None
            for col in app.df_data_combustoleo.columns:
                if "SNR" not in str(col).upper() and "REAL" not in str(col).upper():
                    diario_col_comb = col
                    break
            if not diario_col_comb and len(app.df_data_combustoleo.columns) > 1:
                diario_col_comb = app.df_data_combustoleo.columns[1]
            elif not diario_col_comb:
                diario_col_comb = app.df_data_combustoleo.columns[0]

            if not diario_col_comb:
                raise ValueError("No se encontró la columna de producción diaria ('Combustoleo') en la tabla de Combustoleo.")

            categories_comb = []
            proceso_vals_comb = []
            diario_vals_comb = []
            programa_vals_comb = []
            columna1_vals_comb = []

            # Filtrar df_prod_combustoleo para quedarnos con los años y meses activos
            prod_rows_comb = []
            for idx, row in app.df_prod_combustoleo.iterrows():
                cat = str(row.iloc[0]).strip()
                val = row.iloc[1]

                if not cat:
                    continue

                if not any(c.isalpha() for c in cat):
                    prod_rows_comb.append((cat, val))
                else:
                    if val != "" and val is not None:
                        try:
                            float(val)
                            prod_rows_comb.append((cat, val))
                        except: pass

            # Ajustar al límite de 30 categorías
            if len(prod_rows_comb) > 30:
                prod_rows_comb = prod_rows_comb[-30:]

            # Llenar filas de años y meses (sin celdas vacías al final)
            for i in range(len(prod_rows_comb)):
                cat_val = prod_rows_comb[i][0]
                try:
                    proc_val = float(prod_rows_comb[i][1])
                except:
                    proc_val = None

                categories_comb.append(cat_val)
                proceso_vals_comb.append(proc_val)
                diario_vals_comb.append(None)
                programa_vals_comb.append(None)
                columna1_vals_comb.append(None)

            # Llenar filas diarias
            for i in range(dias_mes):
                categories_comb.append(str(i + 1))
                proceso_vals_comb.append(None)

                d_val = None
                if i < len(app.df_data_combustoleo):
                    try:
                        d_val = float(app.df_data_combustoleo[diario_col_comb].iloc[i])
                    except:
                        d_val = None
                diario_vals_comb.append(d_val)

                p_val = None
                if i < len(app.df_snr_combustoleo):
                    try:
                        p_val = float(app.df_snr_combustoleo.iloc[i, 0])
                    except:
                        p_val = None
                programa_vals_comb.append(p_val)

                c_val = None
                if i < len(app.df_snr_combustoleo):
                    try:
                        c_val = float(app.df_snr_combustoleo.iloc[i, 1])
                    except:
                        c_val = None
                columna1_vals_comb.append(c_val)

            # Actualizar la gráfica de Combustoleo (Diapositiva 7)
            update_slide_chart(chart_comb, categories_comb, proceso_vals_comb, diario_vals_comb, programa_vals_comb, columna1_vals_comb, wine_color, green_color, current_month_name=master_current_month)
 
            # --- 7. PROCESAR DIAPOSITIVA DE GASOLINAS CADEREYTA (DIAPOSITIVA 11) ---
            if app.df_data_cad_gas is not None and app.df_snr_cad_gas is not None and app.df_prod_cad_gas is not None:
                slide_cad_gas = find_slide(slides_map, "cadereyta", "gasolinas")
                chart_cad_gas = None
                for shape in slide_cad_gas.shapes:
                    if shape.has_chart:
                        chart_cad_gas = shape.chart
                        break
                
                if chart_cad_gas:
                    # Buscar columna SNR o usar la segunda por defecto
                    snr_col_cad_gas = None
                    for col in app.df_data_cad_gas.columns:
                        if "SNR" in str(col).upper():
                            snr_col_cad_gas = col
                            break
                    if not snr_col_cad_gas and len(app.df_data_cad_gas.columns) >= 2:
                        snr_col_cad_gas = app.df_data_cad_gas.columns[1]
                    
                    if snr_col_cad_gas:
                        categories_cg = []
                        proceso_vals_cg = []
                        diario_vals_cg = []
                        programa_vals_cg = []
                        columna1_vals_cg = []

                        # Filtrar producción anual (últimas 30 categorías)
                        prod_rows_cg = []
                        for idx, row in app.df_prod_cad_gas.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_cg.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_cg.append((cat, val))
                                    except: pass





                        if len(prod_rows_cg) > 30: prod_rows_cg = prod_rows_cg[-30:]

                        for i in range(len(prod_rows_cg)):
                            categories_cg.append(prod_rows_cg[i][0])
                            try: proceso_vals_cg.append(float(prod_rows_cg[i][1]))
                            except: proceso_vals_cg.append(None)
                            diario_vals_cg.append(None)
                            programa_vals_cg.append(None)
                            columna1_vals_cg.append(None)
 
                        # Llenar datos diarios (días del mes en curso)
                        for i in range(dias_mes):
                            categories_cg.append(str(i + 1))
                            proceso_vals_cg.append(None)
                            
                            try: diario_vals_cg.append(float(app.df_data_cad_gas[snr_col_cad_gas].iloc[i]))
                            except: diario_vals_cg.append(None)
                            
                            try: programa_vals_cg.append(float(app.df_snr_cad_gas.iloc[i, 0]))
                            except: programa_vals_cg.append(None)
                            
                            try: columna1_vals_cg.append(float(app.df_snr_cad_gas.iloc[i, 1]))
                            except: columna1_vals_cg.append(None)
 
                        update_slide_chart(chart_cad_gas, categories_cg, proceso_vals_cg, diario_vals_cg, programa_vals_cg, columna1_vals_cg, wine_color, green_color)
 
                # --- 8. PROCESAR DIAPOSITIVA DE DIESEL CADEREYTA (DIAPOSITIVA 12) ---
                if app.df_data_cad_die is not None and app.df_snr_cad_die is not None and app.df_prod_cad_die is not None:
                    slide_cad_die = find_slide(slides_map, "cadereyta", "diesel")
                    chart_cad_die = None
                    for shape in slide_cad_die.shapes:
                        if shape.has_chart:
                            chart_cad_die = shape.chart
                            break
                    
                    if chart_cad_die:
                        snr_col_cad_die = None
                        for col in app.df_data_cad_die.columns:
                            if "SNR" in str(col).upper():
                                snr_col_cad_die = col
                                break
                        if not snr_col_cad_die and len(app.df_data_cad_die.columns) >= 2:
                            snr_col_cad_die = app.df_data_cad_die.columns[1]
                        
                        if snr_col_cad_die:
                            categories_cd = []
                            proceso_vals_cd = []
                            diario_vals_cd = []
                            programa_vals_cd = []
                            columna1_vals_cd = []
 
                            prod_rows_cd = []
                            for idx, row in app.df_prod_cad_die.iterrows():
                                cat = str(row.iloc[0]).strip()
                                val = row.iloc[1]
                                if not cat: continue
                                if not any(c.isalpha() for c in cat):
                                    prod_rows_cd.append((cat, val))
                                else:
                                    if val != "" and val is not None:
                                        try:
                                            float(val)
                                            prod_rows_cd.append((cat, val))
                                        except: pass
                            
                            if len(prod_rows_cd) > 30: prod_rows_cd = prod_rows_cd[-30:]
 
                            for i in range(len(prod_rows_cd)):
                                categories_cd.append(prod_rows_cd[i][0])
                                try: proceso_vals_cd.append(float(prod_rows_cd[i][1]))
                                except: proceso_vals_cd.append(None)
                                diario_vals_cd.append(None)
                                programa_vals_cd.append(None)
                                columna1_vals_cd.append(None)
 
                            # Llenar datos diarios (días del mes en curso)
                            for i in range(dias_mes):
                                categories_cd.append(str(i + 1))
                                proceso_vals_cd.append(None)
                                
                                try: diario_vals_cd.append(float(app.df_data_cad_die[snr_col_cad_die].iloc[i]))
                                except: diario_vals_cd.append(None)
                                
                                try: programa_vals_cd.append(float(app.df_snr_cad_die.iloc[i, 0]))
                                except: programa_vals_cd.append(None)
                                
                                try: columna1_vals_cd.append(float(app.df_snr_cad_die.iloc[i, 1]))
                                except: columna1_vals_cd.append(None)
 
                            update_slide_chart(chart_cad_die, categories_cd, proceso_vals_cd, diario_vals_cd, programa_vals_cd, columna1_vals_cd, wine_color, green_color)

                # --- 9. PROCESAR DIAPOSITIVA DE COMBUSTOLEO CADEREYTA (DIAPOSITIVA 13) ---
                if slide_exists(slides_map, "cadereyta", "combustoleo") and app.df_data_cad_comb is not None and app.df_snr_cad_comb is not None and app.df_prod_cad_comb is not None:
                    slide_cad_comb = find_slide(slides_map, "cadereyta", "combustoleo")
                    chart_cad_comb = None
                    for shape in slide_cad_comb.shapes:
                        if shape.has_chart:
                            chart_cad_comb = shape.chart
                            break
                    
                    if chart_cad_comb:
                        snr_col_cad_comb = None
                        for col in app.df_data_cad_comb.columns:
                            if "SNR" in str(col).upper():
                                snr_col_cad_comb = col
                                break
                        if not snr_col_cad_comb and len(app.df_data_cad_comb.columns) >= 2:
                            snr_col_cad_comb = app.df_data_cad_comb.columns[1]
                        
                        if snr_col_cad_comb:
                            categories_cc = []
                            proceso_vals_cc = []
                            diario_vals_cc = []
                            programa_vals_cc = []
                            columna1_vals_cc = []
 
                            prod_rows_cc = []
                            for idx, row in app.df_prod_cad_comb.iterrows():
                                cat = str(row.iloc[0]).strip()
                                val = row.iloc[1]
                                if not cat: continue
                                if not any(c.isalpha() for c in cat):
                                    prod_rows_cc.append((cat, val))
                                else:
                                    if val != "" and val is not None:
                                        try:
                                            float(val)
                                            prod_rows_cc.append((cat, val))
                                        except: pass
                            
                            if len(prod_rows_cc) > 30: prod_rows_cc = prod_rows_cc[-30:]
 
                            for i in range(len(prod_rows_cc)):
                                categories_cc.append(prod_rows_cc[i][0])
                                try: proceso_vals_cc.append(float(prod_rows_cc[i][1]))
                                except: proceso_vals_cc.append(None)
                                diario_vals_cc.append(None)
                                programa_vals_cc.append(None)
                                columna1_vals_cc.append(None)
 
                            # Llenar datos diarios (días del mes en curso)
                            for i in range(dias_mes):
                                categories_cc.append(str(i + 1))
                                proceso_vals_cc.append(None)
                                
                                try: diario_vals_cc.append(float(app.df_data_cad_comb[snr_col_cad_comb].iloc[i]))
                                except: diario_vals_cc.append(None)
                                
                                try: programa_vals_cc.append(float(app.df_snr_cad_comb.iloc[i, 0]))
                                except: programa_vals_cc.append(None)
                                
                                try: columna1_vals_cc.append(float(app.df_snr_cad_comb.iloc[i, 1]))
                                except: columna1_vals_cc.append(None)
 
                            update_slide_chart(chart_cad_comb, categories_cc, proceso_vals_cc, diario_vals_cc, programa_vals_cc, columna1_vals_cc, wine_color, green_color)

            # --- 10. PROCESAR DIAPOSITIVA DE CRUDO MADERO (DIAPOSITIVA 16) ---
            if slide_exists(slides_map, "madero", "crudo") and app.df_data_mad_crud is not None and app.df_snr_mad_crud is not None and app.df_prod_mad_crud is not None:
                slide_mad_crud = find_slide(slides_map, "madero", "crudo")
                chart_mad_crud = None
                for shape in slide_mad_crud.shapes:
                    if shape.has_chart:
                        chart_mad_crud = shape.chart
                        break
                
                if chart_mad_crud:
                    categories_mc = []
                    proceso_vals_mc = []
                    diario_vals_mc = []
                    programa_vals_mc = []
                    columna1_vals_mc = []

                    prod_rows_mc = []
                    for idx, row in app.df_prod_mad_crud.iterrows():
                        cat = str(row.iloc[0]).strip()
                        val = row.iloc[1]
                        if not cat: continue
                        if not any(c.isalpha() for c in cat):
                            prod_rows_mc.append((cat, val))
                        else:
                            if val != "" and val is not None:
                                try:
                                    float(val)
                                    prod_rows_mc.append((cat, val))
                                except: pass

                    if len(prod_rows_mc) > 30: prod_rows_mc = prod_rows_mc[-30:]

                    for i in range(len(prod_rows_mc)):
                        categories_mc.append(prod_rows_mc[i][0])
                        try: proceso_vals_mc.append(float(prod_rows_mc[i][1]))
                        except: proceso_vals_mc.append(None)
                        diario_vals_mc.append(None)
                        programa_vals_mc.append(None)
                        columna1_vals_mc.append(None)

                    proceso_col_mad_crud = app.df_data_mad_crud.columns[1]

                    for i in range(dias_mes):
                        categories_mc.append(str(i + 1))
                        proceso_vals_mc.append(None)
                        
                        try: diario_vals_mc.append(float(app.df_data_mad_crud[proceso_col_mad_crud].iloc[i]))
                        except: diario_vals_mc.append(None)
                        
                        try: programa_vals_mc.append(float(app.df_snr_mad_crud.iloc[i, 0]))
                        except: programa_vals_mc.append(None)
                        
                        try: columna1_vals_mc.append(float(app.df_snr_mad_crud.iloc[i, 1]))
                        except: columna1_vals_mc.append(None)

                    update_slide_chart(chart_mad_crud, categories_mc, proceso_vals_mc, diario_vals_mc, programa_vals_mc, columna1_vals_mc, wine_color, green_color)

            # --- 11. PROCESAR DIAPOSITIVA DE GASOLINAS MADERO (DIAPOSITIVA 17) ---
            if slide_exists(slides_map, "madero", "gasolinas") and app.df_data_mad_gas is not None and app.df_snr_mad_gas is not None and app.df_prod_mad_gas is not None:
                slide_mad_gas = find_slide(slides_map, "madero", "gasolinas")
                chart_mad_gas = None
                for shape in slide_mad_gas.shapes:
                    if shape.has_chart:
                        chart_mad_gas = shape.chart
                        break
                
                if chart_mad_gas:
                    snr_col_mad_gas = None
                    for col in app.df_data_mad_gas.columns:
                        if "SNR" in str(col).upper():
                            snr_col_mad_gas = col
                            break
                    if not snr_col_mad_gas and len(app.df_data_mad_gas.columns) >= 2:
                        snr_col_mad_gas = app.df_data_mad_gas.columns[1]

                    if snr_col_mad_gas:
                        categories_mg = []
                        proceso_vals_mg = []
                        diario_vals_mg = []
                        programa_vals_mg = []
                        columna1_vals_mg = []

                        prod_rows_mg = []
                        for idx, row in app.df_prod_mad_gas.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_mg.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_mg.append((cat, val))
                                    except: pass

                        if len(prod_rows_mg) > 30: prod_rows_mg = prod_rows_mg[-30:]

                        for i in range(len(prod_rows_mg)):
                            categories_mg.append(prod_rows_mg[i][0])
                            try: proceso_vals_mg.append(float(prod_rows_mg[i][1]))
                            except: proceso_vals_mg.append(None)
                            diario_vals_mg.append(None)
                            programa_vals_mg.append(None)
                            columna1_vals_mg.append(None)

                        for i in range(dias_mes):
                            categories_mg.append(str(i + 1))
                            proceso_vals_mg.append(None)
                            
                            try: diario_vals_mg.append(float(app.df_data_mad_gas[snr_col_mad_gas].iloc[i]))
                            except: diario_vals_mg.append(None)
                            
                            try: programa_vals_mg.append(float(app.df_snr_mad_gas.iloc[i, 0]))
                            except: programa_vals_mg.append(None)
                            
                            try: columna1_vals_mg.append(float(app.df_snr_mad_gas.iloc[i, 1]))
                            except: columna1_vals_mg.append(None)

                        update_slide_chart(chart_mad_gas, categories_mg, proceso_vals_mg, diario_vals_mg, programa_vals_mg, columna1_vals_mg, wine_color, green_color)

            # --- 12. PROCESAR DIAPOSITIVA DE DIESEL MADERO (DIAPOSITIVA 18) ---
            if slide_exists(slides_map, "madero", "diesel") and app.df_data_mad_die is not None and app.df_snr_mad_die is not None and app.df_prod_mad_die is not None:
                slide_mad_die = find_slide(slides_map, "madero", "diesel")
                chart_mad_die = None
                for shape in slide_mad_die.shapes:
                    if shape.has_chart:
                        chart_mad_die = shape.chart
                        break
                
                if chart_mad_die:
                    snr_col_mad_die = None
                    for col in app.df_data_mad_die.columns:
                        if "SNR" in str(col).upper():
                            snr_col_mad_die = col
                            break
                    if not snr_col_mad_die and len(app.df_data_mad_die.columns) >= 2:
                        snr_col_mad_die = app.df_data_mad_die.columns[1]

                    if snr_col_mad_die:
                        categories_md = []
                        proceso_vals_md = []
                        diario_vals_md = []
                        programa_vals_md = []
                        columna1_vals_md = []

                        prod_rows_md = []
                        for idx, row in app.df_prod_mad_die.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_md.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_md.append((cat, val))
                                    except: pass

                        if len(prod_rows_md) > 30: prod_rows_md = prod_rows_md[-30:]

                        for i in range(len(prod_rows_md)):
                            categories_md.append(prod_rows_md[i][0])
                            try: proceso_vals_md.append(float(prod_rows_md[i][1]))
                            except: proceso_vals_md.append(None)
                            diario_vals_md.append(None)
                            programa_vals_md.append(None)
                            columna1_vals_md.append(None)

                        for i in range(dias_mes):
                            categories_md.append(str(i + 1))
                            proceso_vals_md.append(None)
                            
                            try: diario_vals_md.append(float(app.df_data_mad_die[snr_col_mad_die].iloc[i]))
                            except: diario_vals_md.append(None)
                            
                            try: programa_vals_md.append(float(app.df_snr_mad_die.iloc[i, 0]))
                            except: programa_vals_md.append(None)
                            
                            try: columna1_vals_md.append(float(app.df_snr_mad_die.iloc[i, 1]))
                            except: columna1_vals_md.append(None)

                        update_slide_chart(chart_mad_die, categories_md, proceso_vals_md, diario_vals_md, programa_vals_md, columna1_vals_md, wine_color, green_color)

            # --- 13. PROCESAR DIAPOSITIVA DE TURBOSINA MADERO (DIAPOSITIVA 19) ---
            if slide_exists(slides_map, "madero", "turbosina") and app.df_data_mad_turb is not None and app.df_snr_mad_turb is not None and app.df_prod_mad_turb is not None:
                slide_mad_turb = find_slide(slides_map, "madero", "turbosina")
                chart_mad_turb = None
                for shape in slide_mad_turb.shapes:
                    if shape.has_chart:
                        chart_mad_turb = shape.chart
                        break
                
                if chart_mad_turb:
                    snr_col_mad_turb = None
                    for col in app.df_data_mad_turb.columns:
                        if "SNR" in str(col).upper():
                            snr_col_mad_turb = col
                            break
                    if not snr_col_mad_turb and len(app.df_data_mad_turb.columns) >= 2:
                        snr_col_mad_turb = app.df_data_mad_turb.columns[1]

                    if snr_col_mad_turb:
                        categories_mtu = []
                        proceso_vals_mtu = []
                        diario_vals_mtu = []
                        programa_vals_mtu = []
                        columna1_vals_mtu = []

                        prod_rows_mtu = []
                        for idx, row in app.df_prod_mad_turb.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_mtu.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_mtu.append((cat, val))
                                    except: pass

                        if len(prod_rows_mtu) > 30: prod_rows_mtu = prod_rows_mtu[-30:]

                        for i in range(len(prod_rows_mtu)):
                            categories_mtu.append(prod_rows_mtu[i][0])
                            try: proceso_vals_mtu.append(float(prod_rows_mtu[i][1]))
                            except: proceso_vals_mtu.append(None)
                            diario_vals_mtu.append(None)
                            programa_vals_mtu.append(None)
                            columna1_vals_mtu.append(None)

                        for i in range(dias_mes):
                            categories_mtu.append(str(i + 1))
                            proceso_vals_mtu.append(None)
                            
                            try: diario_vals_mtu.append(float(app.df_data_mad_turb[snr_col_mad_turb].iloc[i]))
                            except: diario_vals_mtu.append(None)
                            
                            try: programa_vals_mtu.append(float(app.df_snr_mad_turb.iloc[i, 0]))
                            except: programa_vals_mtu.append(None)
                            
                            try: columna1_vals_mtu.append(float(app.df_snr_mad_turb.iloc[i, 1]))
                            except: columna1_vals_mtu.append(None)

                        update_slide_chart(chart_mad_turb, categories_mtu, proceso_vals_mtu, diario_vals_mtu, programa_vals_mtu, columna1_vals_mtu, wine_color, green_color)

            # --- 14. PROCESAR DIAPOSITIVA DE COMBUSTOLEO MADERO (DIAPOSITIVA 20) ---
            if slide_exists(slides_map, "madero", "combustoleo") and app.df_data_mad_comb is not None and app.df_snr_mad_comb is not None and app.df_prod_mad_comb is not None:
                slide_mad_comb = find_slide(slides_map, "madero", "combustoleo")
                chart_mad_comb = None
                for shape in slide_mad_comb.shapes:
                    if shape.has_chart:
                        chart_mad_comb = shape.chart
                        break
                
                if chart_mad_comb:
                    snr_col_mad_comb = None
                    for col in app.df_data_mad_comb.columns:
                        if "SNR" in str(col).upper():
                            snr_col_mad_comb = col
                            break
                    if not snr_col_mad_comb and len(app.df_data_mad_comb.columns) >= 2:
                        snr_col_mad_comb = app.df_data_mad_comb.columns[1]

                    if snr_col_mad_comb:
                        categories_mco = []
                        proceso_vals_mco = []
                        diario_vals_mco = []
                        programa_vals_mco = []
                        columna1_vals_mco = []

                        prod_rows_mco = []
                        for idx, row in app.df_prod_mad_comb.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_mco.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_mco.append((cat, val))
                                    except: pass

                        if len(prod_rows_mco) > 30: prod_rows_mco = prod_rows_mco[-30:]

                        for i in range(len(prod_rows_mco)):
                            categories_mco.append(prod_rows_mco[i][0])
                            try: proceso_vals_mco.append(float(prod_rows_mco[i][1]))
                            except: proceso_vals_mco.append(None)
                            diario_vals_mco.append(None)
                            programa_vals_mco.append(None)
                            columna1_vals_mco.append(None)

                        for i in range(dias_mes):
                            categories_mco.append(str(i + 1))
                            proceso_vals_mco.append(None)
                            
                            try: diario_vals_mco.append(float(app.df_data_mad_comb[snr_col_mad_comb].iloc[i]))
                            except: diario_vals_mco.append(None)
                            
                            try: programa_vals_mco.append(float(app.df_snr_mad_comb.iloc[i, 0]))
                            except: programa_vals_mco.append(None)
                            
                            try: columna1_vals_mco.append(float(app.df_snr_mad_comb.iloc[i, 1]))
                            except: columna1_vals_mco.append(None)

                        update_slide_chart(chart_mad_comb, categories_mco, proceso_vals_mco, diario_vals_mco, programa_vals_mco, columna1_vals_mco, wine_color, green_color)

            # --- 15. PROCESAR DIAPOSITIVA DE CRUDO MINATITLAN (DIAPOSITIVA 23) ---
            if slide_exists(slides_map, "minatitlan", "crudo") and app.df_data_mina_crud is not None and app.df_snr_mina_crud is not None and app.df_prod_mina_crud is not None:
                slide_mina_crud = find_slide(slides_map, "minatitlan", "crudo")
                chart_mina_crud = None
                for shape in slide_mina_crud.shapes:
                    if shape.has_chart:
                        chart_mina_crud = shape.chart
                        break
                
                if chart_mina_crud:
                    snr_col_mina_crud = None
                    for col in app.df_data_mina_crud.columns:
                        if "SNR" in str(col).upper():
                            snr_col_mina_crud = col
                            break
                    if not snr_col_mina_crud and len(app.df_data_mina_crud.columns) >= 2:
                        snr_col_mina_crud = app.df_data_mina_crud.columns[1]

                    if snr_col_mina_crud:
                        categories_mic = []
                        proceso_vals_mic = []
                        diario_vals_mic = []
                        programa_vals_mic = []
                        columna1_vals_mic = []

                        prod_rows_mic = []
                        for idx, row in app.df_prod_mina_crud.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_mic.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_mic.append((cat, val))
                                    except: pass

                        if len(prod_rows_mic) > 30: prod_rows_mic = prod_rows_mic[-30:]

                        for i in range(len(prod_rows_mic)):
                            categories_mic.append(prod_rows_mic[i][0])
                            try: proceso_vals_mic.append(float(prod_rows_mic[i][1]))
                            except: proceso_vals_mic.append(None)
                            diario_vals_mic.append(None)
                            programa_vals_mic.append(None)
                            columna1_vals_mic.append(None)

                        for i in range(dias_mes):
                            categories_mic.append(str(i + 1))
                            proceso_vals_mic.append(None)
                            
                            try: diario_vals_mic.append(float(app.df_data_mina_crud[snr_col_mina_crud].iloc[i]))
                            except: diario_vals_mic.append(None)
                            
                            try: programa_vals_mic.append(float(app.df_snr_mina_crud.iloc[i, 0]))
                            except: programa_vals_mic.append(None)
                            
                            try: columna1_vals_mic.append(float(app.df_snr_mina_crud.iloc[i, 1]))
                            except: columna1_vals_mic.append(None)

                        update_slide_chart(chart_mina_crud, categories_mic, proceso_vals_mic, diario_vals_mic, programa_vals_mic, columna1_vals_mic, wine_color, green_color)

            # --- 16. PROCESAR DIAPOSITIVA DE GASOLINAS MINATITLAN (DIAPOSITIVA 24) ---
            if slide_exists(slides_map, "minatitlan", "gasolinas") and app.df_data_mina_gas is not None and app.df_snr_mina_gas is not None and app.df_prod_mina_gas is not None:
                slide_mina_gas = find_slide(slides_map, "minatitlan", "gasolinas")
                chart_mina_gas = None
                for shape in slide_mina_gas.shapes:
                    if shape.has_chart:
                        chart_mina_gas = shape.chart
                        break
                
                if chart_mina_gas:
                    snr_col_mina_gas = None
                    for col in app.df_data_mina_gas.columns:
                        if "SNR" in str(col).upper():
                            snr_col_mina_gas = col
                            break
                    if not snr_col_mina_gas and len(app.df_data_mina_gas.columns) >= 2:
                        snr_col_mina_gas = app.df_data_mina_gas.columns[1]

                    if snr_col_mina_gas:
                        categories_mig = []
                        proceso_vals_mig = []
                        diario_vals_mig = []
                        programa_vals_mig = []
                        columna1_vals_mig = []

                        prod_rows_mig = []
                        for idx, row in app.df_prod_mina_gas.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_mig.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_mig.append((cat, val))
                                    except: pass

                        if len(prod_rows_mig) > 30: prod_rows_mig = prod_rows_mig[-30:]

                        for i in range(len(prod_rows_mig)):
                            categories_mig.append(prod_rows_mig[i][0])
                            try: proceso_vals_mig.append(float(prod_rows_mig[i][1]))
                            except: proceso_vals_mig.append(None)
                            diario_vals_mig.append(None)
                            programa_vals_mig.append(None)
                            columna1_vals_mig.append(None)

                        for i in range(dias_mes):
                            categories_mig.append(str(i + 1))
                            proceso_vals_mig.append(None)
                            
                            try: diario_vals_mig.append(float(app.df_data_mina_gas[snr_col_mina_gas].iloc[i]))
                            except: diario_vals_mig.append(None)
                            
                            try: programa_vals_mig.append(float(app.df_snr_mina_gas.iloc[i, 0]))
                            except: programa_vals_mig.append(None)
                            
                            try: columna1_vals_mig.append(float(app.df_snr_mina_gas.iloc[i, 1]))
                            except: columna1_vals_mig.append(None)

                        update_slide_chart(chart_mina_gas, categories_mig, proceso_vals_mig, diario_vals_mig, programa_vals_mig, columna1_vals_mig, wine_color, green_color)

            # --- 17. PROCESAR DIAPOSITIVA DE DIESEL MINATITLAN (DIAPOSITIVA 25) ---
            if slide_exists(slides_map, "minatitlan", "diesel") and app.df_data_mina_die is not None and app.df_snr_mina_die is not None and app.df_prod_mina_die is not None:
                slide_mina_die = find_slide(slides_map, "minatitlan", "diesel")
                chart_mina_die = None
                for shape in slide_mina_die.shapes:
                    if shape.has_chart:
                        chart_mina_die = shape.chart
                        break
                
                if chart_mina_die:
                    snr_col_mina_die = None
                    for col in app.df_data_mina_die.columns:
                        if "SNR" in str(col).upper():
                            snr_col_mina_die = col
                            break
                    if not snr_col_mina_die and len(app.df_data_mina_die.columns) >= 2:
                        snr_col_mina_die = app.df_data_mina_die.columns[1]

                    if snr_col_mina_die:
                        categories_mid = []
                        proceso_vals_mid = []
                        diario_vals_mid = []
                        programa_vals_mid = []
                        columna1_vals_mid = []

                        prod_rows_mid = []
                        for idx, row in app.df_prod_mina_die.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_mid.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_mid.append((cat, val))
                                    except: pass

                        if len(prod_rows_mid) > 30: prod_rows_mid = prod_rows_mid[-30:]

                        for i in range(len(prod_rows_mid)):
                            categories_mid.append(prod_rows_mid[i][0])
                            try: proceso_vals_mid.append(float(prod_rows_mid[i][1]))
                            except: proceso_vals_mid.append(None)
                            diario_vals_mid.append(None)
                            programa_vals_mid.append(None)
                            columna1_vals_mid.append(None)

                        for i in range(dias_mes):
                            categories_mid.append(str(i + 1))
                            proceso_vals_mid.append(None)
                            
                            try: diario_vals_mid.append(float(app.df_data_mina_die[snr_col_mina_die].iloc[i]))
                            except: diario_vals_mid.append(None)
                            
                            try: programa_vals_mid.append(float(app.df_snr_mina_die.iloc[i, 0]))
                            except: programa_vals_mid.append(None)
                            
                            try: columna1_vals_mid.append(float(app.df_snr_mina_die.iloc[i, 1]))
                            except: columna1_vals_mid.append(None)

                        update_slide_chart(chart_mina_die, categories_mid, proceso_vals_mid, diario_vals_mid, programa_vals_mid, columna1_vals_mid, wine_color, green_color)

            # --- 18. PROCESAR DIAPOSITIVA DE COMBUSTOLEO MINATITLAN (DIAPOSITIVA 26) ---
            if slide_exists(slides_map, "minatitlan", "combustoleo") and app.df_data_mina_comb is not None and app.df_snr_mina_comb is not None and app.df_prod_mina_comb is not None:
                slide_mina_comb = find_slide(slides_map, "minatitlan", "combustoleo")
                chart_mina_comb = None
                for shape in slide_mina_comb.shapes:
                    if shape.has_chart:
                        chart_mina_comb = shape.chart
                        break
                
                if chart_mina_comb:
                    snr_col_mina_comb = None
                    for col in app.df_data_mina_comb.columns:
                        if "SNR" in str(col).upper():
                            snr_col_mina_comb = col
                            break
                    if not snr_col_mina_comb and len(app.df_data_mina_comb.columns) >= 2:
                        snr_col_mina_comb = app.df_data_mina_comb.columns[1]

                    if snr_col_mina_comb:
                        categories_mco = []
                        proceso_vals_mco = []
                        diario_vals_mco = []
                        programa_vals_mco = []
                        columna1_vals_mco = []

                        prod_rows_mco = []
                        for idx, row in app.df_prod_mina_comb.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_mco.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_mco.append((cat, val))
                                    except: pass

                        if len(prod_rows_mco) > 30: prod_rows_mco = prod_rows_mco[-30:]

                        for i in range(len(prod_rows_mco)):
                            categories_mco.append(prod_rows_mco[i][0])
                            try: proceso_vals_mco.append(float(prod_rows_mco[i][1]))
                            except: proceso_vals_mco.append(None)
                            diario_vals_mco.append(None)
                            programa_vals_mco.append(None)
                            columna1_vals_mco.append(None)

                        for i in range(dias_mes):
                            categories_mco.append(str(i + 1))
                            proceso_vals_mco.append(None)
                            
                            try: diario_vals_mco.append(float(app.df_data_mina_comb[snr_col_mina_comb].iloc[i]))
                            except: diario_vals_mco.append(None)
                            
                            try: programa_vals_mco.append(float(app.df_snr_mina_comb.iloc[i, 0]))
                            except: programa_vals_mco.append(None)
                            
                            try: columna1_vals_mco.append(float(app.df_snr_mina_comb.iloc[i, 1]))
                            except: columna1_vals_mco.append(None)

                        update_slide_chart(chart_mina_comb, categories_mco, proceso_vals_mco, diario_vals_mco, programa_vals_mco, columna1_vals_mco, wine_color, green_color)

            # --- 18.1. PROCESAR DIAPOSITIVA DE TURBOSINA MINATITLAN (DIAPOSITIVA 27) ---
            if slide_exists(slides_map, "minatitlan", "turbosina") and getattr(app, 'df_data_mina_turb', None) is not None and getattr(app, 'df_snr_mina_turb', None) is not None and getattr(app, 'df_prod_mina_turb', None) is not None:
                slide_mina_turb = find_slide(slides_map, "minatitlan", "turbosina")
                chart_mina_turb = None
                for shape in slide_mina_turb.shapes:
                    if shape.has_chart:
                        chart_mina_turb = shape.chart
                        break
                
                if chart_mina_turb:
                    snr_col_mina_turb = None
                    for col in app.df_data_mina_turb.columns:
                        if "SNR" in str(col).upper():
                            snr_col_mina_turb = col
                            break
                    if not snr_col_mina_turb and len(app.df_data_mina_turb.columns) >= 2:
                        snr_col_mina_turb = app.df_data_mina_turb.columns[1]

                    if snr_col_mina_turb:
                        categories_mtu = []
                        proceso_vals_mtu = []
                        diario_vals_mtu = []
                        programa_vals_mtu = []
                        columna1_vals_mtu = []

                        prod_rows_mtu = []
                        for idx, row in app.df_prod_mina_turb.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_mtu.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_mtu.append((cat, val))
                                    except: pass

                        if len(prod_rows_mtu) > 30: prod_rows_mtu = prod_rows_mtu[-30:]

                        for i in range(len(prod_rows_mtu)):
                            categories_mtu.append(prod_rows_mtu[i][0])
                            try: proceso_vals_mtu.append(float(prod_rows_mtu[i][1]))
                            except: proceso_vals_mtu.append(None)
                            diario_vals_mtu.append(None)
                            programa_vals_mtu.append(None)
                            columna1_vals_mtu.append(None)

                        for i in range(dias_mes):
                            categories_mtu.append(str(i + 1))
                            proceso_vals_mtu.append(None)
                            
                            try: diario_vals_mtu.append(float(app.df_data_mina_turb[snr_col_mina_turb].iloc[i]))
                            except: diario_vals_mtu.append(None)
                            
                            try: programa_vals_mtu.append(float(app.df_snr_mina_turb.iloc[i, 0]))
                            except: programa_vals_mtu.append(None)
                            
                            try: columna1_vals_mtu.append(float(app.df_snr_mina_turb.iloc[i, 1]))
                            except: columna1_vals_mtu.append(None)

                        update_slide_chart(chart_mina_turb, categories_mtu, proceso_vals_mtu, diario_vals_mtu, programa_vals_mtu, columna1_vals_mtu, wine_color, green_color)

            # --- 19. PROCESAR DIAPOSITIVA DE CRUDO SALAMANCA (DIAPOSITIVA 30) ---
            if slide_exists(slides_map, "salamanca", "crudo") and app.df_data_sala_crud is not None and app.df_snr_sala_crud is not None and app.df_prod_sala_crud is not None:
                slide_sala_crud = find_slide(slides_map, "salamanca", "crudo")
                chart_sala_crud = None
                for shape in slide_sala_crud.shapes:
                    if shape.has_chart:
                        chart_sala_crud = shape.chart
                        break
                
                if chart_sala_crud:
                    snr_col_sala_crud = None
                    for col in app.df_data_sala_crud.columns:
                        if "SNR" in str(col).upper():
                            snr_col_sala_crud = col
                            break
                    if not snr_col_sala_crud and len(app.df_data_sala_crud.columns) >= 2:
                        snr_col_sala_crud = app.df_data_sala_crud.columns[1]

                    if snr_col_sala_crud:
                        categories_sc = []
                        proceso_vals_sc = []
                        diario_vals_sc = []
                        programa_vals_sc = []
                        columna1_vals_sc = []

                        prod_rows_sc = []
                        for idx, row in app.df_prod_sala_crud.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sc.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sc.append((cat, val))
                                    except: pass

                        if len(prod_rows_sc) > 30: prod_rows_sc = prod_rows_sc[-30:]

                        for i in range(len(prod_rows_sc)):
                            categories_sc.append(prod_rows_sc[i][0])
                            try: proceso_vals_sc.append(float(prod_rows_sc[i][1]))
                            except: proceso_vals_sc.append(None)
                            diario_vals_sc.append(None)
                            programa_vals_sc.append(None)
                            columna1_vals_sc.append(None)

                        for i in range(dias_mes):
                            categories_sc.append(str(i + 1))
                            proceso_vals_sc.append(None)
                            
                            try: diario_vals_sc.append(float(app.df_data_sala_crud[snr_col_sala_crud].iloc[i]))
                            except: diario_vals_sc.append(None)
                            
                            try: programa_vals_sc.append(float(app.df_snr_sala_crud.iloc[i, 0]))
                            except: programa_vals_sc.append(None)
                            
                            try: columna1_vals_sc.append(float(app.df_snr_sala_crud.iloc[i, 1]))
                            except: columna1_vals_sc.append(None)

                        update_slide_chart(chart_sala_crud, categories_sc, proceso_vals_sc, diario_vals_sc, programa_vals_sc, columna1_vals_sc, wine_color, green_color)

            # --- 20. PROCESAR DIAPOSITIVA DE GASOLINAS SALAMANCA (DIAPOSITIVA 31) ---
            if slide_exists(slides_map, "salamanca", "gasolinas") and app.df_data_sala_gas is not None and app.df_snr_sala_gas is not None and app.df_prod_sala_gas is not None:
                slide_sala_gas = find_slide(slides_map, "salamanca", "gasolinas")
                chart_sala_gas = None
                for shape in slide_sala_gas.shapes:
                    if shape.has_chart:
                        chart_sala_gas = shape.chart
                        break
                
                if chart_sala_gas:
                    snr_col_sala_gas = None
                    for col in app.df_data_sala_gas.columns:
                        if "SNR" in str(col).upper():
                            snr_col_sala_gas = col
                            break
                    if not snr_col_sala_gas and len(app.df_data_sala_gas.columns) >= 2:
                        snr_col_sala_gas = app.df_data_sala_gas.columns[1]

                    if snr_col_sala_gas:
                        categories_sg = []
                        proceso_vals_sg = []
                        diario_vals_sg = []
                        programa_vals_sg = []
                        columna1_vals_sg = []

                        prod_rows_sg = []
                        for idx, row in app.df_prod_sala_gas.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sg.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sg.append((cat, val))
                                    except: pass

                        if len(prod_rows_sg) > 30: prod_rows_sg = prod_rows_sg[-30:]

                        for i in range(len(prod_rows_sg)):
                            categories_sg.append(prod_rows_sg[i][0])
                            try: proceso_vals_sg.append(float(prod_rows_sg[i][1]))
                            except: proceso_vals_sg.append(None)
                            diario_vals_sg.append(None)
                            programa_vals_sg.append(None)
                            columna1_vals_sg.append(None)

                        for i in range(dias_mes):
                            categories_sg.append(str(i + 1))
                            proceso_vals_sg.append(None)
                            
                            try: diario_vals_sg.append(float(app.df_data_sala_gas[snr_col_sala_gas].iloc[i]))
                            except: diario_vals_sg.append(None)
                            
                            try: programa_vals_sg.append(float(app.df_snr_sala_gas.iloc[i, 0]))
                            except: programa_vals_sg.append(None)
                            
                            try: columna1_vals_sg.append(float(app.df_snr_sala_gas.iloc[i, 1]))
                            except: columna1_vals_sg.append(None)

                        update_slide_chart(chart_sala_gas, categories_sg, proceso_vals_sg, diario_vals_sg, programa_vals_sg, columna1_vals_sg, wine_color, green_color)

            # --- 21. PROCESAR DIAPOSITIVA DE DIESEL SALAMANCA (DIAPOSITIVA 32) ---
            if slide_exists(slides_map, "salamanca", "diesel") and app.df_data_sala_die is not None and app.df_snr_sala_die is not None and app.df_prod_sala_die is not None:
                slide_sala_die = find_slide(slides_map, "salamanca", "diesel")
                chart_sala_die = None
                for shape in slide_sala_die.shapes:
                    if shape.has_chart:
                        chart_sala_die = shape.chart
                        break
                
                if chart_sala_die:
                    snr_col_sala_die = None
                    for col in app.df_data_sala_die.columns:
                        if "SNR" in str(col).upper():
                            snr_col_sala_die = col
                            break
                    if not snr_col_sala_die and len(app.df_data_sala_die.columns) >= 2:
                        snr_col_sala_die = app.df_data_sala_die.columns[1]

                    if snr_col_sala_die:
                        categories_sd = []
                        proceso_vals_sd = []
                        diario_vals_sd = []
                        programa_vals_sd = []
                        columna1_vals_sd = []

                        prod_rows_sd = []
                        for idx, row in app.df_prod_sala_die.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sd.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sd.append((cat, val))
                                    except: pass

                        if len(prod_rows_sd) > 30: prod_rows_sd = prod_rows_sd[-30:]

                        for i in range(len(prod_rows_sd)):
                            categories_sd.append(prod_rows_sd[i][0])
                            try: proceso_vals_sd.append(float(prod_rows_sd[i][1]))
                            except: proceso_vals_sd.append(None)
                            diario_vals_sd.append(None)
                            programa_vals_sd.append(None)
                            columna1_vals_sd.append(None)

                        for i in range(dias_mes):
                            categories_sd.append(str(i + 1))
                            proceso_vals_sd.append(None)
                            
                            try: diario_vals_sd.append(float(app.df_data_sala_die[snr_col_sala_die].iloc[i]))
                            except: diario_vals_sd.append(None)
                            
                            try: programa_vals_sd.append(float(app.df_snr_sala_die.iloc[i, 0]))
                            except: programa_vals_sd.append(None)
                            
                            try: columna1_vals_sd.append(float(app.df_snr_sala_die.iloc[i, 1]))
                            except: columna1_vals_sd.append(None)

                        update_slide_chart(chart_sala_die, categories_sd, proceso_vals_sd, diario_vals_sd, programa_vals_sd, columna1_vals_sd, wine_color, green_color)

            # --- 22. PROCESAR DIAPOSITIVA DE TURBOSINA SALAMANCA (DIAPOSITIVA 33) ---
            if slide_exists(slides_map, "salamanca", "turbosina") and app.df_data_sala_turb is not None and app.df_snr_sala_turb is not None and app.df_prod_sala_turb is not None:
                slide_sala_turb = find_slide(slides_map, "salamanca", "turbosina")
                chart_sala_turb = None
                for shape in slide_sala_turb.shapes:
                    if shape.has_chart:
                        chart_sala_turb = shape.chart
                        break
                
                if chart_sala_turb:
                    snr_col_sala_turb = None
                    for col in app.df_data_sala_turb.columns:
                        if "SNR" in str(col).upper():
                            snr_col_sala_turb = col
                            break
                    if not snr_col_sala_turb and len(app.df_data_sala_turb.columns) >= 2:
                        snr_col_sala_turb = app.df_data_sala_turb.columns[1]

                    if snr_col_sala_turb:
                        categories_stur = []
                        proceso_vals_stur = []
                        diario_vals_stur = []
                        programa_vals_stur = []
                        columna1_vals_stur = []

                        prod_rows_stur = []
                        for idx, row in app.df_prod_sala_turb.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_stur.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_stur.append((cat, val))
                                    except: pass

                        if len(prod_rows_stur) > 30: prod_rows_stur = prod_rows_stur[-30:]

                        for i in range(len(prod_rows_stur)):
                            categories_stur.append(prod_rows_stur[i][0])
                            try: proceso_vals_stur.append(float(prod_rows_stur[i][1]))
                            except: proceso_vals_stur.append(None)
                            diario_vals_stur.append(None)
                            programa_vals_stur.append(None)
                            columna1_vals_stur.append(None)

                        for i in range(dias_mes):
                            categories_stur.append(str(i + 1))
                            proceso_vals_stur.append(None)
                            
                            try: diario_vals_stur.append(float(app.df_data_sala_turb[snr_col_sala_turb].iloc[i]))
                            except: diario_vals_stur.append(None)
                            
                            try: programa_vals_stur.append(float(app.df_snr_sala_turb.iloc[i, 0]))
                            except: programa_vals_stur.append(None)
                            
                            try: columna1_vals_stur.append(float(app.df_snr_sala_turb.iloc[i, 1]))
                            except: columna1_vals_stur.append(None)

                        update_slide_chart(chart_sala_turb, categories_stur, proceso_vals_stur, diario_vals_stur, programa_vals_stur, columna1_vals_stur, wine_color, green_color)

            # --- 23. PROCESAR DIAPOSITIVA DE COMBUSTOLEO SALAMANCA (DIAPOSITIVA 34) ---
            if slide_exists(slides_map, "salamanca", "combustoleo") and app.df_data_sala_comb is not None and app.df_snr_sala_comb is not None and app.df_prod_sala_comb is not None:
                slide_sala_comb = find_slide(slides_map, "salamanca", "combustoleo")
                chart_sala_comb = None
                for shape in slide_sala_comb.shapes:
                    if shape.has_chart:
                        chart_sala_comb = shape.chart
                        break
                
                if chart_sala_comb:
                    snr_col_sala_comb = None
                    for col in app.df_data_sala_comb.columns:
                        if "SNR" in str(col).upper():
                            snr_col_sala_comb = col
                            break
                    if not snr_col_sala_comb and len(app.df_data_sala_comb.columns) >= 2:
                        snr_col_sala_comb = app.df_data_sala_comb.columns[1]

                    if snr_col_sala_comb:
                        categories_sco = []
                        proceso_vals_sco = []
                        diario_vals_sco = []
                        programa_vals_sco = []
                        columna1_vals_sco = []

                        prod_rows_sco = []
                        for idx, row in app.df_prod_sala_comb.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sco.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sco.append((cat, val))
                                    except: pass

                        if len(prod_rows_sco) > 30: prod_rows_sco = prod_rows_sco[-30:]

                        for i in range(len(prod_rows_sco)):
                            categories_sco.append(prod_rows_sco[i][0])
                            try: proceso_vals_sco.append(float(prod_rows_sco[i][1]))
                            except: proceso_vals_sco.append(None)
                            diario_vals_sco.append(None)
                            programa_vals_sco.append(None)
                            columna1_vals_sco.append(None)

                        for i in range(dias_mes):
                            categories_sco.append(str(i + 1))
                            proceso_vals_sco.append(None)
                            
                            try: diario_vals_sco.append(float(app.df_data_sala_comb[snr_col_sala_comb].iloc[i]))
                            except: diario_vals_sco.append(None)
                            
                            try: programa_vals_sco.append(float(app.df_snr_sala_comb.iloc[i, 0]))
                            except: programa_vals_sco.append(None)
                            
                            try: columna1_vals_sco.append(float(app.df_snr_sala_comb.iloc[i, 1]))
                            except: columna1_vals_sco.append(None)

                        update_slide_chart(chart_sala_comb, categories_sco, proceso_vals_sco, diario_vals_sco, programa_vals_sco, columna1_vals_sco, wine_color, green_color)

            # --- 24. PROCESAR DIAPOSITIVA DE CRUDO SALINA CRUZ (DIAPOSITIVA 37) ---
            if slide_exists(slides_map, "salina cruz", "crudo") and app.df_data_sal_crud is not None and app.df_snr_sal_crud is not None and app.df_prod_sal_crud is not None:
                slide_sal_crud = find_slide(slides_map, "salina cruz", "crudo")
                chart_sal_crud = None
                for shape in slide_sal_crud.shapes:
                    if shape.has_chart:
                        chart_sal_crud = shape.chart
                        break
                
                if chart_sal_crud:
                    snr_col_sal_crud = None
                    for col in app.df_data_sal_crud.columns:
                        if "SNR" in str(col).upper():
                            snr_col_sal_crud = col
                            break
                    if not snr_col_sal_crud and len(app.df_data_sal_crud.columns) >= 2:
                        snr_col_sal_crud = app.df_data_sal_crud.columns[1]

                    if snr_col_sal_crud:
                        categories_sc = []
                        proceso_vals_sc = []
                        diario_vals_sc = []
                        programa_vals_sc = []
                        columna1_vals_sc = []

                        prod_rows_sc = []
                        for idx, row in app.df_prod_sal_crud.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sc.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sc.append((cat, val))
                                    except: pass

                        if len(prod_rows_sc) > 30: prod_rows_sc = prod_rows_sc[-30:]

                        for i in range(len(prod_rows_sc)):
                            categories_sc.append(prod_rows_sc[i][0])
                            try: proceso_vals_sc.append(float(prod_rows_sc[i][1]))
                            except: proceso_vals_sc.append(None)
                            diario_vals_sc.append(None)
                            programa_vals_sc.append(None)
                            columna1_vals_sc.append(None)

                        for i in range(dias_mes):
                            categories_sc.append(str(i + 1))
                            proceso_vals_sc.append(None)
                            
                            try: diario_vals_sc.append(float(app.df_data_sal_crud[snr_col_sal_crud].iloc[i]))
                            except: diario_vals_sc.append(None)
                            
                            try: programa_vals_sc.append(float(app.df_snr_sal_crud.iloc[i, 0]))
                            except: programa_vals_sc.append(None)
                            
                            try: columna1_vals_sc.append(float(app.df_snr_sal_crud.iloc[i, 1]))
                            except: columna1_vals_sc.append(None)

                        update_slide_chart(chart_sal_crud, categories_sc, proceso_vals_sc, diario_vals_sc, programa_vals_sc, columna1_vals_sc, wine_color, green_color)

            # --- 25. PROCESAR DIAPOSITIVA DE GASOLINAS SALINA CRUZ (DIAPOSITIVA 38) ---
            if slide_exists(slides_map, "salina cruz", "gasolinas") and app.df_data_sal_gas is not None and app.df_snr_sal_gas is not None and app.df_prod_sal_gas is not None:
                slide_sal_gas = find_slide(slides_map, "salina cruz", "gasolinas")
                chart_sal_gas = None
                for shape in slide_sal_gas.shapes:
                    if shape.has_chart:
                        chart_sal_gas = shape.chart
                        break
                
                if chart_sal_gas:
                    snr_col_sal_gas = None
                    for col in app.df_data_sal_gas.columns:
                        if "SNR" in str(col).upper():
                            snr_col_sal_gas = col
                            break
                    if not snr_col_sal_gas and len(app.df_data_sal_gas.columns) >= 2:
                        snr_col_sal_gas = app.df_data_sal_gas.columns[1]

                    if snr_col_sal_gas:
                        categories_sg = []
                        proceso_vals_sg = []
                        diario_vals_sg = []
                        programa_vals_sg = []
                        columna1_vals_sg = []

                        prod_rows_sg = []
                        for idx, row in app.df_prod_sal_gas.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sg.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sg.append((cat, val))
                                    except: pass

                        if len(prod_rows_sg) > 30: prod_rows_sg = prod_rows_sg[-30:]

                        for i in range(len(prod_rows_sg)):
                            categories_sg.append(prod_rows_sg[i][0])
                            try: proceso_vals_sg.append(float(prod_rows_sg[i][1]))
                            except: proceso_vals_sg.append(None)
                            diario_vals_sg.append(None)
                            programa_vals_sg.append(None)
                            columna1_vals_sg.append(None)

                        for i in range(dias_mes):
                            categories_sg.append(str(i + 1))
                            proceso_vals_sg.append(None)
                            
                            try: diario_vals_sg.append(float(app.df_data_sal_gas[snr_col_sal_gas].iloc[i]))
                            except: diario_vals_sg.append(None)
                            
                            try: programa_vals_sg.append(float(app.df_snr_sal_gas.iloc[i, 0]))
                            except: programa_vals_sg.append(None)
                            
                            try: columna1_vals_sg.append(float(app.df_snr_sal_gas.iloc[i, 1]))
                            except: columna1_vals_sg.append(None)

                        update_slide_chart(chart_sal_gas, categories_sg, proceso_vals_sg, diario_vals_sg, programa_vals_sg, columna1_vals_sg, wine_color, green_color)

            # --- 26. PROCESAR DIAPOSITIVA DE DIESEL SALINA CRUZ (DIAPOSITIVA 39) ---
            if slide_exists(slides_map, "salina cruz", "diesel") and app.df_data_sal_die is not None and app.df_snr_sal_die is not None and app.df_prod_sal_die is not None:
                slide_sal_die = find_slide(slides_map, "salina cruz", "diesel")
                chart_sal_die = None
                for shape in slide_sal_die.shapes:
                    if shape.has_chart:
                        chart_sal_die = shape.chart
                        break
                
                if chart_sal_die:
                    snr_col_sal_die = None
                    for col in app.df_data_sal_die.columns:
                        if "SNR" in str(col).upper():
                            snr_col_sal_die = col
                            break
                    if not snr_col_sal_die and len(app.df_data_sal_die.columns) >= 2:
                        snr_col_sal_die = app.df_data_sal_die.columns[1]

                    if snr_col_sal_die:
                        categories_sd = []
                        proceso_vals_sd = []
                        diario_vals_sd = []
                        programa_vals_sd = []
                        columna1_vals_sd = []

                        prod_rows_sd = []
                        for idx, row in app.df_prod_sal_die.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sd.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sd.append((cat, val))
                                    except: pass

                        if len(prod_rows_sd) > 30: prod_rows_sd = prod_rows_sd[-30:]

                        for i in range(len(prod_rows_sd)):
                            categories_sd.append(prod_rows_sd[i][0])
                            try: proceso_vals_sd.append(float(prod_rows_sd[i][1]))
                            except: proceso_vals_sd.append(None)
                            diario_vals_sd.append(None)
                            programa_vals_sd.append(None)
                            columna1_vals_sd.append(None)

                        for i in range(dias_mes):
                            categories_sd.append(str(i + 1))
                            proceso_vals_sd.append(None)
                            
                            try: diario_vals_sd.append(float(app.df_data_sal_die[snr_col_sal_die].iloc[i]))
                            except: diario_vals_sd.append(None)
                            
                            try: programa_vals_sd.append(float(app.df_snr_sal_die.iloc[i, 0]))
                            except: programa_vals_sd.append(None)
                            
                            try: columna1_vals_sd.append(float(app.df_snr_sal_die.iloc[i, 1]))
                            except: columna1_vals_sd.append(None)

                        update_slide_chart(chart_sal_die, categories_sd, proceso_vals_sd, diario_vals_sd, programa_vals_sd, columna1_vals_sd, wine_color, green_color)

            # --- 27. PROCESAR DIAPOSITIVA DE TURBOSINA SALINA CRUZ (DIAPOSITIVA 40) ---
            if slide_exists(slides_map, "salina cruz", "turbosina") and app.df_data_sal_turb is not None and app.df_snr_sal_turb is not None and app.df_prod_sal_turb is not None:
                slide_sal_turb = find_slide(slides_map, "salina cruz", "turbosina")
                chart_sal_turb = None
                for shape in slide_sal_turb.shapes:
                    if shape.has_chart:
                        chart_sal_turb = shape.chart
                        break
                
                if chart_sal_turb:
                    snr_col_sal_turb = None
                    for col in app.df_data_sal_turb.columns:
                        if "SNR" in str(col).upper():
                            snr_col_sal_turb = col
                            break
                    if not snr_col_sal_turb and len(app.df_data_sal_turb.columns) >= 2:
                        snr_col_sal_turb = app.df_data_sal_turb.columns[1]

                    if snr_col_sal_turb:
                        categories_stur = []
                        proceso_vals_stur = []
                        diario_vals_stur = []
                        programa_vals_stur = []
                        columna1_vals_stur = []

                        prod_rows_stur = []
                        for idx, row in app.df_prod_sal_turb.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_stur.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_stur.append((cat, val))
                                    except: pass

                        if len(prod_rows_stur) > 30: prod_rows_stur = prod_rows_stur[-30:]

                        for i in range(len(prod_rows_stur)):
                            categories_stur.append(prod_rows_stur[i][0])
                            try: proceso_vals_stur.append(float(prod_rows_stur[i][1]))
                            except: proceso_vals_stur.append(None)
                            diario_vals_stur.append(None)
                            programa_vals_stur.append(None)
                            columna1_vals_stur.append(None)

                        for i in range(dias_mes):
                            categories_stur.append(str(i + 1))
                            proceso_vals_stur.append(None)
                            
                            try: diario_vals_stur.append(float(app.df_data_sal_turb[snr_col_sal_turb].iloc[i]))
                            except: diario_vals_stur.append(None)
                            
                            try: programa_vals_stur.append(float(app.df_snr_sal_turb.iloc[i, 0]))
                            except: programa_vals_stur.append(None)
                            
                            try: columna1_vals_stur.append(float(app.df_snr_sal_turb.iloc[i, 1]))
                            except: columna1_vals_stur.append(None)

                        update_slide_chart(chart_sal_turb, categories_stur, proceso_vals_stur, diario_vals_stur, programa_vals_stur, columna1_vals_stur, wine_color, green_color)

            # --- 28. PROCESAR DIAPOSITIVA DE COMBUSTOLEO SALINA CRUZ (DIAPOSITIVA 41) ---
            if slide_exists(slides_map, "salina cruz", "combustoleo") and app.df_data_sal_comb is not None and app.df_snr_sal_comb is not None and app.df_prod_sal_comb is not None:
                slide_sal_comb = find_slide(slides_map, "salina cruz", "combustoleo")
                chart_sal_comb = None
                for shape in slide_sal_comb.shapes:
                    if shape.has_chart:
                        chart_sal_comb = shape.chart
                        break
                
                if chart_sal_comb:
                    snr_col_sal_comb = None
                    for col in app.df_data_sal_comb.columns:
                        if "SNR" in str(col).upper():
                            snr_col_sal_comb = col
                            break
                    if not snr_col_sal_comb and len(app.df_data_sal_comb.columns) >= 2:
                        snr_col_sal_comb = app.df_data_sal_comb.columns[1]

                    if snr_col_sal_comb:
                        categories_sco = []
                        proceso_vals_sco = []
                        diario_vals_sco = []
                        programa_vals_sco = []
                        columna1_vals_sco = []

                        prod_rows_sco = []
                        for idx, row in app.df_prod_sal_comb.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sco.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sco.append((cat, val))
                                    except: pass

                        if len(prod_rows_sco) > 30: prod_rows_sco = prod_rows_sco[-30:]

                        for i in range(len(prod_rows_sco)):
                            categories_sco.append(prod_rows_sco[i][0])
                            try: proceso_vals_sco.append(float(prod_rows_sco[i][1]))
                            except: proceso_vals_sco.append(None)
                            diario_vals_sco.append(None)
                            programa_vals_sco.append(None)
                            columna1_vals_sco.append(None)

                        for i in range(dias_mes):
                            categories_sco.append(str(i + 1))
                            proceso_vals_sco.append(None)
                            
                            try: diario_vals_sco.append(float(app.df_data_sal_comb[snr_col_sal_comb].iloc[i]))
                            except: diario_vals_sco.append(None)
                            
                            try: programa_vals_sco.append(float(app.df_snr_sal_comb.iloc[i, 0]))
                            except: programa_vals_sco.append(None)
                            
                            try: columna1_vals_sco.append(float(app.df_snr_sal_comb.iloc[i, 1]))
                            except: columna1_vals_sco.append(None)

                        update_slide_chart(chart_sal_comb, categories_sco, proceso_vals_sco, diario_vals_sco, programa_vals_sco, columna1_vals_sco, wine_color, green_color)

            # --- 29. PROCESAR DIAPOSITIVA DE CRUDO TULA (DIAPOSITIVA 44) ---
            if slide_exists(slides_map, "tula", "crudo") and app.df_data_tula_crud is not None and app.df_snr_tula_crud is not None and app.df_prod_tula_crud is not None:
                slide_tula_crud = find_slide(slides_map, "tula", "crudo")
                chart_tula_crud = None
                for shape in slide_tula_crud.shapes:
                    if shape.has_chart:
                        chart_tula_crud = shape.chart
                        break
                
                if chart_tula_crud:
                    snr_col_tula_crud = None
                    for col in app.df_data_tula_crud.columns:
                        if "SNR" in str(col).upper():
                            snr_col_tula_crud = col
                            break
                    if not snr_col_tula_crud and len(app.df_data_tula_crud.columns) >= 2:
                        snr_col_tula_crud = app.df_data_tula_crud.columns[1]

                    if snr_col_tula_crud:
                        categories_sc = []
                        proceso_vals_sc = []
                        diario_vals_sc = []
                        programa_vals_sc = []
                        columna1_vals_sc = []

                        prod_rows_sc = []
                        for idx, row in app.df_prod_tula_crud.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sc.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sc.append((cat, val))
                                    except: pass

                        if len(prod_rows_sc) > 30: prod_rows_sc = prod_rows_sc[-30:]

                        for i in range(len(prod_rows_sc)):
                            categories_sc.append(prod_rows_sc[i][0])
                            try: proceso_vals_sc.append(float(prod_rows_sc[i][1]))
                            except: proceso_vals_sc.append(None)
                            diario_vals_sc.append(None)
                            programa_vals_sc.append(None)
                            columna1_vals_sc.append(None)

                        for i in range(dias_mes):
                            categories_sc.append(str(i + 1))
                            proceso_vals_sc.append(None)
                            
                            try: diario_vals_sc.append(float(app.df_data_tula_crud[snr_col_tula_crud].iloc[i]))
                            except: diario_vals_sc.append(None)
                            
                            try: programa_vals_sc.append(float(app.df_snr_tula_crud.iloc[i, 0]))
                            except: programa_vals_sc.append(None)
                            
                            try: columna1_vals_sc.append(float(app.df_snr_tula_crud.iloc[i, 1]))
                            except: columna1_vals_sc.append(None)

                        update_slide_chart(chart_tula_crud, categories_sc, proceso_vals_sc, diario_vals_sc, programa_vals_sc, columna1_vals_sc, wine_color, green_color)

            # --- 30. PROCESAR DIAPOSITIVA DE GASOLINAS TULA (DIAPOSITIVA 45) ---
            if slide_exists(slides_map, "tula", "gasolinas") and app.df_data_tula_gas is not None and app.df_snr_tula_gas is not None and app.df_prod_tula_gas is not None:
                slide_tula_gas = find_slide(slides_map, "tula", "gasolinas")
                chart_tula_gas = None
                for shape in slide_tula_gas.shapes:
                    if shape.has_chart:
                        chart_tula_gas = shape.chart
                        break
                
                if chart_tula_gas:
                    snr_col_tula_gas = None
                    for col in app.df_data_tula_gas.columns:
                        if "SNR" in str(col).upper():
                            snr_col_tula_gas = col
                            break
                    if not snr_col_tula_gas and len(app.df_data_tula_gas.columns) >= 2:
                        snr_col_tula_gas = app.df_data_tula_gas.columns[1]

                    if snr_col_tula_gas:
                        categories_sg = []
                        proceso_vals_sg = []
                        diario_vals_sg = []
                        programa_vals_sg = []
                        columna1_vals_sg = []

                        prod_rows_sg = []
                        for idx, row in app.df_prod_tula_gas.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sg.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sg.append((cat, val))
                                    except: pass

                        if len(prod_rows_sg) > 30: prod_rows_sg = prod_rows_sg[-30:]

                        for i in range(len(prod_rows_sg)):
                            categories_sg.append(prod_rows_sg[i][0])
                            try: proceso_vals_sg.append(float(prod_rows_sg[i][1]))
                            except: proceso_vals_sg.append(None)
                            diario_vals_sg.append(None)
                            programa_vals_sg.append(None)
                            columna1_vals_sg.append(None)

                        for i in range(dias_mes):
                            categories_sg.append(str(i + 1))
                            proceso_vals_sg.append(None)
                            
                            try: diario_vals_sg.append(float(app.df_data_tula_gas[snr_col_tula_gas].iloc[i]))
                            except: diario_vals_sg.append(None)
                            
                            try: programa_vals_sg.append(float(app.df_snr_tula_gas.iloc[i, 0]))
                            except: programa_vals_sg.append(None)
                            
                            try: columna1_vals_sg.append(float(app.df_snr_tula_gas.iloc[i, 1]))
                            except: columna1_vals_sg.append(None)

                        update_slide_chart(chart_tula_gas, categories_sg, proceso_vals_sg, diario_vals_sg, programa_vals_sg, columna1_vals_sg, wine_color, green_color)

            # --- 31. PROCESAR DIAPOSITIVA DE DIESEL TULA (DIAPOSITIVA 46) ---
            if slide_exists(slides_map, "tula", "diesel") and app.df_data_tula_die is not None and app.df_snr_tula_die is not None and app.df_prod_tula_die is not None:
                slide_tula_die = find_slide(slides_map, "tula", "diesel")
                chart_tula_die = None
                for shape in slide_tula_die.shapes:
                    if shape.has_chart:
                        chart_tula_die = shape.chart
                        break
                
                if chart_tula_die:
                    snr_col_tula_die = None
                    for col in app.df_data_tula_die.columns:
                        if "SNR" in str(col).upper():
                            snr_col_tula_die = col
                            break
                    if not snr_col_tula_die and len(app.df_data_tula_die.columns) >= 2:
                        snr_col_tula_die = app.df_data_tula_die.columns[1]

                    if snr_col_tula_die:
                        categories_sd = []
                        proceso_vals_sd = []
                        diario_vals_sd = []
                        programa_vals_sd = []
                        columna1_vals_sd = []

                        prod_rows_sd = []
                        for idx, row in app.df_prod_tula_die.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sd.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sd.append((cat, val))
                                    except: pass

                        if len(prod_rows_sd) > 30: prod_rows_sd = prod_rows_sd[-30:]

                        for i in range(len(prod_rows_sd)):
                            categories_sd.append(prod_rows_sd[i][0])
                            try: proceso_vals_sd.append(float(prod_rows_sd[i][1]))
                            except: proceso_vals_sd.append(None)
                            diario_vals_sd.append(None)
                            programa_vals_sd.append(None)
                            columna1_vals_sd.append(None)

                        for i in range(dias_mes):
                            categories_sd.append(str(i + 1))
                            proceso_vals_sd.append(None)
                            
                            try: diario_vals_sd.append(float(app.df_data_tula_die[snr_col_tula_die].iloc[i]))
                            except: diario_vals_sd.append(None)
                            
                            try: programa_vals_sd.append(float(app.df_snr_tula_die.iloc[i, 0]))
                            except: programa_vals_sd.append(None)
                            
                            try: columna1_vals_sd.append(float(app.df_snr_tula_die.iloc[i, 1]))
                            except: columna1_vals_sd.append(None)

                        update_slide_chart(chart_tula_die, categories_sd, proceso_vals_sd, diario_vals_sd, programa_vals_sd, columna1_vals_sd, wine_color, green_color)

            # --- 32. PROCESAR DIAPOSITIVA DE TURBOSINA TULA (DIAPOSITIVA 47) ---
            if slide_exists(slides_map, "tula", "turbosina") and app.df_data_tula_turb is not None and app.df_snr_tula_turb is not None and app.df_prod_tula_turb is not None:
                slide_tula_turb = find_slide(slides_map, "tula", "turbosina")
                chart_tula_turb = None
                for shape in slide_tula_turb.shapes:
                    if shape.has_chart:
                        chart_tula_turb = shape.chart
                        break
                
                if chart_tula_turb:
                    snr_col_tula_turb = None
                    for col in app.df_data_tula_turb.columns:
                        if "SNR" in str(col).upper():
                            snr_col_tula_turb = col
                            break
                    if not snr_col_tula_turb and len(app.df_data_tula_turb.columns) >= 2:
                        snr_col_tula_turb = app.df_data_tula_turb.columns[1]

                    if snr_col_tula_turb:
                        categories_stur = []
                        proceso_vals_stur = []
                        diario_vals_stur = []
                        programa_vals_stur = []
                        columna1_vals_stur = []

                        prod_rows_stur = []
                        for idx, row in app.df_prod_tula_turb.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_stur.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_stur.append((cat, val))
                                    except: pass

                        if len(prod_rows_stur) > 30: prod_rows_stur = prod_rows_stur[-30:]

                        for i in range(len(prod_rows_stur)):
                            categories_stur.append(prod_rows_stur[i][0])
                            try: proceso_vals_stur.append(float(prod_rows_stur[i][1]))
                            except: proceso_vals_stur.append(None)
                            diario_vals_stur.append(None)
                            programa_vals_stur.append(None)
                            columna1_vals_stur.append(None)

                        for i in range(dias_mes):
                            categories_stur.append(str(i + 1))
                            proceso_vals_stur.append(None)
                            
                            try: diario_vals_stur.append(float(app.df_data_tula_turb[snr_col_tula_turb].iloc[i]))
                            except: diario_vals_stur.append(None)
                            
                            try: programa_vals_stur.append(float(app.df_snr_tula_turb.iloc[i, 0]))
                            except: programa_vals_stur.append(None)
                            
                            try: columna1_vals_stur.append(float(app.df_snr_tula_turb.iloc[i, 1]))
                            except: columna1_vals_stur.append(None)

                        update_slide_chart(chart_tula_turb, categories_stur, proceso_vals_stur, diario_vals_stur, programa_vals_stur, columna1_vals_stur, wine_color, green_color)

            # --- 33. PROCESAR DIAPOSITIVA DE COMBUSTOLEO TULA (DIAPOSITIVA 48) ---
            if slide_exists(slides_map, "tula", "combustoleo") and app.df_data_tula_comb is not None and app.df_snr_tula_comb is not None and app.df_prod_tula_comb is not None:
                slide_tula_comb = find_slide(slides_map, "tula", "combustoleo")
                chart_tula_comb = None
                for shape in slide_tula_comb.shapes:
                    if shape.has_chart:
                        chart_tula_comb = shape.chart
                        break
                
                if chart_tula_comb:
                    snr_col_tula_comb = None
                    for col in app.df_data_tula_comb.columns:
                        if "SNR" in str(col).upper():
                            snr_col_tula_comb = col
                            break
                    if not snr_col_tula_comb and len(app.df_data_tula_comb.columns) >= 2:
                        snr_col_tula_comb = app.df_data_tula_comb.columns[1]

                    if snr_col_tula_comb:
                        categories_sco = []
                        proceso_vals_sco = []
                        diario_vals_sco = []
                        programa_vals_sco = []
                        columna1_vals_sco = []

                        prod_rows_sco = []
                        for idx, row in app.df_prod_tula_comb.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sco.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sco.append((cat, val))
                                    except: pass

                        if len(prod_rows_sco) > 30: prod_rows_sco = prod_rows_sco[-30:]

                        for i in range(len(prod_rows_sco)):
                            categories_sco.append(prod_rows_sco[i][0])
                            try: proceso_vals_sco.append(float(prod_rows_sco[i][1]))
                            except: proceso_vals_sco.append(None)
                            diario_vals_sco.append(None)
                            programa_vals_sco.append(None)
                            columna1_vals_sco.append(None)

                        for i in range(dias_mes):
                            categories_sco.append(str(i + 1))
                            proceso_vals_sco.append(None)
                            
                            try: diario_vals_sco.append(float(app.df_data_tula_comb[snr_col_tula_comb].iloc[i]))
                            except: diario_vals_sco.append(None)
                            
                            try: programa_vals_sco.append(float(app.df_snr_tula_comb.iloc[i, 0]))
                            except: programa_vals_sco.append(None)
                            
                            try: columna1_vals_sco.append(float(app.df_snr_tula_comb.iloc[i, 1]))
                            except: columna1_vals_sco.append(None)

                        update_slide_chart(chart_tula_comb, categories_sco, proceso_vals_sco, diario_vals_sco, programa_vals_sco, columna1_vals_sco, wine_color, green_color)

            # --- 34. PROCESAR DIAPOSITIVA DE CRUDO OLMECA (DIAPOSITIVA 51) ---
            if slide_exists(slides_map, "olmeca", "crudo") and app.df_data_olme_crud is not None and app.df_snr_olme_crud is not None and app.df_prod_olme_crud is not None:
                slide_olme_crud = find_slide(slides_map, "olmeca", "crudo")
                chart_olme_crud = None
                for shape in slide_olme_crud.shapes:
                    if shape.has_chart:
                        chart_olme_crud = shape.chart
                        break
                
                if chart_olme_crud:
                    snr_col_olme_crud = None
                    for col in app.df_data_olme_crud.columns:
                        if "SNR" in str(col).upper():
                            snr_col_olme_crud = col
                            break
                    if not snr_col_olme_crud and len(app.df_data_olme_crud.columns) >= 2:
                        snr_col_olme_crud = app.df_data_olme_crud.columns[1]

                    if snr_col_olme_crud:
                        categories_sc = []
                        proceso_vals_sc = []
                        diario_vals_sc = []
                        programa_vals_sc = []
                        columna1_vals_sc = []

                        prod_rows_sc = []
                        for idx, row in app.df_prod_olme_crud.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sc.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sc.append((cat, val))
                                    except: pass

                        if len(prod_rows_sc) > 30: prod_rows_sc = prod_rows_sc[-30:]

                        for i in range(len(prod_rows_sc)):
                            categories_sc.append(prod_rows_sc[i][0])
                            try: proceso_vals_sc.append(float(prod_rows_sc[i][1]))
                            except: proceso_vals_sc.append(None)
                            diario_vals_sc.append(None)
                            programa_vals_sc.append(None)
                            columna1_vals_sc.append(None)

                        for i in range(dias_mes):
                            categories_sc.append(str(i + 1))
                            proceso_vals_sc.append(None)
                            
                            try: diario_vals_sc.append(float(app.df_data_olme_crud[snr_col_olme_crud].iloc[i]))
                            except: diario_vals_sc.append(None)
                            
                            try: programa_vals_sc.append(float(app.df_snr_olme_crud.iloc[i, 0]))
                            except: programa_vals_sc.append(None)
                            
                            try: columna1_vals_sc.append(float(app.df_snr_olme_crud.iloc[i, 1]))
                            except: columna1_vals_sc.append(None)

                        update_slide_chart(chart_olme_crud, categories_sc, proceso_vals_sc, diario_vals_sc, programa_vals_sc, columna1_vals_sc, wine_color, green_color)

            # --- 35. PROCESAR DIAPOSITIVA DE GASOLINAS OLMECA (DIAPOSITIVA 52) ---
            if slide_exists(slides_map, "olmeca", "gasolinas") and app.df_data_olme_gas is not None and app.df_snr_olme_gas is not None and app.df_prod_olme_gas is not None:
                slide_olme_gas = find_slide(slides_map, "olmeca", "gasolinas")
                chart_olme_gas = None
                for shape in slide_olme_gas.shapes:
                    if shape.has_chart:
                        chart_olme_gas = shape.chart
                        break
                
                if chart_olme_gas:
                    snr_col_olme_gas = None
                    for col in app.df_data_olme_gas.columns:
                        if "SNR" in str(col).upper():
                            snr_col_olme_gas = col
                            break
                    if not snr_col_olme_gas and len(app.df_data_olme_gas.columns) >= 2:
                        snr_col_olme_gas = app.df_data_olme_gas.columns[1]

                    if snr_col_olme_gas:
                        categories_sg = []
                        proceso_vals_sg = []
                        diario_vals_sg = []
                        programa_vals_sg = []
                        columna1_vals_sg = []

                        prod_rows_sg = []
                        for idx, row in app.df_prod_olme_gas.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sg.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sg.append((cat, val))
                                    except: pass

                        if len(prod_rows_sg) > 30: prod_rows_sg = prod_rows_sg[-30:]

                        for i in range(len(prod_rows_sg)):
                            categories_sg.append(prod_rows_sg[i][0])
                            try: proceso_vals_sg.append(float(prod_rows_sg[i][1]))
                            except: proceso_vals_sg.append(None)
                            diario_vals_sg.append(None)
                            programa_vals_sg.append(None)
                            columna1_vals_sg.append(None)

                        for i in range(dias_mes):
                            categories_sg.append(str(i + 1))
                            proceso_vals_sg.append(None)
                            
                            try: diario_vals_sg.append(float(app.df_data_olme_gas[snr_col_olme_gas].iloc[i]))
                            except: diario_vals_sg.append(None)
                            
                            try: programa_vals_sg.append(float(app.df_snr_olme_gas.iloc[i, 0]))
                            except: programa_vals_sg.append(None)
                            
                            try: columna1_vals_sg.append(float(app.df_snr_olme_gas.iloc[i, 1]))
                            except: columna1_vals_sg.append(None)

                        update_slide_chart(chart_olme_gas, categories_sg, proceso_vals_sg, diario_vals_sg, programa_vals_sg, columna1_vals_sg, wine_color, green_color)

            # --- 36. PROCESAR DIAPOSITIVA DE DIESEL OLMECA (DIAPOSITIVA 53) ---
            if slide_exists(slides_map, "olmeca", "diesel") and app.df_data_olme_die is not None and app.df_snr_olme_die is not None and app.df_prod_olme_die is not None:
                slide_olme_die = find_slide(slides_map, "olmeca", "diesel")
                chart_olme_die = None
                for shape in slide_olme_die.shapes:
                    if shape.has_chart:
                        chart_olme_die = shape.chart
                        break
                
                if chart_olme_die:
                    snr_col_olme_die = None
                    for col in app.df_data_olme_die.columns:
                        if "SNR" in str(col).upper():
                            snr_col_olme_die = col
                            break
                    if not snr_col_olme_die and len(app.df_data_olme_die.columns) >= 2:
                        snr_col_olme_die = app.df_data_olme_die.columns[1]

                    if snr_col_olme_die:
                        categories_sd = []
                        proceso_vals_sd = []
                        diario_vals_sd = []
                        programa_vals_sd = []
                        columna1_vals_sd = []

                        prod_rows_sd = []
                        for idx, row in app.df_prod_olme_die.iterrows():
                            cat = str(row.iloc[0]).strip()
                            val = row.iloc[1]
                            if not cat: continue
                            if not any(c.isalpha() for c in cat):
                                prod_rows_sd.append((cat, val))
                            else:
                                if val != "" and val is not None:
                                    try:
                                        float(val)
                                        prod_rows_sd.append((cat, val))
                                    except: pass

                        if len(prod_rows_sd) > 30: prod_rows_sd = prod_rows_sd[-30:]

                        for i in range(len(prod_rows_sd)):
                            categories_sd.append(prod_rows_sd[i][0])
                            try: proceso_vals_sd.append(float(prod_rows_sd[i][1]))
                            except: proceso_vals_sd.append(None)
                            diario_vals_sd.append(None)
                            programa_vals_sd.append(None)
                            columna1_vals_sd.append(None)

                        for i in range(dias_mes):
                            categories_sd.append(str(i + 1))
                            proceso_vals_sd.append(None)
                            
                            try: diario_vals_sd.append(float(app.df_data_olme_die[snr_col_olme_die].iloc[i]))
                            except: diario_vals_sd.append(None)
                            
                            try: programa_vals_sd.append(float(app.df_snr_olme_die.iloc[i, 0]))
                            except: programa_vals_sd.append(None)
                            
                            try: columna1_vals_sd.append(float(app.df_snr_olme_die.iloc[i, 1]))
                            except: columna1_vals_sd.append(None)

                        update_slide_chart(chart_olme_die, categories_sd, proceso_vals_sd, diario_vals_sd, programa_vals_sd, columna1_vals_sd, wine_color, green_color)

        # --- NUEVO: ACTUALIZAR TEXTOS DE PROGRAMAS DE REFINERÍAS DESDE ENVIO PRO DIARIO ---
        try:
            if hasattr(app, 'df_envio') and app.df_envio is not None:
                df_envio = app.df_envio
                
                refinery_map = {
                    "cadereyta": "CY",
                    "madero": "MA",
                    "minatit": "MI",
                    "salamanca": "SA",
                    "salina cruz": "SC",
                    "tula": "TU",
                    "olmeca": "OL"
                }

                ref_row_index = {
                    "CY": 0, "MA": 1, "MI": 2, "OL": 3, "SA": 4, "SC": 5, "TU": 6
                }

                def get_envio_val(row, col=32):
                    try:
                        return float(df_envio.iloc[row, col])
                    except:
                        return 0.0

                for slide in prs.slides:
                    for shape in slide.shapes:
                        if shape.has_text_frame and "Programa:" in shape.text:
                            text_lower = shape.text.lower()
                            
                            # Detect refinery
                            ref_code = None
                            for ref_kw, ref_c in refinery_map.items():
                                if ref_kw in text_lower:
                                    ref_code = ref_c
                                    break
                            
                            if not ref_code:
                                continue
                                
                            offset = ref_row_index[ref_code]
                            new_val = None
                            
                            # Detect product
                            if "crudo" in text_lower:
                                new_val = get_envio_val(11 + offset)
                            elif "gasolina" in text_lower:
                                # Sum of Magna (21), Premium RP (31), Premium ZM (41)
                                new_val = get_envio_val(21 + offset) + get_envio_val(31 + offset) + get_envio_val(41 + offset)
                            elif "turbosina" in text_lower:
                                new_val = get_envio_val(71 + offset)
                            elif "diesel" in text_lower:
                                new_val = get_envio_val(81 + offset)
                            elif "combust" in text_lower:
                                new_val = get_envio_val(121 + offset)
                            elif "asfalto" in text_lower:
                                new_val = get_envio_val(131 + offset)
                                
                            if new_val is not None:
                                rounded_val = round(new_val, 1)
                                for p in shape.text_frame.paragraphs:
                                    if "Programa:" in p.text:
                                        if len(p.runs) > 0:
                                            first_run = p.runs[0]
                                            font_name = first_run.font.name
                                            font_size = first_run.font.size
                                            font_bold = first_run.font.bold
                                            font_color = first_run.font.color.rgb if first_run.font.color.type == 1 else None
                                            
                                            import re
                                            p.text = re.sub(r'Programa:\s*[\d\.]+', f'Programa: {rounded_val}', p.text)
                                            
                                            if len(p.runs) > 0:
                                                new_run = p.runs[0]
                                                new_run.font.name = font_name
                                                new_run.font.size = font_size
                                                new_run.font.bold = font_bold
                                                if font_color:
                                                    new_run.font.color.rgb = font_color
                                        else:
                                            import re
                                            p.text = re.sub(r'Programa:\s*[\d\.]+', f'Programa: {rounded_val}', p.text)
        except Exception as e_text:
            print("Error actualizando textos de programas:", e_text)

        prs.save(save_path)


        app.after(0, app.on_pptx_success, save_path)

    except Exception as e:
        err_details = traceback.format_exc()
        app.after(0, app.on_pptx_error, str(e), err_details)


def reset_slide_mv_chart(app, template_path, save_path):
    """Actualiza la gráfica superior de Proceso de Crudo en la diapositiva de Margen Variable.

    Incluye:
      - Únicamente el último año histórico en la BD (ej. 2025) y su valor de proceso.
      - Los meses transcurridos con producción del año actual (Ene..).
      - Los días del mes en curso con sus series Diario (real) y Programa.
    """
    try:
        from pptx import Presentation
        from pptx.chart.data import CategoryChartData
        from pptx.dml.color import RGBColor
        from pptx.util import Pt

        prs = Presentation(template_path)

        # 1. Encontrar la diapositiva de Seguimiento al Margen Variable
        target_slide = None
        for s in prs.slides:
            for shp in s.shapes:
                if shp.has_text_frame and "margen variable del snr" in shp.text_frame.text.lower():
                    target_slide = s
                    break
            if target_slide:
                break

        if not target_slide and len(prs.slides) > 1:
            target_slide = prs.slides[1]

        if not target_slide:
            raise ValueError("No se encontró la diapositiva de seguimiento al Margen Variable en la plantilla.")

        # 2. Encontrar la gráfica superior (Proceso de Crudo)
        chart_shapes = [s for s in target_slide.shapes if s.has_chart]
        if not chart_shapes:
            raise ValueError("No se encontraron gráficas en la diapositiva de Margen Variable.")

        chart_shapes.sort(key=lambda s: s.top)
        top_chart = chart_shapes[0].chart

        # 3. Extraer ÚLTIMO AÑO de df_prod
        sheet_year = getattr(app, "sheet_year", 2026)
        hist_years = []
        for idx, row in app.df_prod.iterrows():
            cat = str(row.iloc[0]).strip()
            val = row.iloc[1]
            if not cat:
                continue
            if not any(c.isalpha() for c in cat):
                try:
                    int(cat)
                    hist_years.append((cat, val))
                except Exception:
                    pass

        last_year, last_year_val = hist_years[-1] if hist_years else (str(sheet_year), 1112)

        # 4. Extraer meses activos del año actual
        active_months = []
        for idx, row in app.df_prod.iterrows():
            cat = str(row.iloc[0]).strip()
            val = row.iloc[1]
            if any(c.isalpha() for c in cat):
                if val != "" and val is not None:
                    try:
                        active_months.append((cat, float(val)))
                    except Exception:
                        pass

        master_current_month = active_months[-1][0] if active_months else "Oct"

        def _round(v):
            if v is None:
                return None
            try:
                return int(round(float(v)))
            except Exception:
                return v

        categories = [str(last_year)]
        proceso_vals = [_round(last_year_val)]
        programa_vals = [None]
        diario_vals = [None]

        for m_name, m_val in active_months:
            categories.append(m_name)
            proceso_vals.append(_round(m_val))
            programa_vals.append(None)
            diario_vals.append(None)

        dias_mes = days_in_current_month(master_current_month, sheet_year)
        proceso_col = app.df_data.columns[1] if app.df_data.shape[1] > 1 else app.df_data.columns[0]

        for i in range(dias_mes):
            categories.append(str(i + 1))
            proceso_vals.append(None)
            try:
                diario_vals.append(_round(app.df_data[proceso_col].iloc[i]))
            except Exception:
                diario_vals.append(None)
            try:
                programa_vals.append(_round(app.df_snr.iloc[i, 0]))
            except Exception:
                programa_vals.append(None)

        # 5. Reemplazar datos en la gráfica
        cd = CategoryChartData()
        cd.categories = categories
        cd.add_series("PROCESO", tuple(proceso_vals))
        cd.add_series("Programa", tuple(programa_vals))
        cd.add_series("Diario", tuple(diario_vals))

        top_chart.replace_data(cd)

        # 6. Formatear colores
        wine_color = RGBColor(0x62, 0x11, 0x32)   # Color vino corporativo Pemex
        green_color = RGBColor(0x24, 0x5C, 0x4F)  # Color verde corporativo Pemex

        if len(top_chart.series) > 0:
            series = top_chart.series[0]

            for p_idx in range(len(series.points)):
                if p_idx >= len(categories):
                    break
                cat = categories[p_idx]
                if cat.isdigit() and len(cat) <= 2:
                    break
                try:
                    point = series.points[p_idx]
                    fill = point.format.fill
                    fill.solid()

                    if p_idx == 0:
                        # El año pasado: DEBE SER COLOR VINO
                        fill.fore_color.rgb = wine_color
                        try:
                            point.data_label.font.size = Pt(10)
                            point.data_label.font.bold = True
                        except Exception:
                            pass
                    else:
                        # Todos los meses (incluyendo el mes actual): DEBEN SER COLOR VERDE
                        fill.fore_color.rgb = green_color
                        try:
                            point.data_label.font.size = Pt(9)
                            point.data_label.font.bold = False
                        except Exception:
                            pass
                except Exception:
                    pass

        # 7. Asegurar etiquetas de mes en la parte inferior de la gráfica
        try:
            existing_month_boxes = {}
            oct_box = None
            for s in target_slide.shapes:
                if s.has_text_frame and 2900000 < s.top < 3150000:
                    txt = s.text_frame.text.strip()
                    existing_month_boxes[txt] = s
                    if txt == "Oct":
                        oct_box = s

            # Actualizar caja del año si existe manteniendo tipografía y tamaño exacto (Montserrat 9pt negrita)
            for yr_key in ["2024", "2025", "2026", "2027"]:
                if yr_key in existing_month_boxes:
                    yr_box = existing_month_boxes[yr_key]
                    if len(yr_box.text_frame.paragraphs) > 0:
                        p = yr_box.text_frame.paragraphs[0]
                        if len(p.runs) > 0:
                            p.runs[0].text = str(last_year)
                            p.runs[0].font.name = "Montserrat"
                            p.runs[0].font.size = Pt(9)
                            p.runs[0].font.bold = True
                        else:
                            r = p.add_run()
                            r.text = str(last_year)
                            r.font.name = "Montserrat"
                            r.font.size = Pt(9)
                            r.font.bold = True

            # Para cada mes activo, si no tiene cuadro de texto abajo, crearlo
            all_meses_abr = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
            step = 291605
            if oct_box is not None:
                oct_left = oct_box.left
                oct_top = oct_box.top
                oct_w = oct_box.width
                oct_h = oct_box.height
                oct_idx = all_meses_abr.index("Oct")

                import copy
                ns = {
                    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
                }
                for m_name, _ in active_months:
                    if m_name not in existing_month_boxes and m_name in all_meses_abr:
                        m_idx = all_meses_abr.index(m_name)
                        offset = m_idx - oct_idx
                        new_left = oct_left + (step * offset)

                        # Clonar el shape XML de Octubre para heredar exactamente la rotación, márgenes, centrado y tipografía
                        clone_sp = copy.deepcopy(oct_box._element)
                        cNvPr = clone_sp.find('.//p:cNvPr', ns)
                        if cNvPr is not None:
                            cNvPr.attrib['id'] = str(5000 + m_idx)
                            cNvPr.attrib['name'] = f"{m_name} CuadroTexto"

                        off = clone_sp.find('.//a:off', ns)
                        if off is not None:
                            off.attrib['x'] = str(new_left)

                        for t in clone_sp.findall('.//a:t', ns):
                            t.text = m_name

                        target_slide.shapes._spTree.append(clone_sp)
                        existing_month_boxes[m_name] = clone_sp
        except Exception as e_lbl:
            print("Aviso al colocar etiquetas de mes inferiores:", e_lbl)

        # 8. Limpiar cuadros de texto sobre las gráficas de las diapositivas 4 a 10 y dejar un cuadro 'prueba' en la esquina inferior izquierda
        try:
            for s_idx in range(3, min(10, len(prs.slides))):
                slide = prs.slides[s_idx]
                charts = [s for s in slide.shapes if s.has_chart]
                if not charts:
                    continue
                chart_shp = charts[0]
                c_left = chart_shp.left
                c_right = chart_shp.left + chart_shp.width
                c_top = chart_shp.top
                c_bottom = chart_shp.top + chart_shp.height

                # Eliminar figuras de texto superpuestas en la gráfica
                for shp in list(slide.shapes):
                    if shp == chart_shp:
                        continue
                    if shp.has_text_frame and not shp.has_table:
                        s_cx = shp.left + shp.width / 2
                        s_cy = shp.top + shp.height / 2
                        if c_left <= s_cx <= c_right and c_top <= s_cy <= c_bottom:
                            el = shp._element
                            el.getparent().remove(el)

                # Agregar un solo cuadro de texto en la esquina inferior izquierda que diga "prueba"
                tb_left = c_left + 25000
                tb_top = c_bottom - 360000
                tb_w = 1100000
                tb_h = 300000
                txBox = slide.shapes.add_textbox(tb_left, tb_top, tb_w, tb_h)
                txBox.line.color.rgb = RGBColor(0x6E, 0x15, 0x2E)
                txBox.line.width = Pt(1)
                tf = txBox.text_frame
                tf.margin_left = Pt(4)
                tf.margin_top = Pt(2)
                tf.margin_right = Pt(4)
                tf.margin_bottom = Pt(2)
                p = tf.paragraphs[0]
                p.text = "prueba"
                p.font.name = "Montserrat"
                p.font.size = Pt(9)
                p.font.bold = True
                p.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        except Exception as e_clean:
            print("Aviso al limpiar cuadros de texto sobre gráficas de refinerías:", e_clean)

        # 9. Extraer CMP desde el Excel (columnas AA y AB), realizar barrido de tablas y resetear gráficas de las 7 refinerías
        try:
            excel_path = getattr(app, "current_file_path", None)
            if not excel_path or not os.path.exists(excel_path):
                cand = "assets/2026 Datos para MV Nov comparativo con rend.xlsm"
                if os.path.exists(cand):
                    excel_path = cand

            # Barrido completo de tablas en las 7 refinerías
            updated_tables_count, cmp_data = sweep_and_update_refinery_tables(prs, excel_path=excel_path)

            # Si hay archivo de Cargas disponible, también actualizar tablas de Cargas, Productos e Insumos
            cargas_file = getattr(app, "cargas_excel_path", None)
            if not cargas_file or not os.path.exists(cargas_file):
                cand_cargas = "assets/Cargas Nov 26 Auto.xlsm"
                if os.path.exists(cand_cargas):
                    cargas_file = cand_cargas
            if cargas_file and os.path.exists(cargas_file):
                try:
                    sweep_and_update_cargas_tables(prs, cargas_file)
                except Exception as e_cargas:
                    print("Aviso al actualizar tablas de cargas en reseteo:", e_cargas)

            # Resetear las gráficas en las diapositivas de cada refinería
            for s_idx, slide in enumerate(prs.slides):
                matched_ref = None
                for s in slide.shapes:
                    if s.has_text_frame and s.text_frame.text:
                        txt = s.text_frame.text.lower()
                        for ref_key in cmp_data:
                            if ref_key in txt:
                                matched_ref = ref_key
                                break
                    if matched_ref:
                        break

                if not matched_ref:
                    for ref_key, info in cmp_data.items():
                        if info.get("slide_idx") == s_idx:
                            matched_ref = ref_key
                            break

                if not matched_ref:
                    continue

                ref_info = cmp_data[matched_ref]
                val_aa = ref_info.get("val_aa")

                charts = [s for s in slide.shapes if s.has_chart]
                if charts:
                    chart_shp = charts[0]
                    chart_elem = chart_shp.chart._element
                    line_chart = chart_elem.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}lineChart')
                    if line_chart is not None:
                        ser_list = line_chart.findall('{http://schemas.openxmlformats.org/drawingml/2006/chart}ser')
                        if ser_list:
                            # Serie 0: CMP constante en todos los días
                            ser0 = ser_list[0]
                            val0 = ser0.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}val')
                            if val0 is not None and val_aa is not None:
                                str_aa = str(val_aa)
                                for pt in val0.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}pt'):
                                    v = pt.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}v')
                                    if v is not None:
                                        v.text = str_aa

                            # Series 1+: valores reales a quitar
                            for ser in ser_list[1:]:
                                # Quitar etiquetas de datos (data labels)
                                dlbls = ser.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}dLbls')
                                if dlbls is not None:
                                    for dlbl in dlbls.findall('{http://schemas.openxmlformats.org/drawingml/2006/chart}dLbl'):
                                        dlbls.remove(dlbl)

                                # Quitar puntos
                                val = ser.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}val')
                                if val is not None:
                                    cache = val.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}numCache')
                                    if cache is not None:
                                        pt_count = cache.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}ptCount')
                                        if pt_count is not None:
                                            pt_count.set('val', '0')
                                        for pt in cache.findall('{http://schemas.openxmlformats.org/drawingml/2006/chart}pt'):
                                            cache.remove(pt)
                                    lit = val.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}numLit')
                                    if lit is not None:
                                        pt_count = lit.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}ptCount')
                                        if pt_count is not None:
                                            pt_count.set('val', '0')
                                        for pt in lit.findall('{http://schemas.openxmlformats.org/drawingml/2006/chart}pt'):
                                            lit.remove(pt)
        except Exception as e_cmp:
            print("Aviso al actualizar CMP y resetear gráficas de refinerías:", e_cmp)

        prs.save(save_path)
        app.after(0, app.on_reset_slide_success, save_path)

    except Exception as e:
        err_details = traceback.format_exc()
        app.after(0, app.on_reset_slide_error, str(e), err_details)


def _detect_slide_refinery(slide, s_idx, keys, default_idx_map):
    """Devuelve la clave de refinería de una diapositiva.

    Primero busca un cuadro de texto cuyo contenido (normalizado) sea el nombre de la
    refinería (el título de la diapositiva). Si no hay título, usa la posición por defecto
    (diapositivas 4-10). No busca el nombre dentro de textos largos para no confundir
    diapositivas de resumen (p. ej. la 2) con diapositivas de refinería.
    """
    def _n(t):
        t = unicodedata.normalize("NFKD", str(t or "").strip())
        t = "".join(c for c in t if not unicodedata.combining(c))
        return _re.sub(r"\s+", " ", t).lower()

    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            t = _n(sh.text_frame.text)
            if len(t) > 25:
                continue
            for k in keys:
                if t == k:
                    return k
    for k, idx in default_idx_map.items():
        if idx == s_idx and k in keys:
            return k
    return None


def sweep_and_update_refinery_tables(prs, excel_path=None, cmp_data=None):
    """Realiza el barrido de tablas de Margen Variable y Utilidad en las diapositivas de cada refinería.

    Para cada refinería (Cadereyta, Madero, Minatitlán, Salamanca, Salina Cruz, Tula, Olmeca):
    - Extrae de las columnas AA y AB del Excel los valores CMP correspondientes:
        Cadereyta (Fila 124), Madero (Fila 129), Minatitlán (Fila 134), Salamanca (Fila 139),
        Salina Cruz (Fila 144), Tula (Fila 149), Olmeca (Fila 154)
    - Busca la tabla con encabezado 'Margen Variable' y 'Utilidad'
    - Actualiza la fila CMP con los valores de AA y AB conservando formato tipográfico original.
    Retorna (updated_count, cmp_data).
    """
    import openpyxl

    if cmp_data is None:
        if not excel_path or not os.path.exists(excel_path):
            cand = "assets/2026 Datos para MV Nov comparativo con rend.xlsm"
            if os.path.exists(cand):
                excel_path = cand

        if not excel_path or not os.path.exists(excel_path):
            raise ValueError("No se encontró el archivo Excel de Margen Variable para extraer los valores CMP.")

        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
        target_sheet = None
        for s_name in ["EnvioCalculoPromedio", "Calculo promedio", "Calculo MV SNR"]:
            if s_name in wb.sheetnames:
                ws_cand = wb[s_name]
                if ws_cand.cell(row=124, column=27).value is not None:
                    target_sheet = ws_cand
                    break
        if target_sheet is None:
            for s_name in wb.sheetnames:
                ws_cand = wb[s_name]
                v = ws_cand.cell(row=124, column=27).value
                if v is not None and isinstance(v, (int, float)):
                    target_sheet = ws_cand
                    break

        if target_sheet is None:
            raise ValueError("No se encontró la hoja con los datos CMP (filas 124..154, cols AA/AB) en el Excel.")

        refineries_cfg = [
            {"name": "cadereyta", "row": 124, "slide_idx": 3},
            {"name": "madero", "row": 129, "slide_idx": 4},
            {"name": "minatitl", "row": 134, "slide_idx": 5},
            {"name": "salamanca", "row": 139, "slide_idx": 6},
            {"name": "salina cruz", "row": 144, "slide_idx": 7},
            {"name": "tula", "row": 149, "slide_idx": 8},
            {"name": "olmeca", "row": 154, "slide_idx": 9},
        ]

        cmp_data = {}
        for cfg in refineries_cfg:
            val_aa = target_sheet.cell(row=cfg["row"], column=27).value
            val_ab = target_sheet.cell(row=cfg["row"], column=28).value
            cmp_data[cfg["name"]] = {
                "val_aa": val_aa,
                "val_ab": val_ab,
                "slide_idx": cfg["slide_idx"],
            }

    def _update_table_cell(cell, new_val_str):
        if cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
            cell.text_frame.paragraphs[0].runs[0].text = new_val_str
        else:
            cell.text = new_val_str

    updated_count = 0
    default_map = {k: v.get("slide_idx") for k, v in cmp_data.items()}
    for s_idx, slide in enumerate(prs.slides):
        matched_ref = _detect_slide_refinery(slide, s_idx, list(cmp_data.keys()), default_map)
        if not matched_ref:
            continue

        ref_info = cmp_data[matched_ref]
        val_aa = ref_info.get("val_aa")
        val_ab = ref_info.get("val_ab")

        for shp in slide.shapes:
            if shp.has_table:
                tbl = shp.table
                if len(tbl.rows) >= 2 and len(tbl.columns) >= 3:
                    header_txt = tbl.cell(0, 1).text.lower()
                    if "margen variable" in header_txt or "variable" in header_txt or "utilidad" in header_txt:
                        # Fila 1 es CMP
                        if val_aa is not None:
                            try:
                                txt_aa = f"{float(val_aa):.2f}"
                            except Exception:
                                txt_aa = str(val_aa)
                            _update_table_cell(tbl.cell(1, 1), txt_aa)

                        if val_ab is not None:
                            try:
                                txt_ab = f"{float(val_ab):.3f}"
                            except Exception:
                                txt_ab = str(val_ab)
                            _update_table_cell(tbl.cell(1, 2), txt_ab)

                        updated_count += 1

    return updated_count, cmp_data


def sweep_and_update_cargas_tables(prs, cargas_excel_path, warnings=None):
    """Lee la hoja 'Tablas' del Excel de Cargas (ej. Cargas Nov 26 Auto.xlsm) y actualiza
    las tablas de Cargas, Productos e Insumos en las diapositivas de cada refinería.

    Excluye estrictamente la sección 'Variación de Inventarios Intermedios / Transferencias'.
    """
    import openpyxl

    wb = openpyxl.load_workbook(cargas_excel_path, data_only=True, read_only=True)
    if "Tablas" not in wb.sheetnames:
        raise ValueError(f"No se encontró la hoja 'Tablas' en {os.path.basename(cargas_excel_path)}")

    _rows = [list(r) for r in wb["Tablas"].iter_rows(values_only=True)]

    # El Excel usa TODAY()-1 como fecha de corte (Cadereyta!D2). openpyxl solo lee los valores
    # guardados la última vez en Excel; si el archivo no se guardó hoy, los números quedan viejos.
    if warnings is not None and "Cadereyta" in wb.sheetnames:
        try:
            import datetime as _dt
            d2 = next(wb["Cadereyta"].iter_rows(min_row=2, max_row=2, min_col=4, max_col=4, values_only=True))[0]
            if isinstance(d2, _dt.datetime):
                d2 = d2.date()
            expected = _dt.date.today() - _dt.timedelta(days=1)
            if isinstance(d2, _dt.date) and d2 < expected:
                warnings.append(
                    f"El Excel de Cargas fue guardado con fecha de corte {d2:%d/%m/%Y}, pero hoy el corte sería "
                    f"{expected:%d/%m/%Y}. Los valores pueden diferir de lo que ves en pantalla: abre el Excel, "
                    f"guárdalo (Cmd+S) y vuelve a enviar las tablas."
                )
        except Exception:
            pass

    class _WS:
        max_row = len(_rows)

        @staticmethod
        def cell(row, column):
            class _C:
                pass
            c = _C()
            try:
                c.value = _rows[row - 1][column - 1]
            except IndexError:
                c.value = None
            return c

    ws = _WS()

    def _norm(s):
        if s is None:
            return ""
        s = str(s).strip()
        s = unicodedata.normalize("NFKD", s)
        s = "".join(c for c in s if not unicodedata.combining(c))
        s = _re.sub(r"\s+", " ", s)
        return s.lower()

    def _format_val(val, is_pct=False):
        if val is None or val == "":
            return ""
        if isinstance(val, str):
            val_str = val.strip()
            if val_str in [">100", "NA"]:
                return val_str
            if "#" in val_str:
                return "NA"
            try:
                val = float(val_str)
            except ValueError:
                return val_str
        try:
            f_val = float(val)
            if is_pct and f_val > 100:
                return ">100"
            out = f"{f_val:.1f}"
            return "0.0" if out == "-0.0" else out
        except Exception:
            return str(val)

    def _update_cell(cell, text):
        p = cell.text_frame.paragraphs[0] if cell.text_frame.paragraphs else cell.text_frame.add_paragraph()
        if p.runs:
            p.runs[0].text = str(text)
            for r in p.runs[1:]:
                r.text = ""
        else:
            cell.text = str(text)

    # 1. Detectar inicio de cada refinería en la hoja Tablas
    ref_names = ["cadereyta", "madero", "minatitlan", "salamanca", "salina cruz", "tula", "olmeca"]
    starts = {}
    max_scan = ws.max_row + 1
    for r in range(1, max_scan):
        b = _norm(ws.cell(row=r, column=2).value)
        h = _norm(ws.cell(row=r, column=8).value)
        for ref in ref_names:
            if ref not in starts and (ref in b or ref in h):
                starts[ref] = r

    sorted_starts = sorted(starts.items(), key=lambda x: x[1])
    ref_ranges = {}
    for idx, (ref, start_r) in enumerate(sorted_starts):
        end_r = sorted_starts[idx + 1][1] - 1 if idx + 1 < len(sorted_starts) else min(start_r + 30, ws.max_row)
        ref_ranges[ref] = (start_r, end_r)

    # 2. Extraer datos para cada refinería
    cargas_by_ref = {}
    for ref, (r_start, r_end) in ref_ranges.items():
        cargas = {}
        productos = {}
        insumos = {}
        for r in range(r_start, r_end + 1):
            # Columna 2: Cargas
            b = ws.cell(row=r, column=2).value
            if b and "proceso" not in _norm(b) and _norm(b) != ref:
                cargas[_norm(b)] = {
                    "cmp": ws.cell(row=r, column=3).value,
                    "real": ws.cell(row=r, column=4).value,
                    "pct": ws.cell(row=r, column=5).value,
                    "cap": ws.cell(row=r, column=6).value,
                }

            # Columna 8: Productos
            h = ws.cell(row=r, column=8).value
            if h and "producto" not in _norm(h) and _norm(h) != ref:
                productos[_norm(h)] = {
                    "cmp": ws.cell(row=r, column=9).value,
                    "real": ws.cell(row=r, column=10).value,
                    "pct": ws.cell(row=r, column=11).value,
                    "rend": ws.cell(row=r, column=12).value,
                }

            # Columna 14: Insumos
            n = ws.cell(row=r, column=14).value
            if n and "insumos" not in _norm(n) and _norm(n) != ref:
                insumos[_norm(n)] = {
                    "cmp": ws.cell(row=r, column=15).value,
                    "real": ws.cell(row=r, column=16).value,
                    "pct": ws.cell(row=r, column=17).value,
                }

        cargas_by_ref[ref] = {
            "cargas": cargas,
            "productos": productos,
            "insumos": insumos,
        }

    # 3. Mapear a diapositivas y actualizar tablas
    updated_ref_count = 0
    slide_ref_map = {
        "cadereyta": 3,
        "madero": 4,
        "minatitlan": 5,
        "salamanca": 6,
        "salina cruz": 7,
        "tula": 8,
        "olmeca": 9,
    }

    for s_idx, slide in enumerate(prs.slides):
        matched_ref = _detect_slide_refinery(slide, s_idx, ref_names, slide_ref_map)
        if not matched_ref or matched_ref not in cargas_by_ref:
            continue

        ref_data = cargas_by_ref[matched_ref]
        cargas_dict = ref_data["cargas"]
        prod_dict = ref_data["productos"]
        insumo_dict = ref_data["insumos"]

        for shp in slide.shapes:
            if not shp.has_table:
                continue
            tbl = shp.table
            h0 = _norm(tbl.cell(0, 0).text)

            # A. Tabla de Cumplimiento del programa de Cargas
            if "cumplimiento" in h0 or "cargas" in h0:
                for r_i in range(1, len(tbl.rows)):
                    proc = _norm(tbl.cell(r_i, 0).text)
                    if not proc or "proceso" in proc:
                        continue
                    if proc in cargas_dict:
                        info = cargas_dict[proc]
                        if info.get("cmp") is not None and len(tbl.columns) > 1:
                            _update_cell(tbl.cell(r_i, 1), _format_val(info["cmp"]))
                        if info.get("real") is not None and len(tbl.columns) > 2:
                            _update_cell(tbl.cell(r_i, 2), _format_val(info["real"]))
                        if info.get("pct") is not None and len(tbl.columns) > 3:
                            _update_cell(tbl.cell(r_i, 3), _format_val(info["pct"], is_pct=True))
                        if info.get("cap") is not None and len(tbl.columns) > 4:
                            _update_cell(tbl.cell(r_i, 4), _format_val(info["cap"]))

            # B. Tabla de Productos e Insumos
            elif "producto" in h0:
                skip_inv = False
                in_insumos = False
                for r_i in range(1, len(tbl.rows)):
                    p_name = _norm(tbl.cell(r_i, 0).text)

                    # Omitir sección de Variación de Inventarios Intermedios / Transferencias
                    if "variacion de inventarios" in p_name:
                        skip_inv = True
                        continue

                    if skip_inv:
                        if p_name in insumo_dict or "insumos" in p_name or p_name in ["mtbe", "metanol", "isobutano", "duba", "muba zmm", "muba rp", "hidrogeno", "propano", "mezcla licuables", "premium uba zmvm imp."]:
                            skip_inv = False
                            in_insumos = True
                        else:
                            continue

                    if not p_name or "producciones" in p_name or "insumos" in p_name:
                        if "insumos" in p_name:
                            in_insumos = True
                        continue

                    # Si estamos en la sección de insumos, priorizar insumos
                    info = None
                    if in_insumos:
                        info = insumo_dict.get(p_name) or prod_dict.get(p_name)
                    else:
                        info = prod_dict.get(p_name) or insumo_dict.get(p_name)

                    if info:
                        if info.get("cmp") is not None and len(tbl.columns) > 1:
                            _update_cell(tbl.cell(r_i, 1), _format_val(info["cmp"]))
                        if info.get("real") is not None and len(tbl.columns) > 2:
                            _update_cell(tbl.cell(r_i, 2), _format_val(info["real"]))
                        if info.get("pct") is not None and len(tbl.columns) > 3:
                            _update_cell(tbl.cell(r_i, 3), _format_val(info["pct"], is_pct=True))
                        if info.get("rend") is not None and len(tbl.columns) > 4 and not in_insumos:
                            _update_cell(tbl.cell(r_i, 4), _format_val(info["rend"]))

        updated_ref_count += 1

    return updated_ref_count


def update_refinery_tables_mv(app, pptx_path, save_path, cargas_excel_path=None):
    """Actualiza las tablas de las refinerías en la presentación.

    1. Actualiza la tabla CMP (Margen Variable / Utilidad) con columnas AA y AB del Excel de Margen Variable.
    2. Si se proporciona cargas_excel_path, actualiza las tablas de Cargas, Productos e Insumos
       leyendo la hoja 'Tablas' (excluyendo Variación de Inventarios Intermedios / Transferencias).
    """
    try:
        from pptx import Presentation

        prs = Presentation(pptx_path)

        # 1. Tablas de Cargas / Productos / Insumos (hoja 'Tablas' del Excel de Cargas)
        stale_warnings = []
        cargas_count = 0
        if cargas_excel_path and os.path.exists(cargas_excel_path):
            cargas_count = sweep_and_update_cargas_tables(prs, cargas_excel_path, warnings=stale_warnings)

        # 2. Tabla CMP Margen Variable / Utilidad (opcional: si falla no bloquea el paso 1)
        updated_cmp_count = 0
        warning = ""
        try:
            excel_path = getattr(app, "current_file_path", None)
            if not excel_path or not os.path.exists(excel_path):
                cand = "assets/2026 Datos para MV Nov comparativo con rend.xlsm"
                excel_path = cand if os.path.exists(cand) else None
            if not excel_path:
                raise ValueError("No se encontró el Excel de Margen Variable.")
            updated_cmp_count, _ = sweep_and_update_refinery_tables(prs, excel_path=excel_path)
        except Exception as e_cmp:
            warning = f"No se actualizó la tabla de Margen Variable/Utilidad: {e_cmp}"
            print("Aviso:", warning)
        warning = "\n".join(stale_warnings + ([warning] if warning else []))

        prs.save(save_path)
        app.after(0, app.on_enviar_tablas_success, save_path, updated_cmp_count, cargas_count, warning)

    except Exception as e:
        err_details = traceback.format_exc()
        app.after(0, app.on_enviar_tablas_error, str(e), err_details)


