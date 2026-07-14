import traceback
import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor

def update_slide_chart(chart, categories, proceso_vals, diario_vals, programa_vals, columna1_vals, wine_color, green_color):
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
        last_month_idx = -1
        for i in range(17):
            if i < len(categories):
                cat = categories[i]
                val = proceso_vals[i]
                # Si es un mes (contiene letras)
                if any(c.isalpha() for c in cat):
                    if val is not None and val != 0:
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

        # Pintar todas las barras
        for p_idx in range(min(17, len(series.points))):
            try:
                cat = categories[p_idx]
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
                    # Si es un año, aplicamos el gris explícitamente
                    fill.fore_color.rgb = gray_color
                    
            except Exception:
                pass


def export_to_pptx(app, file_path, save_path):
    try:
        from pptx import Presentation
        from pptx.dml.color import RGBColor

        prs = Presentation(file_path)
        if len(prs.slides) < 18:
            raise ValueError("La presentación debe tener al menos 18 diapositivas (incluyendo las de Madero -Crudo, Gasolinas y Diesel).")
 
        # --- 1. PROCESAR DIAPOSITIVA DE CRUDO (DIAPOSITIVA 2) ---
        slide = prs.slides[1]
        chart = None

        for shape in slide.shapes:
            if shape.has_chart:
                chart = shape.chart
                break

        if not chart:
            raise ValueError("No se encontró ninguna gráfica en la segunda diapositiva (Crudo).")

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
                try:
                    p_val = float(val)
                    if p_val != 0: prod_rows.append((cat, val))
                except: pass
        
        if len(prod_rows) > 30: prod_rows = prod_rows[-30:]

        for i in range(len(prod_rows)):
            categories.append(prod_rows[i][0])
            try: proceso_vals.append(float(prod_rows[i][1]))
            except: proceso_vals.append(None)
            diario_vals.append(None)
            programa_vals.append(None)
            columna1_vals.append(None)

        # Llenar datos diarios (31 días)
        for i in range(31):
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
        slide_cad = prs.slides[9]
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
                    try:
                        p_val = float(val)
                        if p_val != 0: prod_rows_cad.append((cat, val))
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

            for i in range(31):
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
        slide_gas = prs.slides[2]
        chart_gas = None
        for shape in slide_gas.shapes:
            if shape.has_chart:
                chart_gas = shape.chart
                break

        if not chart_gas:
            raise ValueError("No se encontró ninguna gráfica en la tercera diapositiva (Gasolinas).")

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
                try:
                    p_val = float(val)
                except:
                    p_val = 0
                if p_val != 0:
                    prod_rows_gas.append((cat, val))

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
        for i in range(31):
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
        slide_die = prs.slides[3]
        chart_die = None
        for shape in slide_die.shapes:
            if shape.has_chart:
                chart_die = shape.chart
                break

        if not chart_die:
            raise ValueError("No se encontró ninguna gráfica en la cuarta diapositiva (Diesel).")

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
                try:
                    p_val = float(val)
                except:
                    p_val = 0
                if p_val != 0:
                    prod_rows_die.append((cat, val))

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
        for i in range(31):
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
        slide_turb = prs.slides[4]
        chart_turb = None
        for shape in slide_turb.shapes:
            if shape.has_chart:
                chart_turb = shape.chart
                break

        if not chart_turb:
            raise ValueError("No se encontró ninguna gráfica en la quinta diapositiva (Turbosina).")

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
                try:
                    p_val = float(val)
                except:
                    p_val = 0
                if p_val != 0:
                    prod_rows_turb.append((cat, val))

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
        for i in range(31):
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
            slide_asf = prs.slides[5]
            chart_asf = None
            for shape in slide_asf.shapes:
                if shape.has_chart:
                    chart_asf = shape.chart
                    break

            if not chart_asf:
                raise ValueError("No se encontró ninguna gráfica en la sexta diapositiva (Asfalto).")

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
                    try:
                        p_val = float(val)
                    except:
                        p_val = 0
                    if p_val != 0:
                        prod_rows_asf.append((cat, val))

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
            for i in range(31):
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
            update_slide_chart(chart_asf, categories_asf, proceso_vals_asf, diario_vals_asf, programa_vals_asf, columna1_vals_asf, wine_color, green_color)


        # --- 6. PROCESAR DIAPOSITIVA DE COMBUSTOLEO (DIAPOSITIVA 7) ---
        if app.df_data_combustoleo is not None and app.df_snr_combustoleo is not None and app.df_prod_combustoleo is not None:
            slide_comb = prs.slides[6]
            chart_comb = None
            for shape in slide_comb.shapes:
                if shape.has_chart:
                    chart_comb = shape.chart
                    break

            if not chart_comb:
                raise ValueError("No se encontró ninguna gráfica en la séptima diapositiva (Combustoleo).")

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
                    try:
                        p_val = float(val)
                    except:
                        p_val = 0
                    if p_val != 0:
                        prod_rows_comb.append((cat, val))

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
            for i in range(31):
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
            update_slide_chart(chart_comb, categories_comb, proceso_vals_comb, diario_vals_comb, programa_vals_comb, columna1_vals_comb, wine_color, green_color)
 
            # --- 7. PROCESAR DIAPOSITIVA DE GASOLINAS CADEREYTA (DIAPOSITIVA 11) ---
            if app.df_data_cad_gas is not None and app.df_snr_cad_gas is not None and app.df_prod_cad_gas is not None:
                slide_cad_gas = prs.slides[10]
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
                                try:
                                    if float(val) != 0: prod_rows_cg.append((cat, val))
                                except: pass





                        if len(prod_rows_cg) > 30: prod_rows_cg = prod_rows_cg[-30:]

                        for i in range(len(prod_rows_cg)):
                            categories_cg.append(prod_rows_cg[i][0])
                            try: proceso_vals_cg.append(float(prod_rows_cg[i][1]))
                            except: proceso_vals_cg.append(None)
                            diario_vals_cg.append(None)
                            programa_vals_cg.append(None)
                            columna1_vals_cg.append(None)
 
                        # Llenar datos diarios (31 días)
                        for i in range(31):
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
                    slide_cad_die = prs.slides[11]
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
                                    try:
                                        if float(val) != 0: prod_rows_cd.append((cat, val))
                                    except: pass
                            
                            if len(prod_rows_cd) > 30: prod_rows_cd = prod_rows_cd[-30:]
 
                            for i in range(len(prod_rows_cd)):
                                categories_cd.append(prod_rows_cd[i][0])
                                try: proceso_vals_cd.append(float(prod_rows_cd[i][1]))
                                except: proceso_vals_cd.append(None)
                                diario_vals_cd.append(None)
                                programa_vals_cd.append(None)
                                columna1_vals_cd.append(None)
 
                            # Llenar datos diarios (31 días)
                            for i in range(31):
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
                if len(prs.slides) > 12 and app.df_data_cad_comb is not None and app.df_snr_cad_comb is not None and app.df_prod_cad_comb is not None:
                    slide_cad_comb = prs.slides[12]
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
                                    try:
                                        if float(val) != 0: prod_rows_cc.append((cat, val))
                                    except: pass
                            
                            if len(prod_rows_cc) > 30: prod_rows_cc = prod_rows_cc[-30:]
 
                            for i in range(len(prod_rows_cc)):
                                categories_cc.append(prod_rows_cc[i][0])
                                try: proceso_vals_cc.append(float(prod_rows_cc[i][1]))
                                except: proceso_vals_cc.append(None)
                                diario_vals_cc.append(None)
                                programa_vals_cc.append(None)
                                columna1_vals_cc.append(None)
 
                            # Llenar datos diarios (31 días)
                            for i in range(31):
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
            if len(prs.slides) > 15 and app.df_data_mad_crud is not None and app.df_snr_mad_crud is not None and app.df_prod_mad_crud is not None:
                slide_mad_crud = prs.slides[15]
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
                            try:
                                if float(val) != 0: prod_rows_mc.append((cat, val))
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

                    for i in range(31):
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
            if len(prs.slides) > 16 and app.df_data_mad_gas is not None and app.df_snr_mad_gas is not None and app.df_prod_mad_gas is not None:
                slide_mad_gas = prs.slides[16]
                chart_mad_gas = None
                for shape in slide_mad_gas.shapes:
                    if shape.has_chart:
                        chart_mad_gas = shape.chart
                        break
                
                if chart_mad_gas:
                    snr_col_mad_gas = None
                    for col in app.df_data_mad_gas.columns:
                        if "SNR" in str(col).upper() or "MADERO" in str(col).upper() or "PRODUCCIÓN" in str(col).upper():
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
                                try:
                                    if float(val) != 0: prod_rows_mg.append((cat, val))
                                except: pass

                        if len(prod_rows_mg) > 30: prod_rows_mg = prod_rows_mg[-30:]

                        for i in range(len(prod_rows_mg)):
                            categories_mg.append(prod_rows_mg[i][0])
                            try: proceso_vals_mg.append(float(prod_rows_mg[i][1]))
                            except: proceso_vals_mg.append(None)
                            diario_vals_mg.append(None)
                            programa_vals_mg.append(None)
                            columna1_vals_mg.append(None)

                        for i in range(31):
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
            if len(prs.slides) > 17 and app.df_data_mad_die is not None and app.df_snr_mad_die is not None and app.df_prod_mad_die is not None:
                slide_mad_die = prs.slides[17]
                chart_mad_die = None
                for shape in slide_mad_die.shapes:
                    if shape.has_chart:
                        chart_mad_die = shape.chart
                        break
                
                if chart_mad_die:
                    snr_col_mad_die = None
                    for col in app.df_data_mad_die.columns:
                        if "SNR" in str(col).upper() or "MADERO" in str(col).upper() or "DIE" in str(col).upper():
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
                                try:
                                    if float(val) != 0: prod_rows_md.append((cat, val))
                                except: pass

                        if len(prod_rows_md) > 30: prod_rows_md = prod_rows_md[-30:]

                        for i in range(len(prod_rows_md)):
                            categories_md.append(prod_rows_md[i][0])
                            try: proceso_vals_md.append(float(prod_rows_md[i][1]))
                            except: proceso_vals_md.append(None)
                            diario_vals_md.append(None)
                            programa_vals_md.append(None)
                            columna1_vals_md.append(None)

                        for i in range(31):
                            categories_md.append(str(i + 1))
                            proceso_vals_md.append(None)
                            
                            try: diario_vals_md.append(float(app.df_data_mad_die[snr_col_mad_die].iloc[i]))
                            except: diario_vals_md.append(None)
                            
                            try: programa_vals_md.append(float(app.df_snr_mad_die.iloc[i, 0]))
                            except: programa_vals_md.append(None)
                            
                            try: columna1_vals_md.append(float(app.df_snr_mad_die.iloc[i, 1]))
                            except: columna1_vals_md.append(None)

                        update_slide_chart(chart_mad_die, categories_md, proceso_vals_md, diario_vals_md, programa_vals_md, columna1_vals_md, wine_color, green_color)

        prs.save(save_path)


        app.after(0, app.on_pptx_success, save_path)

    except Exception as e:
        err_details = traceback.format_exc()
        app.after(0, app.on_pptx_error, str(e), err_details)

