"""Diálogos 'Agregar año extra' y 'Coordenadas Excel' con el mismo diseño que la ventana principal."""
import re
import threading
from datetime import datetime
from tkinter import messagebox

import customtkinter as ctk

import db_helper

# ═══ Paleta (igual que app.py) ═══
BG = "#14161b"
CARD = "#1c1f27"
BORDER = "#2a2e39"
ACCENT, ACCENT_H = "#3b82f6", "#2563eb"
NEUTRAL, NEUTRAL_H = "#2b303b", "#363c4a"
DANGER, DANGER_H = "#ef4444", "#dc2626"
OK = "#4ade80"
MUTED, TEXT = "#8b93a5", "#e6e8ee"

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def _button(parent, text, command, kind="neutral", width=140, height=34):
    styles = {
        "primary": dict(fg_color=ACCENT, hover_color=ACCENT_H, text_color="#ffffff", border_width=0),
        "neutral": dict(fg_color=NEUTRAL, hover_color=NEUTRAL_H, text_color=TEXT, border_width=1, border_color=BORDER),
        "ghost": dict(fg_color="transparent", hover_color=NEUTRAL_H, text_color=TEXT, border_width=1, border_color=BORDER),
        "danger": dict(fg_color="transparent", hover_color="#3a1d22", text_color=DANGER, border_width=1, border_color="#5a2a31"),
    }
    return ctk.CTkButton(parent, text=text, command=command, width=width, height=height, corner_radius=8,
                         font=("Roboto", 12, "bold"), **styles[kind])


def _entry(parent, width=140, placeholder=""):
    return ctk.CTkEntry(parent, width=width, height=34, corner_radius=8, fg_color=NEUTRAL,
                        border_color=BORDER, text_color=TEXT, placeholder_text=placeholder,
                        placeholder_text_color=MUTED, font=("Roboto", 12))


def _combo(parent, values, width=260, command=None):
    return ctk.CTkComboBox(parent, values=values, width=width, height=34, corner_radius=8, state="readonly",
                           fg_color=NEUTRAL, border_color=BORDER, button_color=NEUTRAL_H,
                           button_hover_color=ACCENT, font=("Roboto", 12), command=command)


def _label(parent, text, muted=False, bold=False, size=12):
    return ctk.CTkLabel(parent, text=text, text_color=MUTED if muted else TEXT,
                        font=("Roboto", size, "bold" if bold else "normal"))


def _card(parent, **pack):
    f = ctk.CTkFrame(parent, corner_radius=12, fg_color=CARD, border_width=1, border_color=BORDER)
    f.pack(fill="x", **pack)
    return f


def _make_dialog(app, title, width, height):
    dialog = ctk.CTkToplevel(app)
    dialog.title(title)
    dialog.configure(fg_color=BG)
    app.update_idletasks()
    x = app.winfo_x() + (app.winfo_width() - width) // 2
    y = app.winfo_y() + (app.winfo_height() - height) // 2
    dialog.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")
    dialog.resizable(False, False)
    dialog.transient(app)
    dialog.after(120, lambda: _safe_grab(dialog))
    return dialog


def _safe_grab(dialog):
    try:
        dialog.grab_set()
        dialog.focus_set()
    except Exception:
        pass


def _header(dialog, title, subtitle):
    box = ctk.CTkFrame(dialog, fg_color="transparent")
    box.pack(fill="x", padx=22, pady=(18, 12))
    _label(box, title, bold=True, size=18).pack(anchor="w")
    _label(box, subtitle, muted=True).pack(anchor="w", pady=(2, 0))


def _reload(app, message):
    path = getattr(app, "current_file_path", None)
    if path:
        app.set_loading_state(True, message)
        threading.Thread(target=app.async_load_data, args=(path,), daemon=True).start()


def _process_names(app):
    return [v for v in app.cb_proceso.cget("values") if v != "Titulares"]


# ════════════════════════════════════════════════════════════════════════
#  AGREGAR AÑO / MES EXTRA
# ════════════════════════════════════════════════════════════════════════
def open_add_year_dialog(app):
    dialog = _make_dialog(app, "Producción extra", 600, 700)
    _header(dialog, "Producción extra",
            "Agrega o corrige meses y años que no vienen en el Excel (se guardan en la base de datos local).")

    sheet_year = getattr(app, "sheet_year", None) or datetime.now().year
    state = {"dirty": False}

    # ── Formulario ──
    form = _card(dialog, padx=22, pady=(0, 10))
    grid = ctk.CTkFrame(form, fg_color="transparent")
    grid.pack(fill="x", padx=16, pady=14)
    grid.grid_columnconfigure(1, weight=1)

    procesos = _process_names(app)
    _label(grid, "Proceso", muted=True).grid(row=0, column=0, sticky="w", pady=6)
    cb_proc = _combo(grid, procesos, width=300)
    cb_proc.grid(row=0, column=1, sticky="e", pady=6)
    current = app.cb_proceso.get()
    cb_proc.set(current if current in procesos else procesos[0])

    _label(grid, "Tipo", muted=True).grid(row=1, column=0, sticky="w", pady=6)
    seg = ctk.CTkSegmentedButton(grid, values=["Mes", "Año completo"], width=300, height=34,
                                 fg_color=NEUTRAL, selected_color=ACCENT, selected_hover_color=ACCENT_H,
                                 unselected_color=NEUTRAL, unselected_hover_color=NEUTRAL_H,
                                 font=("Roboto", 12, "bold"))
    seg.set("Mes")
    seg.grid(row=1, column=1, sticky="e", pady=6)

    _label(grid, "Año", muted=True).grid(row=2, column=0, sticky="w", pady=6)
    ent_anio = _entry(grid, width=300, placeholder=f"Ej. {sheet_year}")
    ent_anio.grid(row=2, column=1, sticky="e", pady=6)
    ent_anio.insert(0, str(sheet_year))

    lbl_mes = _label(grid, "Mes", muted=True)
    lbl_mes.grid(row=3, column=0, sticky="w", pady=6)
    cb_mes = _combo(grid, MESES, width=300)
    cb_mes.grid(row=3, column=1, sticky="e", pady=6)
    cb_mes.set(MESES[max(0, datetime.now().month - 1)])

    _label(grid, "Producción (Mbd)", muted=True).grid(row=4, column=0, sticky="w", pady=6)
    ent_prod = _entry(grid, width=300, placeholder="Ej. 1013.8")
    ent_prod.grid(row=4, column=1, sticky="e", pady=6)

    lbl_msg = ctk.CTkLabel(form, text="", font=("Roboto", 11), text_color=DANGER, anchor="w", justify="left", wraplength=520)
    lbl_msg.pack(fill="x", padx=16, pady=(0, 4))

    row_btns = ctk.CTkFrame(form, fg_color="transparent")
    row_btns.pack(fill="x", padx=16, pady=(0, 14))

    def on_tipo(_=None):
        es_mes = seg.get() == "Mes"
        cb_mes.configure(state="readonly" if es_mes else "disabled")
        lbl_mes.configure(text_color=MUTED if es_mes else "#4b5160")
    seg.configure(command=on_tipo)

    def set_msg(text, ok=False):
        lbl_msg.configure(text=text, text_color=OK if ok else DANGER)

    # ── Lista de registros guardados ──
    lst_card = _card(dialog, padx=22, pady=(0, 10))
    lst_title = _label(lst_card, "", bold=True)
    lst_title.pack(anchor="w", padx=16, pady=(12, 4))
    lst = ctk.CTkScrollableFrame(lst_card, fg_color="transparent", height=170)
    lst.pack(fill="x", padx=8, pady=(0, 10))

    def refresh_list():
        for w in lst.winfo_children():
            w.destroy()
        proc = cb_proc.get()
        rows = db_helper.get_extra_prod(proc)
        lst_title.configure(text=f"Registros guardados · {proc}")
        if not rows:
            _label(lst, "Sin registros extra para este proceso.", muted=True).pack(anchor="w", padx=8, pady=6)
            return

        def sort_key(r):
            anio, mes, _ = r
            idx = MESES.index(mes) if mes in MESES else -1
            return (str(anio), idx)

        for anio, mes, prod in sorted(rows, key=sort_key):
            fila = ctk.CTkFrame(lst, fg_color=NEUTRAL, corner_radius=8)
            fila.pack(fill="x", padx=4, pady=3)
            etiqueta = f"{anio} · {'Año completo' if mes == 'AÑO' else mes}"
            _label(fila, etiqueta, bold=True).pack(side="left", padx=12, pady=8)
            _label(fila, f"{float(prod):,.1f} Mbd", muted=True).pack(side="left", padx=4)
            _button(fila, "Eliminar", lambda a=anio, m=mes: delete_row(a, m), "danger", width=80, height=26).pack(side="right", padx=(4, 8), pady=6)
            _button(fila, "Editar", lambda a=anio, m=mes, p=prod: edit_row(a, m, p), "ghost", width=70, height=26).pack(side="right", padx=4, pady=6)

    def edit_row(anio, mes, prod):
        seg.set("Año completo" if mes == "AÑO" else "Mes")
        on_tipo()
        ent_anio.delete(0, "end"); ent_anio.insert(0, str(anio))
        if mes != "AÑO":
            cb_mes.set(mes)
        ent_prod.delete(0, "end"); ent_prod.insert(0, f"{float(prod):g}")
        set_msg("Editando registro existente: modifica el valor y pulsa Guardar.", ok=True)

    def delete_row(anio, mes):
        etiqueta = "el año completo" if mes == "AÑO" else f"el mes {mes}"
        if messagebox.askyesno("Eliminar registro", f"¿Eliminar {etiqueta} de {anio} para '{cb_proc.get()}'?", parent=dialog):
            db_helper.delete_extra_prod(cb_proc.get(), anio, mes)
            state["dirty"] = True
            refresh_list()
            set_msg("Registro eliminado.", ok=True)

    cb_proc.configure(command=lambda _=None: (refresh_list(), set_msg("")))

    # ── Guardar ──
    def on_save(_event=None):
        anio = ent_anio.get().strip()
        if not (anio.isdigit() and len(anio) == 4):
            set_msg("El año debe tener 4 dígitos (ej. 2025).")
            return
        if int(anio) > sheet_year:
            set_msg(f"El año no puede ser posterior a {sheet_year} (el año de la hoja de Excel).")
            return
        if int(anio) < 2000:
            set_msg("El año debe ser 2000 o posterior.")
            return
        try:
            prod = float(ent_prod.get().strip().replace(",", ""))
            if prod < 0:
                raise ValueError
        except ValueError:
            set_msg("La producción debe ser un número mayor o igual a 0.")
            return

        mes = cb_mes.get() if seg.get() == "Mes" else "AÑO"
        db_helper.save_extra_prod(cb_proc.get(), anio, mes, prod)
        app.cambios_sin_guardar = True
        state["dirty"] = True
        ent_prod.delete(0, "end")
        refresh_list()
        set_msg(f"Guardado: {anio} · {'año completo' if mes == 'AÑO' else mes} = {prod:g} Mbd.", ok=True)

    _button(row_btns, "Guardar", on_save, "primary", width=130).pack(side="right")
    _button(row_btns, "Limpiar campos", lambda: (ent_prod.delete(0, "end"), set_msg("")), "ghost", width=130).pack(side="right", padx=8)
    dialog.bind("<Return>", on_save)

    # ── Pie ──
    foot = ctk.CTkFrame(dialog, fg_color="transparent")
    foot.pack(fill="x", padx=22, pady=(4, 16), side="bottom")

    def on_clear_all():
        if messagebox.askyesno("Confirmar",
                               "¿Borrar TODOS los años y meses extra de TODOS los procesos?\n"
                               "(No afecta al Excel original.)", parent=dialog):
            db_helper.clear_db()
            app.cambios_sin_guardar = True
            state["dirty"] = True
            refresh_list()
            set_msg("Base de datos local limpiada.", ok=True)

    def on_close():
        dirty = state["dirty"]
        dialog.destroy()
        if dirty:
            _reload(app, "Recargando datos del Excel...")

    _button(foot, "Limpiar toda la base…", on_clear_all, "danger", width=170).pack(side="left")
    _button(foot, "Cerrar", on_close, "ghost", width=110).pack(side="right")
    dialog.protocol("WM_DELETE_WINDOW", on_close)

    on_tipo()
    refresh_list()
    ent_prod.focus_set()


# ════════════════════════════════════════════════════════════════════════
#  COORDENADAS EXCEL
# ════════════════════════════════════════════════════════════════════════
_RE_ROWS = re.compile(r"^\d+-\d+$")
_RE_COL = r"[A-Za-z]{1,3}"
_RE_COLS = re.compile(rf"^({_RE_COL}-{_RE_COL}|{_RE_COL}(,{_RE_COL})*)$")


def _col_to_num(col):
    n = 0
    for c in col.upper():
        n = n * 26 + (ord(c) - 64)
    return n - 1


def _num_to_col(n):
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def validate_rows(text):
    """Devuelve (inicio0, fin) o None si el texto vacío; lanza ValueError si es inválido."""
    text = text.replace(" ", "")
    if not text:
        return None
    if not _RE_ROWS.match(text):
        raise ValueError("Filas: usa el formato Inicio-Fin (ej. 21-51).")
    a, b = (int(x) for x in text.split("-"))
    if a < 1 or b < a:
        raise ValueError("Filas: el inicio debe ser ≥ 1 y no mayor que el fin.")
    return a - 1, b


def validate_cols(text):
    """Devuelve la lista de índices de columna o None si vacío; lanza ValueError si es inválido."""
    text = text.replace(" ", "")
    if not text:
        return None
    if not _RE_COLS.match(text):
        raise ValueError("Columnas: usa letras como A-H, o separadas por coma (L,M).")
    if "-" in text:
        a, b = (_col_to_num(x) for x in text.split("-"))
        if b < a:
            raise ValueError("Columnas: el rango está al revés (ej. AX-AW).")
        return list(range(a, b + 1))
    return [_col_to_num(x) for x in text.split(",")]


def open_config_coords_dialog(app):
    dialog = _make_dialog(app, "Coordenadas Excel", 820, 760)
    _header(dialog, "Coordenadas de Excel",
            "Indica de qué columnas y filas se leen los datos de cada proceso. Déjalo vacío para usar las coordenadas por defecto.")

    # ── Proceso + estado ──
    top = _card(dialog, padx=22, pady=(0, 10))
    top_in = ctk.CTkFrame(top, fg_color="transparent")
    top_in.pack(fill="x", padx=16, pady=12)
    _label(top_in, "Proceso", muted=True).pack(side="left", padx=(0, 10))
    specific = [v for v in app.cb_proceso.cget("values") if " -" in v]
    cb_proc = _combo(top_in, specific, width=280)
    cb_proc.pack(side="left")
    current = app.cb_proceso.get()
    cb_proc.set(current if current in specific else specific[0])
    lbl_state = ctk.CTkLabel(top_in, text="", font=("Roboto", 11, "bold"))
    lbl_state.pack(side="right")

    # ── Pestañas ──
    tabs_card = _card(dialog, padx=22, pady=(0, 10))
    seg = ctk.CTkSegmentedButton(tabs_card, values=["Diaria", "Programa", "Histórica"], height=32,
                                 fg_color=NEUTRAL, selected_color=ACCENT, selected_hover_color=ACCENT_H,
                                 unselected_color=NEUTRAL, unselected_hover_color=NEUTRAL_H,
                                 font=("Roboto", 12, "bold"))
    seg.pack(fill="x", padx=16, pady=(14, 8))
    seg.set("Diaria")

    info = _label(tabs_card, "", muted=True)
    info.pack(anchor="w", padx=16)

    fields = ctk.CTkFrame(tabs_card, fg_color="transparent")
    fields.pack(fill="x", padx=16, pady=(8, 4))
    fields.grid_columnconfigure((1, 3), weight=1)

    prefixes = {"Diaria": "d", "Programa": "p", "Histórica": "h"}
    hints = {
        "Diaria": ("Producción diaria del mes (una fila por día).", "Ej. L,M  ó  A-H", "Ej. 21-51"),
        "Programa": ("Metas de planeación (CMP / PODIM).", "Ej. BI  ó  AE-AF", "Ej. 74-104"),
        "Histórica": ("Años y meses de la tabla histórica (Año/Mes, Producción).", "Ej. AW-AX", "Ej. 21-40"),
    }
    entries = {}
    for pre in prefixes.values():
        entries[pre + "_cols"] = _entry(fields, width=170)
        entries[pre + "_filas"] = _entry(fields, width=170)
    _label(fields, "Columnas", muted=True).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
    _label(fields, "Filas", muted=True).grid(row=0, column=2, sticky="w", padx=(16, 8), pady=4)

    lbl_err = ctk.CTkLabel(tabs_card, text="", font=("Roboto", 11), text_color=DANGER, anchor="w")
    lbl_err.pack(fill="x", padx=16, pady=(0, 12))

    # ── Vista previa ──
    prev_card = _card(dialog, padx=22, pady=(0, 10))
    prev_top = ctk.CTkFrame(prev_card, fg_color="transparent")
    prev_top.pack(fill="x", padx=16, pady=(12, 4))
    _label(prev_top, "Vista previa (se actualiza al escribir)", bold=True).pack(side="left")
    lbl_prev_src = _label(prev_top, "", muted=True, size=11)
    lbl_prev_src.pack(side="right")
    preview = ctk.CTkTextbox(prev_card, height=190, fg_color=NEUTRAL, text_color=TEXT, corner_radius=8,
                             font=("Menlo", 11), wrap="none", border_width=1, border_color=BORDER)
    preview.pack(fill="x", padx=16, pady=(0, 14))

    # ── Lógica ──
    after_id = {"id": None}

    def active_prefix():
        return prefixes[seg.get()]

    def show_tab(_=None):
        pre = active_prefix()
        for p in prefixes.values():
            for kind in ("cols", "filas"):
                entries[f"{p}_{kind}"].grid_forget()
        info.configure(text=hints[seg.get()][0])
        entries[pre + "_cols"].configure(placeholder_text=hints[seg.get()][1])
        entries[pre + "_filas"].configure(placeholder_text=hints[seg.get()][2])
        entries[pre + "_cols"].grid(row=0, column=1, sticky="ew", pady=4)
        entries[pre + "_filas"].grid(row=0, column=3, sticky="ew", pady=4)
        validate_and_preview()
    seg.configure(command=show_tab)

    def parse_pair(pre):
        """Valida columnas y filas de una pestaña. Devuelve (cols, rows) o lanza ValueError."""
        for kind, fn in (("cols", validate_cols), ("filas", validate_rows)):
            e = entries[f"{pre}_{kind}"]
            try:
                fn(e.get())
                e.configure(border_color=BORDER)
            except ValueError:
                e.configure(border_color=DANGER)
        return (validate_cols(entries[pre + "_cols"].get()),
                validate_rows(entries[pre + "_filas"].get()))

    def render_preview(cols, rows):
        df = getattr(app, "cached_df_sheet", None)
        if df is None:
            lbl_prev_src.configure(text="")
            return "Carga un archivo de Excel para ver la vista previa."
        lbl_prev_src.configure(text="Hoja en memoria")
        r0, r1 = rows
        r1 = min(r1, r0 + 8, len(df))
        cols = [c for c in cols if c < df.shape[1]]
        if not cols or r0 >= len(df):
            return "El rango queda fuera de la hoja."
        w = 13
        head = "fila".rjust(5) + "".join(_num_to_col(c).center(w) for c in cols)
        lines = [head, "─" * len(head)]
        for r in range(r0, r1):
            cells = []
            for c in cols:
                v = df.iat[r, c]
                if v != v or v is None:
                    s = "·"
                elif isinstance(v, float):
                    s = f"{v:.2f}".rstrip("0").rstrip(".")
                else:
                    s = str(v)
                cells.append(s[: w - 1].center(w))
            lines.append(str(r + 1).rjust(5) + "".join(cells))
        total = rows[1] - rows[0]
        if total > 9:
            lines.append(f"… ({total} filas en total)")
        return "\n".join(lines)

    def validate_and_preview():
        pre = active_prefix()
        try:
            cols, rows = parse_pair(pre)
            lbl_err.configure(text="")
        except ValueError as e:
            lbl_err.configure(text=str(e))
            set_preview("Corrige las coordenadas para ver la vista previa.")
            return
        if not cols or not rows:
            lbl_err.configure(text="")
            set_preview("Sin coordenadas personalizadas en esta pestaña: se usan las del programa.\n"
                        "Escribe columnas y filas para previsualizarlas aquí.")
            return
        set_preview(render_preview(cols, rows))

    def set_preview(text):
        preview.configure(state="normal")
        preview.delete("1.0", "end")
        preview.insert("1.0", text)
        preview.configure(state="disabled")

    def schedule(_=None):
        if after_id["id"]:
            dialog.after_cancel(after_id["id"])
        after_id["id"] = dialog.after(350, validate_and_preview)

    for e in entries.values():
        e.bind("<KeyRelease>", schedule)

    def load_proc(_=None):
        for e in entries.values():
            e.delete(0, "end")
        coords = db_helper.get_coordenadas_override(cb_proc.get())
        if coords:
            for pre, name in (("d", "diaria"), ("p", "programa"), ("h", "historica")):
                entries[pre + "_filas"].insert(0, coords.get(name + "_filas") or "")
                entries[pre + "_cols"].insert(0, coords.get(name + "_cols") or "")
            lbl_state.configure(text="● Personalizadas", text_color=ACCENT)
        else:
            lbl_state.configure(text="● Por defecto", text_color=OK)
        show_tab()
    cb_proc.configure(command=load_proc)

    def on_save():
        values = {}
        for name, pre in prefixes.items():
            try:
                parse_pair(pre)
            except ValueError as e:
                seg.set(name)
                show_tab()
                lbl_err.configure(text=f"{name}: {e}")
                return
            values[pre] = (entries[pre + "_filas"].get().strip(), entries[pre + "_cols"].get().strip())
        proceso = cb_proc.get()
        if not any(f or c for f, c in values.values()):
            db_helper.delete_coordenadas_override(proceso)
        else:
            db_helper.save_coordenadas_override(proceso, values["d"][0], values["d"][1],
                                                values["p"][0], values["p"][1], values["h"][0], values["h"][1])
        app.cambios_sin_guardar = True
        dialog.destroy()
        _reload(app, "Recargando datos del Excel con nuevas coordenadas...")

    def on_restore():
        proceso = cb_proc.get()
        if not messagebox.askyesno("Restaurar", f"¿Quitar la configuración personalizada de '{proceso}' "
                                   "y volver a las coordenadas por defecto?", parent=dialog):
            return
        db_helper.delete_coordenadas_override(proceso)
        app.cambios_sin_guardar = True
        dialog.destroy()
        _reload(app, "Recargando datos del Excel con coordenadas por defecto...")

    foot = ctk.CTkFrame(dialog, fg_color="transparent")
    foot.pack(fill="x", padx=22, pady=(0, 16), side="bottom")
    _button(foot, "Restaurar por defecto", on_restore, "danger", width=170).pack(side="left")
    _button(foot, "Guardar y recargar", on_save, "primary", width=170).pack(side="right")
    _button(foot, "Cancelar", dialog.destroy, "ghost", width=100).pack(side="right", padx=8)

    load_proc()
