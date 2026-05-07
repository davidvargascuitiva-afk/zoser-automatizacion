import threading
import os
from datetime import datetime
from typing import List

import customtkinter as ctk
from tkinter import filedialog, messagebox

from app.models.sensor import Sensor
from app.models.proyecto import ProyectoConfig
from app.core.parser_primarios import detectar_sensores, cargar_datos_primarios
from app.core.parser_fallas import cargar_datos_fallas
from app.core.data_cleaner import limpiar_datos, clip_calificacion
from app.core.data_simulator import simular_sensores_faltantes
from app.core.excel_writer import llenar_plantilla

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

SIDEBAR_W = 230
ACCENT    = "#1a5fa8"
GREEN     = "#1a7a1a"
GREEN_H   = "#135213"
AMBER     = "#f0a500"
MUTED     = "gray55"


class CalificadorApp(ctk.CTk):
    # ══════════════════════════════════════════════════════════════
    #  INIT
    # ══════════════════════════════════════════════════════════════
    def __init__(self):
        super().__init__()
        self.title("CalificadorIA — Cámaras de Estabilidad")
        self.geometry("1140x740")
        self.minsize(920, 640)

        self.sensores: List[Sensor] = []
        self.sensor_vars: list = []
        self._paso = 0
        self._nav_btns: list = []
        self._pages: list = []

        self._build_shell()
        self._page_archivos()
        self._page_sensores()
        self._page_config()
        self._page_equipo()
        self._page_generar()
        self._go(0)

    # ══════════════════════════════════════════════════════════════
    #  SHELL  (sidebar + header + content + navbar)
    # ══════════════════════════════════════════════════════════════
    def _build_shell(self):
        # ── sidebar ───────────────────────────────────────────────
        self.sidebar = ctk.CTkFrame(self, width=SIDEBAR_W, corner_radius=0,
                                    fg_color=("gray84", "gray12"))
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # logo
        logo = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo.pack(fill="x", padx=20, pady=(28, 10))
        ctk.CTkLabel(logo, text="CalificadorIA",
                     font=ctk.CTkFont(size=21, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(logo, text="Cámaras de Estabilidad",
                     font=ctk.CTkFont(size=11), text_color=MUTED).pack(anchor="w", pady=(2, 0))

        ctk.CTkFrame(self.sidebar, height=1, fg_color=("gray65", "gray28")).pack(
            fill="x", padx=18, pady=(10, 14))

        # step nav buttons
        steps = [
            ("1", "Archivos"),
            ("2", "Sensores"),
            ("3", "Configuración"),
            ("4", "Datos del equipo"),
            ("5", "Generar"),
        ]
        self._nav_btns = []
        for i, (num, name) in enumerate(steps):
            btn = ctk.CTkButton(
                self.sidebar,
                text=f"   {num}   {name}",
                anchor="w",
                height=42,
                corner_radius=8,
                fg_color="transparent",
                hover_color=("gray74", "gray22"),
                text_color=("gray25", "gray78"),
                font=ctk.CTkFont(size=13),
                command=lambda idx=i: self._go(idx),
            )
            btn.pack(fill="x", padx=10, pady=2)
            self._nav_btns.append(btn)

        ctk.CTkLabel(self.sidebar, text="v1.0  ·  CAMECAL",
                     font=ctk.CTkFont(size=10), text_color="gray40").pack(
            side="bottom", pady=14)

        # ── main area ─────────────────────────────────────────────
        self.main = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main.pack(side="left", fill="both", expand=True)

        # header strip
        self.header = ctk.CTkFrame(self.main, height=60, corner_radius=0,
                                   fg_color=("gray88", "gray15"))
        self.header.pack(fill="x")
        self.header.pack_propagate(False)
        self.lbl_title = ctk.CTkLabel(self.header, text="",
                                      font=ctk.CTkFont(size=18, weight="bold"))
        self.lbl_title.pack(side="left", padx=26, pady=14)
        self.lbl_sub = ctk.CTkLabel(self.header, text="",
                                    font=ctk.CTkFont(size=12), text_color=MUTED)
        self.lbl_sub.pack(side="left", padx=4)

        # bottom nav bar (pack BEFORE page_area so it anchors to bottom)
        nav_bar = ctk.CTkFrame(self.main, height=62, corner_radius=0,
                               fg_color=("gray88", "gray15"))
        nav_bar.pack(fill="x", side="bottom")
        nav_bar.pack_propagate(False)

        self.btn_prev = ctk.CTkButton(
            nav_bar, text="← Anterior", width=130, height=36,
            command=lambda: self._go(self._paso - 1),
        )
        self.btn_prev.pack(side="left", padx=22, pady=13)

        self.btn_next = ctk.CTkButton(
            nav_bar, text="Siguiente →", width=130, height=36,
            command=self._go_next,
        )
        self.btn_next.pack(side="right", padx=22, pady=13)

        # page area (takes remaining space)
        self.page_area = ctk.CTkFrame(self.main, fg_color="transparent", corner_radius=0)
        self.page_area.pack(fill="both", expand=True)

    # ══════════════════════════════════════════════════════════════
    #  NAVIGATION
    # ══════════════════════════════════════════════════════════════
    def _go(self, idx: int):
        n = len(self._pages)
        if idx < 0 or idx >= n:
            return
        self._paso = idx

        # swap visible page
        for i, page in enumerate(self._pages):
            if i == idx:
                page.pack(fill="both", expand=True)
            else:
                page.pack_forget()

        # header text
        info = [
            ("Archivos de entrada",    "Carga los tres archivos Excel requeridos"),
            ("Configuración de sensores", "Asigna la posición de cada sensor en la cámara"),
            ("Parámetros del ensayo",  "Fechas, rangos del equipo y setpoints"),
            ("Datos del equipo",       "Información del cliente y del equipo calificado"),
            ("Generar reporte",        "Revisa el resumen y genera el Excel final"),
        ]
        self.lbl_title.configure(text=info[idx][0])
        self.lbl_sub.configure(text=info[idx][1])

        # highlight active step in sidebar
        for i, btn in enumerate(self._nav_btns):
            if i == idx:
                btn.configure(fg_color=ACCENT, text_color="white",
                              font=ctk.CTkFont(size=13, weight="bold"),
                              hover_color=ACCENT)
            else:
                btn.configure(fg_color="transparent",
                              text_color=("gray25", "gray78"),
                              font=ctk.CTkFont(size=13, weight="normal"),
                              hover_color=("gray74", "gray22"))

        # prev/next buttons
        self.btn_prev.configure(state="normal" if idx > 0 else "disabled")
        if idx < n - 1:
            self.btn_next.configure(text="Siguiente →", state="normal")
        else:
            self.btn_next.configure(text="", state="disabled")

    def _go_next(self):
        if self._paso == 0:
            # Step 1→2: detect sensors first, then navigate
            self._detectar_sensores(on_success=lambda: self._go(1))
        else:
            self._go(self._paso + 1)

    # ══════════════════════════════════════════════════════════════
    #  PAGE HELPERS
    # ══════════════════════════════════════════════════════════════
    def _new_page(self) -> ctk.CTkScrollableFrame:
        page = ctk.CTkScrollableFrame(self.page_area, fg_color="transparent",
                                      corner_radius=0)
        self._pages.append(page)
        return page

    def _card(self, parent, title: str = "") -> ctk.CTkFrame:
        outer = ctk.CTkFrame(parent, fg_color=("gray91", "gray17"), corner_radius=12)
        outer.pack(fill="x", padx=24, pady=(0, 14))
        if title:
            ctk.CTkLabel(outer, text=title,
                         font=ctk.CTkFont(size=13, weight="bold")).pack(
                anchor="w", padx=18, pady=(14, 4))
            ctk.CTkFrame(outer, height=1, fg_color=("gray72", "gray30")).pack(
                fill="x", padx=18, pady=(0, 10))
        return outer

    def _row(self, parent) -> ctk.CTkFrame:
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", padx=18, pady=8)
        return f

    def _lbl(self, parent, text: str, width: int = 0, **kw):
        kw.setdefault("anchor", "w")
        l = ctk.CTkLabel(parent, text=text, **kw)
        if width:
            l.configure(width=width)
        l.pack(side="left")
        return l

    def _entry(self, parent, var, width: int = 200, **kw) -> ctk.CTkEntry:
        e = ctk.CTkEntry(parent, textvariable=var, width=width, height=34, **kw)
        e.pack(side="left", padx=(0, 8))
        return e

    def _unit(self, parent, text: str):
        ctk.CTkLabel(parent, text=text, text_color=MUTED,
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 28))

    def _file_row(self, card, icon: str, label: str, key: str, save: bool = False):
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=9)
        ctk.CTkLabel(row, text=f"{icon}  {label}", width=220,
                     anchor="w", font=ctk.CTkFont(size=12)).pack(side="left")
        var = ctk.StringVar()
        setattr(self, f"var_{key}", var)
        ctk.CTkEntry(row, textvariable=var,
                     placeholder_text="Ningún archivo seleccionado",
                     height=34).pack(side="left", fill="x", expand=True, padx=(0, 10))
        cmd = (lambda k=key: self._browse_save(k)) if save else (lambda k=key: self._browse(k))
        ctk.CTkButton(row, text="Examinar", width=100, height=34, command=cmd).pack(side="left")

    # ══════════════════════════════════════════════════════════════
    #  PAGE 1 — ARCHIVOS
    # ══════════════════════════════════════════════════════════════
    def _page_archivos(self):
        page = self._new_page()
        ctk.CTkFrame(page, height=10, fg_color="transparent").pack()

        card = self._card(page, "Archivos de entrada")
        self._file_row(card, "📊", "Excel Primarios (.xlsx)", "primarios")
        self._file_row(card, "⚡", "Excel Fallas (.xlsx)", "fallas")
        self._file_row(card, "📋", "Plantilla (.xlsx)", "plantilla")
        ctk.CTkFrame(card, height=6, fg_color="transparent").pack()

        card2 = self._card(page, "Carpeta de salida")
        row_sal = ctk.CTkFrame(card2, fg_color="transparent")
        row_sal.pack(fill="x", padx=18, pady=9)
        ctk.CTkLabel(row_sal, text="💾  Carpeta donde guardar el reporte:", width=280,
                     anchor="w", font=ctk.CTkFont(size=12)).pack(side="left")
        self.var_carpeta_salida = ctk.StringVar()
        ctk.CTkEntry(row_sal, textvariable=self.var_carpeta_salida,
                     placeholder_text="Selecciona una carpeta...",
                     height=34).pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row_sal, text="Examinar", width=100, height=34,
                      command=self._browse_carpeta).pack(side="left")
        ctk.CTkLabel(card2,
                     text="El archivo se guardará con el nombre del N° de Ensayo (ej: B6367-110226.xlsx).",
                     text_color=MUTED, font=ctk.CTkFont(size=11)).pack(
            anchor="w", padx=18, pady=(0, 12))

        card3 = self._card(page, "Sensores instalados en la cámara")
        row = self._row(card3)
        self._lbl(row, "Número total de sensores:", width=270)
        self.var_num_sensores = ctk.StringVar(value="9")
        self._entry(row, self.var_num_sensores, width=72, justify="center")
        ctk.CTkLabel(card3,
                     text="Si el archivo tiene menos sensores que este número, "
                          "los faltantes se simularán automáticamente.",
                     text_color=MUTED, font=ctk.CTkFont(size=11),
                     wraplength=640, justify="left").pack(anchor="w", padx=18, pady=(0, 14))

    # ══════════════════════════════════════════════════════════════
    #  PAGE 2 — SENSORES
    # ══════════════════════════════════════════════════════════════
    def _page_sensores(self):
        page = self._new_page()
        ctk.CTkFrame(page, height=10, fg_color="transparent").pack()

        card = self._card(page, "Sensores detectados")
        ctk.CTkLabel(card,
                     text="Asigna la posición (1–9) de cada sensor dentro de la cámara. "
                          "Desmarca los que no deban incluirse en el reporte.",
                     text_color=MUTED, font=ctk.CTkFont(size=11),
                     wraplength=740, justify="left").pack(anchor="w", padx=18, pady=(0, 10))

        self.sensores_frame = ctk.CTkFrame(card, fg_color="transparent")
        self.sensores_frame.pack(fill="x", padx=10, pady=(0, 4))

        ctk.CTkLabel(card,
                     text="Los sensores en color dorado [SIM] serán generados sintéticamente "
                          "siguiendo la tendencia de los sensores reales.",
                     text_color=AMBER, font=ctk.CTkFont(size=11)).pack(
            anchor="w", padx=18, pady=(4, 14))

    def _render_sensores(self):
        for w in self.sensores_frame.winfo_children():
            w.destroy()
        self.sensor_vars = []

        # table header
        hdr = ctk.CTkFrame(self.sensores_frame, fg_color=("gray78", "gray23"), corner_radius=6)
        hdr.pack(fill="x", padx=6, pady=(4, 2))
        for txt, w in [("Serial / ID", 195), ("Modelo", 140),
                       ("Temp", 58), ("HR", 58), ("Posición", 90), ("Usar", 70)]:
            ctk.CTkLabel(hdr, text=txt, width=w,
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=("gray20", "gray75")).pack(side="left", padx=6, pady=7)

        for i, sensor in enumerate(self.sensores):
            es_sim = sensor.col_idx_temp is None and sensor.col_idx_hum is None
            bg = ("gray87", "gray19") if i % 2 == 0 else ("gray83", "gray21")
            row = ctk.CTkFrame(self.sensores_frame, fg_color=bg, corner_radius=4)
            row.pack(fill="x", padx=6, pady=1)

            serial_txt = f"[SIM] {sensor.serial}" if es_sim else sensor.serial
            ctk.CTkLabel(row, text=serial_txt, width=195, anchor="w",
                         text_color=AMBER if es_sim else None,
                         font=ctk.CTkFont(size=12)).pack(side="left", padx=8, pady=9)
            ctk.CTkLabel(row, text=sensor.nombre, width=140, anchor="w",
                         font=ctk.CTkFont(size=12)).pack(side="left", padx=6)
            ctk.CTkLabel(row, text="✓" if sensor.tiene_temperatura else "—",
                         width=58).pack(side="left", padx=6)
            ctk.CTkLabel(row, text="✓" if sensor.tiene_humedad else "—",
                         width=58).pack(side="left", padx=6)

            pos_var = ctk.StringVar(value=str(sensor.posicion))
            ctk.CTkEntry(row, textvariable=pos_var, width=64,
                         justify="center", height=30).pack(side="left", padx=8)

            usar_var = ctk.BooleanVar(value=True)
            ctk.CTkCheckBox(row, text="", variable=usar_var, width=60).pack(side="left", padx=10)
            self.sensor_vars.append({"sensor": sensor, "pos": pos_var, "usar": usar_var})

    # ══════════════════════════════════════════════════════════════
    #  PAGE 3 — CONFIGURACIÓN
    # ══════════════════════════════════════════════════════════════
    def _page_config(self):
        page = self._new_page()
        ctk.CTkFrame(page, height=10, fg_color="transparent").pack()

        # inicio 24h
        card1 = self._card(page, "Inicio del periodo de 24 horas")
        row1 = self._row(card1)
        self._lbl(row1, "Fecha:", width=110)
        self.var_ini_fecha = ctk.StringVar()
        self._entry(row1, self.var_ini_fecha, width=155, placeholder_text="DD/MM/AAAA")
        self._lbl(row1, "Hora:", width=70)
        self.var_ini_hora = ctk.StringVar()
        self._entry(row1, self.var_ini_hora, width=110, placeholder_text="HH:MM")
        ctk.CTkFrame(card1, height=6, fg_color="transparent").pack()

        # prueba de fallas
        card2 = self._card(page, "Prueba de fallas")
        row2 = self._row(card2)
        self._lbl(row2, "Tipo de prueba:", width=150)
        self.var_tipo = ctk.StringVar(value="PO")
        ctk.CTkSegmentedButton(row2, values=["PO", "PP"],
                               variable=self.var_tipo, width=150).pack(side="left", padx=(0, 24))
        ctk.CTkLabel(card2,
                     text="PO: 30 min corte energía + 30 min recuperación  (60 registros)\n"
                          "PP: 5 min puerta abierta  + 30 min recuperación  (35 registros)",
                     text_color=MUTED, font=ctk.CTkFont(size=11)).pack(
            anchor="w", padx=18, pady=(0, 6))
        row3 = self._row(card2)
        self._lbl(row3, "Fecha inicio falla:", width=150)
        self.var_falla_fecha = ctk.StringVar()
        self._entry(row3, self.var_falla_fecha, width=155, placeholder_text="DD/MM/AAAA")
        self._lbl(row3, "Hora:", width=70)
        self.var_falla_hora = ctk.StringVar()
        self._entry(row3, self.var_falla_hora, width=110, placeholder_text="HH:MM")
        ctk.CTkFrame(card2, height=6, fg_color="transparent").pack()

        # rango operativo
        card3 = self._card(page, "Rango operativo del equipo")
        r1 = self._row(card3)
        self._lbl(r1, "Temp mínima:", width=140)
        self.var_tmin = ctk.StringVar(value="28")
        self._entry(r1, self.var_tmin, width=80, justify="center")
        self._unit(r1, "°C")
        self._lbl(r1, "Temp máxima:", width=140)
        self.var_tmax = ctk.StringVar(value="32")
        self._entry(r1, self.var_tmax, width=80, justify="center")
        self._unit(r1, "°C")

        r2 = self._row(card3)
        self._lbl(r2, "HR mínima:", width=140)
        self.var_hmin = ctk.StringVar(value="70")
        self._entry(r2, self.var_hmin, width=80, justify="center")
        self._unit(r2, "%")
        self._lbl(r2, "HR máxima:", width=140)
        self.var_hmax = ctk.StringVar(value="80")
        self._entry(r2, self.var_hmax, width=80, justify="center")
        self._unit(r2, "%")
        ctk.CTkFrame(card3, height=6, fg_color="transparent").pack()

        # setpoints
        card4 = self._card(page, "Setpoints de trabajo")
        r3 = self._row(card4)
        self._lbl(r3, "Temperatura:", width=140)
        self.var_sp_temp = ctk.StringVar(value="30")
        self._entry(r3, self.var_sp_temp, width=80, justify="center")
        self._unit(r3, "°C")
        self._lbl(r3, "Humedad relativa:", width=160)
        self.var_sp_hum = ctk.StringVar(value="75")
        self._entry(r3, self.var_sp_hum, width=80, justify="center")
        self._unit(r3, "%")
        ctk.CTkFrame(card4, height=6, fg_color="transparent").pack()
        ctk.CTkFrame(page, height=10, fg_color="transparent").pack()

    # ══════════════════════════════════════════════════════════════
    #  PAGE 4 — DATOS DEL EQUIPO
    # ══════════════════════════════════════════════════════════════
    def _page_equipo(self):
        page = self._new_page()
        ctk.CTkFrame(page, height=10, fg_color="transparent").pack()

        card = self._card(page, "Identificación del cliente y del equipo")
        self.equipo_vars: dict = {}
        fields = [
            ("Empresa / Cliente:",  "empresa",   "Nombre de la empresa"),
            ("Marca del equipo:",   "marca",     "Ej: BINDER, Memmert, Thermo"),
            ("Ubicación:",          "ubicacion", "Sala, edificio o área donde está el equipo"),
            ("Código del equipo:",  "codigo",    "Código interno de inventario"),
            ("N° de Ensayo:",       "ensayo",    "Ej: B6367-110226"),
        ]
        for lbl, key, placeholder in fields:
            row = self._row(card)
            self._lbl(row, lbl, width=200)
            var = ctk.StringVar()
            self.equipo_vars[key] = var
            ctk.CTkEntry(row, textvariable=var, placeholder_text=placeholder,
                         height=34).pack(side="left", fill="x", expand=True, padx=(0, 18))
        ctk.CTkFrame(card, height=6, fg_color="transparent").pack()

        card2 = self._card(page, "Lectura del controlador del equipo (Tabla 5)")
        ctk.CTkLabel(card2,
                     text="Temperatura y HR que muestra el controlador del equipo al momento del ensayo. "
                          "Si se omite, se usa el setpoint como valor de referencia.",
                     text_color=MUTED, font=ctk.CTkFont(size=11),
                     wraplength=660, justify="left").pack(anchor="w", padx=18, pady=(0, 8))
        row_lec = self._row(card2)
        self._lbl(row_lec, "Temp del equipo:", width=180)
        self.equipo_vars['lec_temp'] = ctk.StringVar()
        self._entry(row_lec, self.equipo_vars['lec_temp'],
                    width=100, justify="center", placeholder_text="30.0")
        self._unit(row_lec, "°C")
        self._lbl(row_lec, "HR del equipo:", width=140)
        self.equipo_vars['lec_hum'] = ctk.StringVar()
        self._entry(row_lec, self.equipo_vars['lec_hum'],
                    width=100, justify="center", placeholder_text="75.0")
        self._unit(row_lec, "%")
        ctk.CTkFrame(card2, height=6, fg_color="transparent").pack()
        ctk.CTkFrame(page, height=10, fg_color="transparent").pack()

    # ══════════════════════════════════════════════════════════════
    #  PAGE 5 — GENERAR
    # ══════════════════════════════════════════════════════════════
    def _page_generar(self):
        page = self._new_page()
        ctk.CTkFrame(page, height=10, fg_color="transparent").pack()

        card = self._card(page, "Registro de procesamiento")
        self.log = ctk.CTkTextbox(card, height=390,
                                  font=ctk.CTkFont(family="Courier", size=11))
        self.log.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        self.btn_gen = ctk.CTkButton(
            page,
            text="  GENERAR REPORTE EXCEL",
            command=self._iniciar_generacion,
            font=ctk.CTkFont(size=15, weight="bold"),
            height=54,
            corner_radius=10,
            fg_color=GREEN,
            hover_color=GREEN_H,
        )
        self.btn_gen.pack(fill="x", padx=24, pady=(4, 6))

        ctk.CTkLabel(page,
                     text="El proceso puede tardar entre 10 y 40 segundos según el tamaño de los archivos.",
                     text_color=MUTED, font=ctk.CTkFont(size=11)).pack(pady=(0, 12))

    # ══════════════════════════════════════════════════════════════
    #  FILE BROWSER HELPERS
    # ══════════════════════════════════════════════════════════════
    def _parsear_fecha_hora(self, fecha: str, hora: str, msg_error: str) -> datetime:
        fecha = fecha.strip()
        hora = hora.strip().lower().replace(".", "").replace(" ", "")
        # normalizar separadores de fecha
        fecha = fecha.replace("-", "/")
        formatos = [
            "%d/%m/%Y %H:%M",
            "%d/%m/%Y %I:%M%p",
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y %I:%M:%S%p",
        ]
        texto = f"{fecha} {hora}"
        for fmt in formatos:
            try:
                return datetime.strptime(texto, fmt)
            except ValueError:
                continue
        raise ValueError(f"{msg_error}. Escribe la hora como HH:MM (ej: 11:00 o 14:30)")

    def _browse(self, key: str):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            getattr(self, f"var_{key}").set(path)

    def _browse_carpeta(self):
        carpeta = filedialog.askdirectory(title="Selecciona la carpeta de destino")
        if carpeta:
            self.var_carpeta_salida.set(carpeta)

    # ══════════════════════════════════════════════════════════════
    #  SENSOR DETECTION
    # ══════════════════════════════════════════════════════════════
    def _detectar_sensores(self, on_success=None):
        ruta = self.var_primarios.get()
        if not ruta:
            messagebox.showerror("Error", "Selecciona el archivo de Primarios primero.")
            return
        try:
            num = int(self.var_num_sensores.get())
        except ValueError:
            messagebox.showerror("Error", "El número de sensores debe ser un entero.")
            return
        try:
            detectados = detectar_sensores(ruta)
            self.sensores = self._completar_sensores(detectados, num)
            self._render_sensores()
            simulados = sum(1 for s in self.sensores
                            if s.col_idx_temp is None and s.col_idx_hum is None)
            msg = f"Sensores encontrados en el archivo: {len(detectados)}\nSensores configurados: {num}"
            if simulados:
                msg += f"\nSensores a simular: {simulados}  [marcados en dorado]"
            messagebox.showinfo("Sensores detectados", msg)
            if on_success:
                on_success()
        except Exception as e:
            messagebox.showerror("Error al leer archivo", str(e))

    def _completar_sensores(self, detectados: List[Sensor], num_total: int) -> List[Sensor]:
        posiciones_usadas = {s.posicion for s in detectados}
        resultado = list(detectados)
        pos_sig = max(posiciones_usadas, default=0) + 1
        while len(resultado) < num_total:
            while pos_sig in posiciones_usadas:
                pos_sig += 1
            resultado.append(Sensor(
                serial=f"SIM-{pos_sig:02d}",
                nombre="RHTemp101A",
                descripcion="Sensor Simulado",
                posicion=pos_sig,
                tiene_temperatura=True,
                tiene_humedad=True,
                col_idx_temp=None,
                col_idx_hum=None,
            ))
            posiciones_usadas.add(pos_sig)
            pos_sig += 1
        return resultado

    # ══════════════════════════════════════════════════════════════
    #  GENERATION PIPELINE
    # ══════════════════════════════════════════════════════════════
    def _log(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.update_idletasks()

    def _iniciar_generacion(self):
        self.btn_gen.configure(state="disabled")
        self.log.delete("1.0", "end")
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        try:
            self._log("Validando configuración...")
            config = self._build_config()

            self._log("Aplicando configuración de sensores...")
            sensores = self._sensores_activos()

            self._log(f"Leyendo primarios: {os.path.basename(config.ruta_primarios)}")
            sensores = cargar_datos_primarios(
                config.ruta_primarios, sensores, config.inicio_24h)
            self._log(f"  {len(sensores)} sensores cargados")

            self._log("Limpiando valores fuera de rango...")
            sensores = limpiar_datos(
                sensores,
                config.rango_temp_min, config.rango_temp_max,
                config.rango_hum_min,  config.rango_hum_max,
            )

            timestamps = None
            for s in sensores:
                if s.datos is not None and len(s.datos) > 0:
                    timestamps = s.datos['timestamp']
                    break
            if timestamps is None:
                raise ValueError("No se encontraron datos en el rango de 24 horas indicado.")

            for s in sensores:
                if s.tiene_temperatura and not s.tiene_datos_temp():
                    self._log(f"  ⚠ {s.serial}: sin datos de temperatura — se simulará")
                if s.tiene_humedad and not s.tiene_datos_hum():
                    self._log(f"  ⚠ {s.serial}: sin datos de HR — se simulará")

            sensores = simular_sensores_faltantes(
                sensores, timestamps,
                config.rango_temp_min, config.rango_temp_max,
                config.rango_hum_min,  config.rango_hum_max,
                num_sensores=config.num_sensores,
                setpoint_temp=config.setpoint_temp,
                setpoint_hum=config.setpoint_hum,
            )

            self._log("Aplicando tolerancia de calificación (±2 °C / ±5 %HR)...")
            sensores = clip_calificacion(sensores, config.setpoint_temp, config.setpoint_hum)

            self._log(f"Leyendo fallas ({config.tipo_prueba}): "
                      f"{os.path.basename(config.ruta_fallas)}")
            df_ft, df_fh = cargar_datos_fallas(
                config.ruta_fallas, config.tipo_prueba, config.inicio_falla)
            self._log(f"  {len(df_ft)} registros temp  |  {len(df_fh)} registros HR")

            self._log("Generando Excel final...")
            llenar_plantilla(
                config.ruta_plantilla, config.ruta_salida,
                sensores, timestamps, df_ft, df_fh, config,
            )

            self._log(f"\n✓  REPORTE GENERADO EXITOSAMENTE")
            self._log(f"   {config.ruta_salida}")

        except Exception as exc:
            import traceback
            self._log(f"\n✗  ERROR: {exc}")
            self._log(traceback.format_exc())
        finally:
            self.btn_gen.configure(state="normal")

    # ══════════════════════════════════════════════════════════════
    #  CONFIG BUILDER  &  SENSOR FILTER
    # ══════════════════════════════════════════════════════════════
    def _build_config(self) -> ProyectoConfig:
        c = ProyectoConfig()
        c.ruta_primarios = self.var_primarios.get()
        c.ruta_fallas    = self.var_fallas.get()
        c.ruta_plantilla = self.var_plantilla.get()

        for ruta, nombre in [(c.ruta_primarios, "Primarios"),
                              (c.ruta_fallas,   "Fallas"),
                              (c.ruta_plantilla, "Plantilla")]:
            if not ruta:
                raise ValueError(f"Selecciona el archivo de {nombre}.")

        carpeta = self.var_carpeta_salida.get().strip()
        if not carpeta:
            raise ValueError("Selecciona la carpeta donde guardar el reporte.")

        # El nombre se construye más adelante, cuando ya tengamos el N° de Ensayo.
        # Se asigna en la segunda mitad de este método.

        c.inicio_24h = self._parsear_fecha_hora(
            self.var_ini_fecha.get(), self.var_ini_hora.get(),
            "Fecha/hora de inicio 24h inválida")

        c.inicio_falla = self._parsear_fecha_hora(
            self.var_falla_fecha.get(), self.var_falla_hora.get(),
            "Fecha/hora de inicio de falla inválida")

        c.tipo_prueba = self.var_tipo.get()

        try:
            c.rango_temp_min = float(self.var_tmin.get())
            c.rango_temp_max = float(self.var_tmax.get())
            c.rango_hum_min  = float(self.var_hmin.get())
            c.rango_hum_max  = float(self.var_hmax.get())
            c.setpoint_temp  = float(self.var_sp_temp.get())
            c.setpoint_hum   = float(self.var_sp_hum.get())
        except ValueError:
            raise ValueError("Los rangos y setpoints deben ser valores numéricos.")

        try:
            c.num_sensores = int(self.var_num_sensores.get())
        except ValueError:
            pass

        c.empresa       = self.equipo_vars['empresa'].get()
        c.marca_equipo  = self.equipo_vars['marca'].get()
        c.ubicacion     = self.equipo_vars['ubicacion'].get()
        c.codigo_equipo = self.equipo_vars['codigo'].get()
        c.ensayo        = self.equipo_vars['ensayo'].get()

        # Construir ruta de salida: carpeta + nombre basado en ensayo
        nombre_base = c.ensayo.strip() if c.ensayo.strip() else "reporte"
        nombre_base = nombre_base.replace("/", "-").replace("\\", "-")
        c.ruta_salida = os.path.join(carpeta, f"{nombre_base}.xlsx")

        lec_t = self.equipo_vars['lec_temp'].get().strip()
        lec_h = self.equipo_vars['lec_hum'].get().strip()
        try:
            c.lectura_equipo_temp = float(lec_t) if lec_t else c.setpoint_temp
            c.lectura_equipo_hum  = float(lec_h) if lec_h else c.setpoint_hum
        except ValueError:
            raise ValueError("La lectura del equipo debe ser un número decimal (ej: 30.2).")
        return c

    def _sensores_activos(self) -> List[Sensor]:
        result = []
        for sv in self.sensor_vars:
            if not sv['usar'].get():
                continue
            sensor = sv['sensor']
            try:
                sensor.posicion = int(sv['pos'].get())
            except ValueError:
                raise ValueError(f"Posición inválida para el sensor {sensor.serial}")
            result.append(sensor)
        return result
