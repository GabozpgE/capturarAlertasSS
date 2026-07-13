# -*- coding: utf-8 -*-
"""
Ventana de Cese de Difusion.

Permite marcar personas como LOCALIZADAS a partir de su FUB. Dos modos:
  - Manual: se escribe un FUB (o varios, uno por linea) y se marca.
  - Lista: se pega una lista de FUBs de golpe (uno por linea).

En ambos casos el motor pone STATUS=LOCALIZADO, la fecha de localizado, y
pinta de verde las celdas con texto de cada fila encontrada.

Se apoya en capturar_alerta.marcar_localizados (la logica real).
"""

import threading
import queue
import datetime as dt
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path

import capturar_alerta as motor

COLOR_FONDO = "#f4f6f8"
COLOR_TITULO = "#1f7a4d"      # verde, para diferenciar de la captura normal
COLOR_TENUE = "#5b6b7b"
COLOR_INFO = "#2d6cdf"


class VentanaCese(tk.Toplevel):
    """Ventana (hija de la principal) para el proceso de cese de difusion."""

    def __init__(self, padre, excel_inicial=""):
        super().__init__(padre)
        self.title("Cese de difusion")
        self.geometry("720x680")
        self.minsize(640, 600)
        self.configure(bg=COLOR_FONDO)
        self.transient(padre)

        self.var_excel = tk.StringVar(value=excel_inicial)
        self.var_hoja = tk.StringVar()
        hoy = dt.date.today()
        self.var_dia = tk.StringVar(value=f"{hoy.day:02d}")
        self.var_mes = tk.StringVar(value=f"{hoy.month:02d}")
        self.var_anio = tk.StringVar(value=str(hoy.year))

        self.cola_logs = queue.Queue()
        self.procesando = False

        self._construir()
        self.after(100, self._vaciar_cola)
        # Si nos pasaron un Excel de entrada, cargar sus hojas de una vez
        if excel_inicial and Path(excel_inicial).is_file():
            self._cargar_hojas(excel_inicial)

    def _construir(self):
        cont = ttk.Frame(self, padding=16)
        cont.pack(fill="both", expand=True)

        ttk.Label(cont, text="Cese de difusion",
                  font=("Segoe UI", 16, "bold"),
                  foreground=COLOR_TITULO).pack(anchor="w")
        ttk.Label(cont,
                  text="Marca personas como LOCALIZADAS por su FUB. Escribe uno "
                       "o pega una lista (uno por linea).",
                  foreground=COLOR_TENUE, wraplength=680,
                  justify="left").pack(anchor="w", pady=(2, 12))

        # --- Excel y hoja ---
        f_excel = ttk.Frame(cont)
        f_excel.pack(fill="x", pady=4)
        ttk.Label(f_excel, text="Archivo de Excel:", width=16).pack(side="left")
        ttk.Entry(f_excel, textvariable=self.var_excel).pack(
            side="left", fill="x", expand=True, padx=(0, 6), ipady=3)
        ttk.Button(f_excel, text="Elegir...",
                   command=self._elegir_excel).pack(side="left")

        f_hoja = ttk.Frame(cont)
        f_hoja.pack(fill="x", pady=4)
        ttk.Label(f_hoja, text="Hoja:", width=16).pack(side="left")
        self.combo_hoja = ttk.Combobox(f_hoja, textvariable=self.var_hoja,
                                       state="readonly")
        self.combo_hoja.pack(side="left", fill="x", expand=True, ipady=2)
        self.combo_hoja.set("(elige primero el Excel)")

        # --- Fecha de localizado ---
        f_fecha = ttk.Frame(cont)
        f_fecha.pack(fill="x", pady=8)
        ttk.Label(f_fecha, text="Fecha de localizado:",
                  width=16).pack(side="left")
        ttk.Label(f_fecha, text="Dia:").pack(side="left")
        ttk.Entry(f_fecha, textvariable=self.var_dia, width=4).pack(
            side="left", padx=(2, 8))
        ttk.Label(f_fecha, text="Mes:").pack(side="left")
        ttk.Entry(f_fecha, textvariable=self.var_mes, width=4).pack(
            side="left", padx=(2, 8))
        ttk.Label(f_fecha, text="Año:").pack(side="left")
        ttk.Entry(f_fecha, textvariable=self.var_anio, width=6).pack(
            side="left", padx=(2, 8))
        ttk.Button(f_fecha, text="Hoy", command=self._fecha_hoy).pack(side="left")

        # --- Caja para los FUBs (sirve para uno o para lista pegada) ---
        ttk.Label(cont, text="FUB a localizar (uno por linea):",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(6, 2))
        self.caja_fubs = scrolledtext.ScrolledText(
            cont, height=8, wrap="word", font=("Consolas", 10))
        self.caja_fubs.pack(fill="x")
        ttk.Label(cont,
                  text="Puedes escribir un solo FUB, o pegar muchos de golpe "
                       "(cada uno en su propia linea).",
                  foreground=COLOR_TENUE, font=("Segoe UI", 8)).pack(anchor="w")

        # --- Boton procesar ---
        self.boton = ttk.Button(cont, text="Marcar como localizados",
                                command=self._al_procesar)
        self.boton.pack(pady=12)

        # --- Resultados ---
        ttk.Label(cont, text="Resultados:",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.caja_log = scrolledtext.ScrolledText(
            cont, height=9, wrap="word", state="disabled",
            font=("Consolas", 9))
        self.caja_log.pack(fill="both", expand=True, pady=(2, 0))

    # ---------- Acciones ----------
    def _elegir_excel(self):
        ruta = filedialog.askopenfilename(
            title="Archivo Excel",
            filetypes=[("Archivos de Excel", "*.xlsx *.xlsm"),
                       ("Todos los archivos", "*.*")])
        if ruta:
            self.var_excel.set(ruta)
            self._cargar_hojas(ruta)

    def _cargar_hojas(self, ruta_excel):
        hojas = motor.listar_hojas(ruta_excel)
        if hojas:
            self.combo_hoja.configure(values=hojas)
            self.var_hoja.set(hojas[0])
        else:
            self.combo_hoja.configure(values=[])
            self.combo_hoja.set("(no se pudieron leer las hojas)")

    def _fecha_hoy(self):
        hoy = dt.date.today()
        self.var_dia.set(f"{hoy.day:02d}")
        self.var_mes.set(f"{hoy.month:02d}")
        self.var_anio.set(str(hoy.year))

    def _leer_fecha(self):
        try:
            return dt.date(int(self.var_anio.get()), int(self.var_mes.get()),
                           int(self.var_dia.get())), None
        except (ValueError, TypeError):
            return None, ("La fecha de localizado no es valida. Revisa dia, "
                          "mes y año.")

    def _al_procesar(self):
        if self.procesando:
            return

        excel = self.var_excel.get().strip()
        hoja = self.var_hoja.get().strip()

        if not excel or not Path(excel).is_file():
            messagebox.showerror("Excel no valido",
                                 "Elige un archivo de Excel valido.")
            return
        if not hoja or hoja.startswith("("):
            messagebox.showwarning("Falta la hoja",
                                   "Elige el Excel y luego la hoja.")
            return

        fecha, err = self._leer_fecha()
        if err:
            messagebox.showwarning("Fecha invalida", err)
            return

        # Sacar los FUBs de la caja: uno por linea, sin vacios
        texto = self.caja_fubs.get("1.0", "end")
        fubs = [linea.strip() for linea in texto.splitlines() if linea.strip()]
        if not fubs:
            messagebox.showwarning(
                "Sin FUB",
                "Escribe al menos un FUB (o pega una lista) para localizar.")
            return

        self.caja_log.configure(state="normal")
        self.caja_log.delete("1.0", "end")
        self.caja_log.configure(state="disabled")

        self.procesando = True
        self.boton.configure(text="Procesando...", state="disabled")

        hilo = threading.Thread(
            target=self._trabajo,
            args=(excel, fubs, hoja, fecha),
            daemon=True)
        hilo.start()

    def _trabajo(self, excel, fubs, hoja, fecha):
        try:
            resumen = motor.marcar_localizados(
                excel, fubs, hoja=hoja, fecha_localizado=fecha,
                log=self._log_hilo)
            self.cola_logs.put(("RESUMEN", resumen))
        except Exception as e:
            self.cola_logs.put(("LOG", f"ERROR inesperado: {e}"))
        finally:
            self.cola_logs.put(("FIN", None))

    def _log_hilo(self, mensaje):
        self.cola_logs.put(("LOG", str(mensaje)))

    def _vaciar_cola(self):
        try:
            while True:
                tipo, dato = self.cola_logs.get_nowait()
                if tipo == "LOG":
                    self._escribir(dato)
                elif tipo == "RESUMEN":
                    self._mostrar_resumen(dato)
                elif tipo == "FIN":
                    self.procesando = False
                    self.boton.configure(text="Marcar como localizados",
                                         state="normal")
        except queue.Empty:
            pass
        self.after(100, self._vaciar_cola)

    def _escribir(self, texto):
        self.caja_log.configure(state="normal")
        self.caja_log.insert("end", texto + "\n")
        self.caja_log.see("end")
        self.caja_log.configure(state="disabled")

    def _mostrar_resumen(self, resumen):
        loc = len(resumen.get("localizados", []))
        no = len(resumen.get("no_encontrados", []))
        texto = f"Localizados: {loc}\nNo encontrados: {no}"
        if no:
            texto += ("\n\nAlgunos FUB no se encontraron en la hoja. Revisa "
                      "el detalle en el area de resultados.")
            messagebox.showwarning("Proceso terminado", texto)
        else:
            messagebox.showinfo("Proceso terminado", texto)
