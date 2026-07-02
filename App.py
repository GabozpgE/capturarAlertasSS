# -*- coding: utf-8 -*-
"""
Interfaz grafica para la captura de Alertas Nacionales.

Para ejecutar:
    python app.py

Requisitos: los mismos del motor (pdfplumber, openpyxl). Tkinter ya viene
incluido con Python, no se instala nada extra.
"""

import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path

import capturar_alerta as motor

# Paleta de colores simple y sobria
COLOR_FONDO = "#f4f6f8"
COLOR_TITULO = "#1f3a5f"
COLOR_TEXTO_TENUE = "#5b6b7b"
COLOR_INFO = "#2d6cdf"

# Textos de ayuda de cada campo (los muestra el boton "?")
AYUDAS = {
    "carpeta": (
        "Carpeta de PDFs",
        "Selecciona la carpeta donde guardaste los archivos PDF de las "
        "fichas de Alerta Nacional descargadas del correo.\n\n"
        "El programa leera TODOS los PDF que esten dentro de esa carpeta.\n\n"
        "Consejo: ten una carpeta dedicada solo para los PDF nuevos por "
        "procesar."
    ),
    "excel": (
        "Archivo de Excel",
        "Selecciona el archivo de Excel donde se guardaran los registros "
        "(tu base de datos).\n\n"
        "Importante: cierra el archivo en Excel antes de procesar. Si esta "
        "abierto, el programa no podra guardar los cambios.\n\n"
        "El programa hace una copia de seguridad automatica antes de "
        "escribir."
    ),
    "hoja": (
        "Hoja del Excel",
        "Elige la hoja (pestaña) del Excel donde van los registros.\n\n"
        "La lista se llena sola al elegir el archivo de Excel. Solo abre el "
        "menu y selecciona la hoja correcta.\n\n"
        
    ),
    "procesados": (
        "Carpeta de procesados",
        "Carpeta donde se moveran los PDF despues de procesarlos, para no "
        "mezclarlos con los nuevos.\n\n"
        "Se ordenan asi:\n"
        "  - Los agregados con exito: en esta carpeta.\n"
        "  - Los duplicados: en una subcarpeta 'duplicados'.\n"
        "  - Los que fallan: en una subcarpeta 'errores'.\n\n"
        "Si la carpeta no existe, el programa la crea sola."
    ),
}


class AppCaptura:
    """Ventana principal de la aplicacion."""

    def __init__(self, raiz):
        self.raiz = raiz
        raiz.title("Captura de Alertas Nacionales")
        raiz.geometry("820x620")
        raiz.minsize(720, 560)
        raiz.configure(bg=COLOR_FONDO)

        self.var_carpeta = tk.StringVar()
        self.var_excel = tk.StringVar()
        self.var_procesados = tk.StringVar()
        self.var_hoja = tk.StringVar()

        self.cola_logs = queue.Queue()
        self.procesando = False

        self._configurar_estilos()
        self._construir_interfaz()
        self.raiz.after(100, self._vaciar_cola)

    def _configurar_estilos(self):
        estilo = ttk.Style()
        try:
            estilo.theme_use("clam")
        except tk.TclError:
            pass
        estilo.configure("TFrame", background=COLOR_FONDO)
        estilo.configure("TLabel", background=COLOR_FONDO)
        estilo.configure("Card.TFrame", background="white", relief="solid",
                         borderwidth=1)
        estilo.configure("Titulo.TLabel", background=COLOR_FONDO,
                         foreground=COLOR_TITULO,
                         font=("Segoe UI", 17, "bold"))
        estilo.configure("Sub.TLabel", background=COLOR_FONDO,
                         foreground=COLOR_TEXTO_TENUE, font=("Segoe UI", 10))
        estilo.configure("Campo.TLabel", background="white",
                         font=("Segoe UI", 10, "bold"))
        estilo.configure("Procesar.TButton", font=("Segoe UI", 11, "bold"),
                         padding=10)
        estilo.configure("Paso.TLabel", background="white",
                         foreground=COLOR_INFO, font=("Segoe UI", 11, "bold"))

    def _construir_interfaz(self):
        cont = ttk.Frame(self.raiz, padding=18)
        cont.pack(fill="both", expand=True)

        # Encabezado
        ttk.Label(cont, text="Captura de Alertas Nacionales",
                  style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(cont,
                  text="Completa los 4 pasos y presiona Procesar.",
                  style="Sub.TLabel").pack(anchor="w", pady=(2, 14))

        # Tarjeta con los campos
        tarjeta = ttk.Frame(cont, style="Card.TFrame", padding=16)
        tarjeta.pack(fill="x")

        self._fila(tarjeta, "1", "Carpeta de PDFs", self.var_carpeta,
                   self._elegir_carpeta_pdfs, "carpeta")
        self._separador(tarjeta)
        self._fila(tarjeta, "2", "Archivo de Excel", self.var_excel,
                   self._elegir_excel, "excel")
        self._separador(tarjeta)
        self._fila_hoja(tarjeta, "3", "Hoja del Excel", "hoja")
        self._separador(tarjeta)
        self._fila(tarjeta, "4", "Carpeta de procesados", self.var_procesados,
                   self._elegir_carpeta_procesados, "procesados")

        # Boton procesar
        self.boton = ttk.Button(cont, text="Procesar",
                                style="Procesar.TButton",
                                command=self._al_presionar_procesar)
        self.boton.pack(pady=16)

        # Resultados
        ttk.Label(cont, text="Resultados",
                  style="Sub.TLabel").pack(anchor="w")
        self.caja_log = scrolledtext.ScrolledText(
            cont, height=12, wrap="word", state="disabled",
            font=("Consolas", 9), relief="solid", borderwidth=1)
        self.caja_log.pack(fill="both", expand=True, pady=(4, 0))

    def _separador(self, padre):
        ttk.Separator(padre, orient="horizontal").pack(fill="x", pady=8)

    def _fila(self, padre, num, etiqueta, variable, comando, clave_ayuda):
        """Fila estandar: numero de paso, etiqueta, campo, info y boton."""
        fila = ttk.Frame(padre, style="Card.TFrame")
        fila.configure(padding=0)
        fila.pack(fill="x")

        # Encabezado de la fila: paso + etiqueta + info
        cab = ttk.Frame(fila, style="Card.TFrame")
        cab.pack(fill="x")
        ttk.Label(cab, text=f"Paso {num}", style="Paso.TLabel").pack(side="left")
        ttk.Label(cab, text="  " + etiqueta, style="Campo.TLabel").pack(side="left")
        self._boton_info(cab, clave_ayuda)

        # Linea de entrada + boton elegir
        linea = ttk.Frame(fila, style="Card.TFrame")
        linea.pack(fill="x", pady=(4, 0))
        ttk.Entry(linea, textvariable=variable).pack(
            side="left", fill="x", expand=True, padx=(0, 8), ipady=3)
        ttk.Button(linea, text="Elegir...", command=comando).pack(side="left")

    def _fila_hoja(self, padre, num, etiqueta, clave_ayuda):
        """Fila especial: en vez de 'Elegir', un menu desplegable de hojas."""
        fila = ttk.Frame(padre, style="Card.TFrame")
        fila.pack(fill="x")

        cab = ttk.Frame(fila, style="Card.TFrame")
        cab.pack(fill="x")
        ttk.Label(cab, text=f"Paso {num}", style="Paso.TLabel").pack(side="left")
        ttk.Label(cab, text="  " + etiqueta, style="Campo.TLabel").pack(side="left")
        self._boton_info(cab, clave_ayuda)

        linea = ttk.Frame(fila, style="Card.TFrame")
        linea.pack(fill="x", pady=(4, 0))
        self.combo_hoja = ttk.Combobox(
            linea, textvariable=self.var_hoja, state="readonly")
        self.combo_hoja.pack(side="left", fill="x", expand=True, ipady=2)
        self.combo_hoja.set("(elige primero el Excel)")

    def _boton_info(self, padre, clave_ayuda):
        """Pequeño boton circular '?' que abre una ventanita con ayuda."""
        b = tk.Label(padre, text=" ? ", fg="white", bg=COLOR_INFO,
                     font=("Segoe UI", 9, "bold"), cursor="hand2")
        b.pack(side="left", padx=8)
        b.bind("<Button-1>", lambda e: self._mostrar_ayuda(clave_ayuda))

    def _mostrar_ayuda(self, clave):
        titulo, texto = AYUDAS[clave]
        messagebox.showinfo(titulo, texto)

    # ---------- Acciones ----------
    def _elegir_carpeta_pdfs(self):
        ruta = filedialog.askdirectory(title="Carpeta con los PDF")
        if ruta:
            self.var_carpeta.set(ruta)
            if not self.var_procesados.get():
                self.var_procesados.set(str(Path(ruta) / "procesados"))

    def _elegir_excel(self):
        ruta = filedialog.askopenfilename(
            title="Archivo Excel destino",
            filetypes=[("Archivos de Excel", "*.xlsx *.xlsm"),
                       ("Todos los archivos", "*.*")])
        if ruta:
            self.var_excel.set(ruta)
            self._cargar_hojas(ruta)

    def _cargar_hojas(self, ruta_excel):
        """Llena el desplegable de hojas leyendo el Excel elegido."""
        hojas = motor.listar_hojas(ruta_excel)
        if hojas:
            self.combo_hoja.configure(values=hojas)
            self.var_hoja.set(hojas[0])  # selecciona la primera por defecto
        else:
            self.combo_hoja.configure(values=[])
            self.combo_hoja.set("(no se pudieron leer las hojas)")

    def _elegir_carpeta_procesados(self):
        ruta = filedialog.askdirectory(title="Carpeta para PDF procesados")
        if ruta:
            self.var_procesados.set(ruta)

    # ---------- Procesar ----------
    def _al_presionar_procesar(self):
        if self.procesando:
            return

        carpeta = self.var_carpeta.get().strip()
        excel = self.var_excel.get().strip()
        procesados = self.var_procesados.get().strip()
        hoja = self.var_hoja.get().strip()

        if not carpeta or not excel or not procesados:
            messagebox.showwarning(
                "Faltan datos",
                "Completa las rutas de los pasos 1, 2 y 4 antes de procesar.")
            return
        if not hoja or hoja.startswith("("):
            messagebox.showwarning(
                "Falta la hoja",
                "Elige primero el Excel (paso 2) y luego la hoja (paso 3).")
            return
        if not Path(carpeta).is_dir():
            messagebox.showerror("Carpeta no valida",
                                 f"La carpeta de PDF no existe:\n{carpeta}")
            return
        if not Path(excel).is_file():
            messagebox.showerror("Excel no valido",
                                 f"No encuentro el archivo Excel:\n{excel}")
            return

        self.caja_log.configure(state="normal")
        self.caja_log.delete("1.0", "end")
        self.caja_log.configure(state="disabled")

        self.procesando = True
        self.boton.configure(text="Procesando...", state="disabled")

        hilo = threading.Thread(
            target=self._trabajo_en_segundo_plano,
            args=(carpeta, excel, procesados, hoja),
            daemon=True)
        hilo.start()

    def _trabajo_en_segundo_plano(self, carpeta, excel, procesados, hoja):
        try:
            resumen = motor.procesar_carpeta(
                carpeta, excel, procesados,
                hoja=hoja, log=self._log_desde_hilo)
            self.cola_logs.put(("RESUMEN", resumen))
        except Exception as e:
            self.cola_logs.put(("LOG", f"ERROR inesperado: {e}"))
        finally:
            self.cola_logs.put(("FIN", None))

    def _log_desde_hilo(self, mensaje):
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
                    self.boton.configure(text="Procesar", state="normal")
        except queue.Empty:
            pass
        self.raiz.after(100, self._vaciar_cola)

    def _escribir(self, texto):
        self.caja_log.configure(state="normal")
        self.caja_log.insert("end", texto + "\n")
        self.caja_log.see("end")
        self.caja_log.configure(state="disabled")

    def _mostrar_resumen(self, resumen):
        ag = len(resumen.get("agregados", []))
        du = len(resumen.get("duplicados", []))
        er = len(resumen.get("errores", []))
        texto = (f"Agregados con exito: {ag}\n"
                 f"Duplicados (omitidos): {du}\n"
                 f"Con error: {er}")
        if er:
            texto += ("\n\nHubo archivos con error. Revisa el detalle en el "
                      "area de resultados; esos PDF se movieron a la "
                      "subcarpeta 'errores'.")
            messagebox.showwarning("Proceso terminado", texto)
        else:
            messagebox.showinfo("Proceso terminado", texto)


def main():
    raiz = tk.Tk()
    AppCaptura(raiz)
    raiz.mainloop()


if __name__ == "__main__":
    main()