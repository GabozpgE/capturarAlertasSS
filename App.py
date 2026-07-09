# -*- coding: utf-8 -*-
"""
Interfaz grafica para la captura de Alertas Nacionales.

Solo recoge rutas y opciones y las pasa al motor (capturar_alerta.py).

Para ejecutar:
    python app.py
"""

import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path

import capturar_alerta as motor
import config_campos as cfg
from ventana_config import VentanaConfig

COLOR_FONDO = "#f4f6f8"
COLOR_TITULO = "#1f3a5f"
COLOR_TEXTO_TENUE = "#5b6b7b"
COLOR_INFO = "#2d6cdf"

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
        "Selecciona el archivo de Excel donde se guardaran los registros.\n\n"
        "Importante: cierra el archivo en Excel antes de procesar. Si esta "
        "abierto, el programa no podra guardar los cambios.\n\n"
        "El programa hace una copia de seguridad automatica antes de "
        "escribir."
    ),
    "hojas": (
        "Hojas donde registrar",
        "Puedes registrar en una o en dos hojas al mismo tiempo.\n\n"
        "Marca la casilla de cada hoja donde quieras que se guarden los "
        "registros y elige, en su menu, cual es esa hoja.\n\n"
        "  - Hoja completa: normalmente la que tiene TODAS las columnas.\n"
        "  - Hoja personalizada: la que preparaste con solo las columnas "
        "que te interesan.\n\n"
        "El programa detecta las columnas por su encabezado, asi que cada "
        "hoja recibe solo los datos de las columnas que tenga.\n\n"
        "Debes marcar al menos una."
    ),
    "procesados": (
        "Carpeta de procesados",
        "Carpeta donde se moveran los PDF despues de procesarlos.\n\n"
        "Se ordenan asi:\n"
        "  - Agregados con exito: en esta carpeta.\n"
        "  - Duplicados: en una subcarpeta 'duplicados'.\n"
        "  - Con error: en una subcarpeta 'errores'.\n\n"
        "Si la carpeta no existe, el programa la crea sola."
    ),
}


class AppCaptura:
    def __init__(self, raiz):
        self.raiz = raiz
        raiz.title("Captura de Alertas Nacionales")
        raiz.geometry("840x680")
        raiz.minsize(740, 620)
        raiz.configure(bg=COLOR_FONDO)

        self.var_carpeta = tk.StringVar()
        self.var_excel = tk.StringVar()
        self.var_procesados = tk.StringVar()
        # Hoja completa y personalizada: cada una con su activador y su nombre
        self.var_usar_completa = tk.BooleanVar(value=True)
        self.var_hoja_completa = tk.StringVar()
        self.var_usar_personal = tk.BooleanVar(value=False)
        self.var_hoja_personal = tk.StringVar()

        self.cola_logs = queue.Queue()
        self.procesando = False

        # Cargar la configuracion de campos. Detectamos si se acaba de crear
        # (no existia) para mostrar el modal de bienvenida una sola vez.
        self._config_recien_creada = False
        self.config = cfg.cargar_configuracion(log=self._marcar_config_creada)

        self._configurar_estilos()
        self._construir_interfaz()
        self.raiz.after(100, self._vaciar_cola)
        # Mostrar bienvenida despues de que la ventana este lista
        if self._config_recien_creada:
            self.raiz.after(400, self._mostrar_bienvenida)

    def _marcar_config_creada(self, mensaje):
        """El cargador avisa por 'log'; si creo config nueva, lo recordamos."""
        if "creo una por defecto" in mensaje or "regenero" in mensaje:
            self._config_recien_creada = True

    def _mostrar_bienvenida(self):
        messagebox.showinfo(
            "Bienvenido",
            "Se creo una configuracion de campos por defecto.\n\n"
            "El programa ya esta listo para usarse. Si algun dia cambia el "
            "formato del PDF o los nombres de las columnas del Excel, puedes "
            "ajustar los campos con el boton 'Configurar campos', sin ayuda "
            "de un programador.")

    def _configurar_estilos(self):
        e = ttk.Style()
        try:
            e.theme_use("clam")
        except tk.TclError:
            pass
        e.configure("TFrame", background=COLOR_FONDO)
        e.configure("TLabel", background=COLOR_FONDO)
        e.configure("Card.TFrame", background="white")
        e.configure("Titulo.TLabel", background=COLOR_FONDO,
                    foreground=COLOR_TITULO, font=("Segoe UI", 17, "bold"))
        e.configure("Sub.TLabel", background=COLOR_FONDO,
                    foreground=COLOR_TEXTO_TENUE, font=("Segoe UI", 10))
        e.configure("Campo.TLabel", background="white",
                    font=("Segoe UI", 10, "bold"))
        e.configure("Paso.TLabel", background="white",
                    foreground=COLOR_INFO, font=("Segoe UI", 11, "bold"))
        e.configure("Check.TCheckbutton", background="white",
                    font=("Segoe UI", 10, "bold"))
        e.configure("Procesar.TButton", font=("Segoe UI", 11, "bold"),
                    padding=10)

    def _construir_interfaz(self):
        cont = ttk.Frame(self.raiz, padding=18)
        cont.pack(fill="both", expand=True)

        # Encabezado con titulo a la izquierda y boton de config a la derecha
        cab = ttk.Frame(cont)
        cab.pack(fill="x")
        izq = ttk.Frame(cab)
        izq.pack(side="left", fill="x", expand=True)
        ttk.Label(izq, text="Captura de Alertas Nacionales",
                  style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(izq, text="Completa los pasos y presiona Procesar.",
                  style="Sub.TLabel").pack(anchor="w", pady=(2, 14))
        ttk.Button(cab, text="Configurar campos",
                   command=self._abrir_configuracion).pack(side="right",
                                                           anchor="n")

        tarjeta = ttk.Frame(cont, style="Card.TFrame", padding=16)
        tarjeta.pack(fill="x")

        self._fila(tarjeta, "1", "Carpeta de PDFs", self.var_carpeta,
                   self._elegir_carpeta_pdfs, "carpeta")
        self._separador(tarjeta)
        self._fila(tarjeta, "2", "Archivo de Excel", self.var_excel,
                   self._elegir_excel, "excel")
        self._separador(tarjeta)
        self._bloque_hojas(tarjeta, "3")
        self._separador(tarjeta)
        self._fila(tarjeta, "4", "Carpeta de procesados", self.var_procesados,
                   self._elegir_carpeta_procesados, "procesados")

        self.boton = ttk.Button(cont, text="Procesar", style="Procesar.TButton",
                                command=self._al_presionar_procesar)
        self.boton.pack(pady=16)

        ttk.Label(cont, text="Resultados", style="Sub.TLabel").pack(anchor="w")
        self.caja_log = scrolledtext.ScrolledText(
            cont, height=10, wrap="word", state="disabled",
            font=("Consolas", 9), relief="solid", borderwidth=1)
        self.caja_log.pack(fill="both", expand=True, pady=(4, 0))

    def _separador(self, padre):
        ttk.Separator(padre, orient="horizontal").pack(fill="x", pady=8)

    def _fila(self, padre, num, etiqueta, variable, comando, clave_ayuda):
        fila = ttk.Frame(padre, style="Card.TFrame")
        fila.pack(fill="x")
        cab = ttk.Frame(fila, style="Card.TFrame")
        cab.pack(fill="x")
        ttk.Label(cab, text=f"Paso {num}", style="Paso.TLabel").pack(side="left")
        ttk.Label(cab, text="  " + etiqueta, style="Campo.TLabel").pack(side="left")
        self._boton_info(cab, clave_ayuda)
        linea = ttk.Frame(fila, style="Card.TFrame")
        linea.pack(fill="x", pady=(4, 0))
        ttk.Entry(linea, textvariable=variable).pack(
            side="left", fill="x", expand=True, padx=(0, 8), ipady=3)
        ttk.Button(linea, text="Elegir...", command=comando).pack(side="left")

    def _bloque_hojas(self, padre, num):
        """Paso 3: dos hojas activables, cada una con casilla + desplegable."""
        fila = ttk.Frame(padre, style="Card.TFrame")
        fila.pack(fill="x")
        cab = ttk.Frame(fila, style="Card.TFrame")
        cab.pack(fill="x")
        ttk.Label(cab, text=f"Paso {num}", style="Paso.TLabel").pack(side="left")
        ttk.Label(cab, text="  Hojas donde registrar",
                  style="Campo.TLabel").pack(side="left")
        self._boton_info(cab, "hojas")

        # Fila hoja completa
        f1 = ttk.Frame(fila, style="Card.TFrame")
        f1.pack(fill="x", pady=(6, 0))
        ttk.Checkbutton(f1, text="Hoja completa", style="Check.TCheckbutton",
                        variable=self.var_usar_completa,
                        command=self._actualizar_estado_combos).pack(side="left")
        self.combo_completa = ttk.Combobox(
            f1, textvariable=self.var_hoja_completa, state="disabled", width=30)
        self.combo_completa.pack(side="left", padx=(10, 0), fill="x", expand=True)

        # Fila hoja personalizada
        f2 = ttk.Frame(fila, style="Card.TFrame")
        f2.pack(fill="x", pady=(6, 0))
        ttk.Checkbutton(f2, text="Hoja personalizada", style="Check.TCheckbutton",
                        variable=self.var_usar_personal,
                        command=self._actualizar_estado_combos).pack(side="left")
        self.combo_personal = ttk.Combobox(
            f2, textvariable=self.var_hoja_personal, state="disabled", width=30)
        self.combo_personal.pack(side="left", padx=(10, 0), fill="x", expand=True)

    def _boton_info(self, padre, clave_ayuda):
        b = tk.Label(padre, text=" ? ", fg="white", bg=COLOR_INFO,
                     font=("Segoe UI", 9, "bold"), cursor="hand2")
        b.pack(side="left", padx=8)
        b.bind("<Button-1>", lambda e: self._mostrar_ayuda(clave_ayuda))

    def _mostrar_ayuda(self, clave):
        titulo, texto = AYUDAS[clave]
        messagebox.showinfo(titulo, texto)

    def _actualizar_estado_combos(self):
        """Habilita el desplegable de cada hoja solo si su casilla esta marcada."""
        self.combo_completa.configure(
            state="readonly" if self.var_usar_completa.get() else "disabled")
        self.combo_personal.configure(
            state="readonly" if self.var_usar_personal.get() else "disabled")

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
        hojas = motor.listar_hojas(ruta_excel)
        if hojas:
            self.combo_completa.configure(values=hojas)
            self.combo_personal.configure(values=hojas)
            # Sugerencias sensatas: completa = primera; personalizada = segunda
            self.var_hoja_completa.set(hojas[0])
            if len(hojas) > 1:
                self.var_hoja_personal.set(hojas[1])
            else:
                self.var_hoja_personal.set(hojas[0])
            self._actualizar_estado_combos()
        else:
            for combo in (self.combo_completa, self.combo_personal):
                combo.configure(values=[])
            self.var_hoja_completa.set("")
            self.var_hoja_personal.set("")

    def _elegir_carpeta_procesados(self):
        ruta = filedialog.askdirectory(title="Carpeta para PDF procesados")
        if ruta:
            self.var_procesados.set(ruta)

    def _abrir_configuracion(self):
        """Abre la ventana de configuracion de campos."""
        if self.procesando:
            messagebox.showinfo("Espera",
                                "Termina el proceso actual antes de configurar.")
            return
        VentanaConfig(self.raiz, al_guardar=self._recargar_config)

    def _recargar_config(self):
        """Se llama tras guardar la config: recarga la copia en memoria."""
        self.config = cfg.cargar_configuracion()
        messagebox.showinfo(
            "Configuracion actualizada",
            "Los cambios se aplicaran en el proximo proceso.")

    # ---------- Procesar ----------
    def _al_presionar_procesar(self):
        if self.procesando:
            return

        carpeta = self.var_carpeta.get().strip()
        excel = self.var_excel.get().strip()
        procesados = self.var_procesados.get().strip()

        if not carpeta or not excel or not procesados:
            messagebox.showwarning(
                "Faltan datos",
                "Completa las rutas de los pasos 1, 2 y 4 antes de procesar.")
            return
        if not Path(carpeta).is_dir():
            messagebox.showerror("Carpeta no valida",
                                 f"La carpeta de PDF no existe:\n{carpeta}")
            return
        if not Path(excel).is_file():
            messagebox.showerror("Excel no valido",
                                 f"No encuentro el archivo Excel:\n{excel}")
            return

        # Reunir las hojas activas (sin repetir)
        hojas = []
        if self.var_usar_completa.get():
            h = self.var_hoja_completa.get().strip()
            if h:
                hojas.append(h)
        if self.var_usar_personal.get():
            h = self.var_hoja_personal.get().strip()
            if h and h not in hojas:
                hojas.append(h)

        if not hojas:
            messagebox.showwarning(
                "Falta elegir hoja",
                "Marca al menos una hoja (completa o personalizada) y "
                "eligela en su menu.")
            return

        self.caja_log.configure(state="normal")
        self.caja_log.delete("1.0", "end")
        self.caja_log.configure(state="disabled")

        self.procesando = True
        self.boton.configure(text="Procesando...", state="disabled")

        hilo = threading.Thread(
            target=self._trabajo_en_segundo_plano,
            args=(carpeta, excel, procesados, hojas),
            daemon=True)
        hilo.start()

    def _trabajo_en_segundo_plano(self, carpeta, excel, procesados, hojas):
        try:
            resumen = motor.procesar_carpeta(
                carpeta, excel, procesados,
                hojas=hojas, config=self.config, log=self._log_desde_hilo)
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