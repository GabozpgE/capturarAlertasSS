# -*- coding: utf-8 -*-
"""
Ventana de configuracion de campos.

Permite al usuario ver, editar, agregar y quitar los campos que el programa
lee del PDF y escribe en el Excel, sin tocar el codigo.

Diseño (Opcion 2): una lista clara de campos; para editar uno, se abre una
ventanita chica enfocada solo en ese campo. Menos abrumador para el usuario.

Se apoya en el modulo config_campos (el "cerebro" que guarda/valida/respalda).
"""

import tkinter as tk
from tkinter import messagebox, ttk

import config_campos as cfg

COLOR_FONDO = "#f4f6f8"
COLOR_TITULO = "#1f3a5f"
COLOR_INFO = "#2d6cdf"
COLOR_TENUE = "#5b6b7b"

# Explicacion de cada tipo de campo, para que el usuario entienda que puede tocar
EXPLICA_TIPO = {
    "sistema": "Campo que genera el programa (no viene del PDF). "
               "Solo puedes cambiar a que columna va.",
    "especial": "Campo con lectura automatica especial. "
                "Solo puedes cambiar a que columna va.",
    "simple": "Campo normal. Puedes cambiar tanto la etiqueta del PDF "
              "como la columna del Excel.",
}


class VentanaConfig(tk.Toplevel):
    """Ventana (hija de la principal) para configurar los campos."""

    def __init__(self, padre, al_guardar=None):
        super().__init__(padre)
        # 'al_guardar' es un callback que la ventana principal nos pasa para
        # enterarse cuando el usuario guarda algo. Si no lo pasan, no pasa nada.
        self.al_guardar = al_guardar  # callback opcional al guardar
        self.title("Configurar campos")
        self.geometry("720x620")
        self.minsize(640, 540)
        self.configure(bg=COLOR_FONDO)
        # transient + grab_set = ventana modal. Mientras esta abierta, el
        # usuario no puede tocar la ventana principal por detras. Es a proposito:
        # no queremos que procese PDFs mientras esta cambiando la config.
        self.transient(padre)
        self.grab_set()  # modal: bloquea la ventana principal mientras esta abierta

        # Trabajamos sobre una copia recien cargada del disco. Todo lo que el
        # usuario edite vive aqui en memoria hasta que le da a Guardar; asi, si
        # se arrepiente y cancela, el archivo real nunca se toco.
        self.config = cfg.cargar_configuracion()

        self._construir()
        self._refrescar_lista()

    # ---------- Construccion ----------
    def _construir(self):
        cont = ttk.Frame(self, padding=16)
        cont.pack(fill="both", expand=True)

        ttk.Label(cont, text="Configuracion de campos",
                  font=("Segoe UI", 15, "bold"),
                  foreground=COLOR_TITULO).pack(anchor="w")
        ttk.Label(cont,
                  text="Aqui defines que campos se leen del PDF y a que "
                       "columna del Excel van.\nHaz clic en un campo y luego "
                       "en 'Editar' para modificarlo.",
                  foreground=COLOR_TENUE).pack(anchor="w", pady=(2, 12))

        # La tabla de campos es un Treeview de 4 columnas. Solo mostramos datos
        # (show="headings"), no hay arbol desplegable ni nada de eso.
        marco_tabla = ttk.Frame(cont)
        marco_tabla.pack(fill="both", expand=True)

        cols = ("campo", "tipo", "etiqueta", "columna")
        self.tabla = ttk.Treeview(marco_tabla, columns=cols, show="headings",
                                  height=14)
        self.tabla.heading("campo", text="Campo")
        self.tabla.heading("tipo", text="Tipo")
        self.tabla.heading("etiqueta", text="Etiqueta en el PDF")
        self.tabla.heading("columna", text="Columna en el Excel")
        self.tabla.column("campo", width=150)
        self.tabla.column("tipo", width=80, anchor="center")
        self.tabla.column("etiqueta", width=200)
        self.tabla.column("columna", width=200)
        self.tabla.pack(side="left", fill="both", expand=True)

        # Barra de scroll pegada a la tabla, porque con 20+ campos no caben
        # todos a la vista.
        scroll = ttk.Scrollbar(marco_tabla, orient="vertical",
                               command=self.tabla.yview)
        scroll.pack(side="right", fill="y")
        self.tabla.configure(yscrollcommand=scroll.set)
        # Doble clic sobre una fila = editar ese campo. Es el atajo que la gente
        # espera por instinto, aparte del boton de abajo.
        self.tabla.bind("<Double-1>", lambda e: self._editar_seleccionado())

        # Botones que actuan sobre el campo seleccionado en la lista.
        botones = ttk.Frame(cont)
        botones.pack(fill="x", pady=(10, 0))
        ttk.Button(botones, text="Editar campo",
                   command=self._editar_seleccionado).pack(side="left")
        ttk.Button(botones, text="Agregar campo nuevo",
                   command=self._agregar_campo).pack(side="left", padx=6)
        ttk.Button(botones, text="Quitar campo",
                   command=self._quitar_campo).pack(side="left")

        # Pie: Guardar y Cancelar a la derecha (lo mas usado), y Restaurar a la
        # izquierda, aparte, porque es una accion destructiva que no queremos
        # que se apriete sin querer.
        pie = ttk.Frame(cont)
        pie.pack(fill="x", pady=(14, 0))
        ttk.Button(pie, text="Guardar cambios",
                   command=self._guardar).pack(side="right")
        ttk.Button(pie, text="Cancelar",
                   command=self.destroy).pack(side="right", padx=6)
        ttk.Button(pie, text="Restaurar valores por defecto",
                   command=self._restaurar).pack(side="left")

    # ---------- Datos <-> tabla ----------
    def _refrescar_lista(self):
        """Vuelve a pintar la tabla con los campos actuales de self.config."""
        # Borron y cuenta nueva: vaciamos la tabla y la volvemos a llenar desde
        # cero. Es mas simple y menos propenso a bugs que andar actualizando
        # filas sueltas cada vez que algo cambia.
        for item in self.tabla.get_children():
            self.tabla.delete(item)
        for i, campo in enumerate(self.config.get("campos", [])):
            # Usamos el indice de la lista como iid de la fila. Asi, cuando el
            # usuario selecciona una fila, sabemos exactamente que campo es sin
            # tener que buscarlo por nombre.
            etq = ", ".join(campo.get("etiquetas_pdf", [])) or "(automatico)"
            col = ", ".join(campo.get("encabezados_excel", []))
            tipo = campo.get("tipo", "simple")
            self.tabla.insert("", "end", iid=str(i),
                              values=(campo.get("clave", ""), tipo, etq, col))

    def _indice_seleccionado(self):
        sel = self.tabla.selection()
        if not sel:
            return None
        return int(sel[0])

    # ---------- Acciones ----------
    def _editar_seleccionado(self):
        idx = self._indice_seleccionado()
        if idx is None:
            messagebox.showinfo("Editar campo",
                                "Primero selecciona un campo de la lista.")
            return
        # Abrimos el editor chico para ESE campo. El truco esta en el lambda:
        # cuando el editor confirme, nos devuelve el campo ya editado y nosotros
        # lo metemos de vuelta en la posicion idx de la lista.
        campo = self.config["campos"][idx]
        EditorCampo(self, campo, es_nuevo=False,
                    al_confirmar=lambda c: self._aplicar_edicion(idx, c))

    def _agregar_campo(self):
        # Los campos que agrega el usuario siempre son 'simple'. Los especiales
        # y de sistema tienen logica en el codigo, asi que no dejamos que se
        # inventen unos nuevos de esos tipos desde la interfaz.
        nuevo = {"clave": "", "etiquetas_pdf": [""],
                 "encabezados_excel": [""], "tipo": "simple",
                 "obligatorio": False}
        EditorCampo(self, nuevo, es_nuevo=True,
                    al_confirmar=self._aplicar_nuevo)

    # Estos dos son los callbacks que el editor llama al aceptar. Uno pisa el
    # campo existente, el otro agrega uno al final. Ambos repintan la tabla.
    def _aplicar_edicion(self, idx, campo):
        self.config["campos"][idx] = campo
        self._refrescar_lista()

    def _aplicar_nuevo(self, campo):
        self.config["campos"].append(campo)
        self._refrescar_lista()

    def _quitar_campo(self):
        idx = self._indice_seleccionado()
        if idx is None:
            messagebox.showinfo("Quitar campo",
                                "Primero selecciona un campo de la lista.")
            return
        campo = self.config["campos"][idx]
        clave = campo.get("clave", "")
        # nombre y fub son la columna vertebral del sistema: sin nombre no
        # sabemos donde escribir, y sin fub no podemos detectar duplicados. Asi
        # que los blindamos para que nadie los borre por accidente.
        if clave in ("nombre", "fub"):
            messagebox.showwarning(
                "No se puede quitar",
                f"El campo '{clave}' es indispensable para el sistema y "
                "no se puede quitar.")
            return
        # Para lo demas, pedimos confirmacion antes de borrar. Nunca borres algo
        # de golpe sin preguntar.
        if messagebox.askyesno("Quitar campo",
                               f"Quitar el campo '{clave}'?\n"
                               "Ya no se leera ni registrara."):
            del self.config["campos"][idx]
            self._refrescar_lista()

    def _restaurar(self):
        # El boton de panico: si el usuario enredo todo, esto vuelve a la
        # configuracion de fabrica. Avisamos claro que va a perder sus cambios.
        if messagebox.askyesno(
                "Restaurar valores por defecto",
                "Esto reemplaza toda la configuracion por la original.\n"
                "Se perderan tus cambios personalizados. Continuar?"):
            self.config = cfg.configuracion_por_defecto()
            self._refrescar_lista()

    def _guardar(self):
        # Ojo: no validamos aqui a mano. Delegamos en guardar_configuracion, que
        # valida ANTES de escribir. Si algo esta mal, no guarda nada y nos
        # devuelve el motivo, que le mostramos al usuario tal cual. El archivo
        # bueno anterior queda intacto pase lo que pase.
        ok, msg = cfg.guardar_configuracion(self.config)
        if ok:
            messagebox.showinfo("Guardado", msg)
            # Le avisamos a la ventana principal que recargue su copia.
            if self.al_guardar:
                self.al_guardar()
            self.destroy()
        else:
            messagebox.showerror("No se pudo guardar", msg)


class EditorCampo(tk.Toplevel):
    """Ventanita chica para editar UN campo (Opcion 2)."""

    def __init__(self, padre, campo, es_nuevo, al_confirmar):
        super().__init__(padre)
        self.al_confirmar = al_confirmar
        self.campo = dict(campo)  # copia de trabajo
        self.es_nuevo = es_nuevo
        self.tipo = self.campo.get("tipo", "simple")

        self.title("Nuevo campo" if es_nuevo else "Editar campo")
        self.geometry("520x420")
        self.configure(bg=COLOR_FONDO)
        self.transient(padre)
        self.grab_set()

        self._construir()

    def _construir(self):
        cont = ttk.Frame(self, padding=16)
        cont.pack(fill="both", expand=True)

        titulo = "Agregar campo nuevo" if self.es_nuevo else \
            f"Editar: {self.campo.get('clave','')}"
        ttk.Label(cont, text=titulo, font=("Segoe UI", 13, "bold"),
                  foreground=COLOR_TITULO).pack(anchor="w", pady=(0, 4))

        # Explicacion segun el tipo
        ttk.Label(cont, text=EXPLICA_TIPO.get(self.tipo, ""),
                  foreground=COLOR_TENUE, wraplength=480,
                  justify="left").pack(anchor="w", pady=(0, 12))

        # Nombre interno (clave): editable solo si es nuevo
        ttk.Label(cont, text="Nombre del campo:",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.var_clave = tk.StringVar(value=self.campo.get("clave", ""))
        entrada_clave = ttk.Entry(cont, textvariable=self.var_clave)
        entrada_clave.pack(fill="x", pady=(2, 2))
        if not self.es_nuevo:
            entrada_clave.configure(state="disabled")  # no se renombra la clave
            ttk.Label(cont, text="(el nombre interno no se cambia)",
                      foreground=COLOR_TENUE,
                      font=("Segoe UI", 8)).pack(anchor="w")
        ttk.Label(cont, text="").pack()  # espaciador

        # Etiqueta del PDF: solo editable en campos 'simple'
        editable_etq = (self.tipo == "simple")
        ttk.Label(cont, text="Etiqueta como aparece en el PDF:",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        etq_actual = ", ".join(self.campo.get("etiquetas_pdf", []))
        self.var_etq = tk.StringVar(value=etq_actual)
        e_etq = ttk.Entry(cont, textvariable=self.var_etq)
        e_etq.pack(fill="x", pady=(2, 2))
        if not editable_etq:
            e_etq.configure(state="disabled")
            ttk.Label(cont, text="(este campo se lee de forma automatica)",
                      foreground=COLOR_TENUE,
                      font=("Segoe UI", 8)).pack(anchor="w")
        else:
            ttk.Label(cont,
                      text="Si hay varias, separalas con comas. "
                           "Ej.: Sexo:, Género:",
                      foreground=COLOR_TENUE,
                      font=("Segoe UI", 8)).pack(anchor="w")
        ttk.Label(cont, text="").pack()

        # Columna del Excel: editable siempre
        ttk.Label(cont, text="Columna(s) en el Excel:",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        col_actual = ", ".join(self.campo.get("encabezados_excel", []))
        self.var_col = tk.StringVar(value=col_actual)
        ttk.Entry(cont, textvariable=self.var_col).pack(fill="x", pady=(2, 2))
        ttk.Label(cont,
                  text="El encabezado exacto de la columna. Si tiene varios "
                       "nombres posibles, separalos con comas.",
                  foreground=COLOR_TENUE, wraplength=480,
                  font=("Segoe UI", 8)).pack(anchor="w")

        # Botones
        pie = ttk.Frame(cont)
        pie.pack(fill="x", pady=(16, 0))
        ttk.Button(pie, text="Aceptar",
                   command=self._aceptar).pack(side="right")
        ttk.Button(pie, text="Cancelar",
                   command=self.destroy).pack(side="right", padx=6)

    def _aceptar(self):
        # Aqui recogemos lo que el usuario tecleo y lo validamos antes de darlo
        # por bueno. Estas son validaciones de "forma"; la validacion de fondo
        # (que no haya claves repetidas, etc.) la hace el cerebro al guardar.
        clave = self.var_clave.get().strip()
        if not clave:
            messagebox.showwarning("Falta el nombre",
                                   "El campo necesita un nombre.")
            return
        # El nombre interno solo lo validamos en campos NUEVOS (los existentes
        # tienen la clave bloqueada). Lo forzamos a formato tipo variable:
        # minusculas, numeros y guion bajo. Nada de espacios ni acentos, porque
        # esa clave se usa internamente como identificador, no es para mostrar.
        if self.es_nuevo:
            import re
            if not re.match(r"^[a-z0-9_]+$", clave):
                messagebox.showwarning(
                    "Nombre no valido",
                    "El nombre interno solo puede tener letras minusculas, "
                    "numeros y guion bajo (_), sin espacios ni acentos.\n"
                    "Ejemplo: telefono_contacto")
                return

        # Los campos de etiquetas y columnas se escriben separados por comas.
        # Aqui los partimos y de paso tiramos los vacios (por si dejaron una
        # coma suelta o espacios de mas).
        etiquetas = [e.strip() for e in self.var_etq.get().split(",")
                     if e.strip()]
        columnas = [c.strip() for c in self.var_col.get().split(",")
                    if c.strip()]

        # Sin columna de destino el campo no sirve de nada: no sabriamos donde
        # escribirlo en el Excel. La etiqueta si puede ir vacia (campos
        # especiales/sistema que no se leen del PDF).
        if not columnas:
            messagebox.showwarning(
                "Falta la columna",
                "Indica al menos una columna del Excel para este campo.")
            return

        # Todo bien: volcamos los datos al campo y se lo devolvemos a quien nos
        # abrio (via el callback). El que decide donde va es la ventana grande.
        self.campo["clave"] = clave
        self.campo["etiquetas_pdf"] = etiquetas
        self.campo["encabezados_excel"] = columnas
        self.al_confirmar(self.campo)
        self.destroy()