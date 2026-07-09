# -*- coding: utf-8 -*-
"""
Modulo de configuracion de campos para la Captura de Alertas Nacionales.

Este modulo es el "cerebro" que permite que el usuario defina, sin tocar el
codigo, que campos se leen del PDF y a que columna del Excel van.

La configuracion se guarda en un archivo JSON en una subcarpeta 'config'
junto al programa. El modulo se encarga de:
  - crear una configuracion por defecto si no existe
  - leerla y validarla (si esta dañada, avisa y usa el respaldo)
  - guardarla de forma segura (validando antes, respaldando el archivo bueno)

ESTRUCTURA de cada campo en la configuracion:
{
    "clave":       "municipio",          # nombre interno (no cambia)
    "etiquetas_pdf": ["MUNICIPIO:"],     # como aparece en el PDF (editable)
    "encabezados_excel": ["MUNICIPIO"],  # como se llama la columna (editable)
    "tipo":        "simple",             # simple | especial
    "obligatorio": false                 # si es indispensable
}

- Los campos "simple" se extraen con el patron etiqueta: valor. El usuario
  puede editar sus etiquetas y encabezados con seguridad.
- Los campos "especial" tienen logica propia en el codigo (nombre, lugar,
  caracteristicas fisicas, etc.). El usuario puede editar sus encabezados de
  Excel, pero NO la forma de extraerlos (por eso no exponemos su patron).
"""

import json
import shutil
import sys
from pathlib import Path
from datetime import datetime


# Ubicacion del archivo de configuracion: subcarpeta 'config' junto al programa.
def _carpeta_base() -> Path:
    """
    Devuelve la carpeta donde debe vivir la config, funcione como script .py
    o como .exe empaquetado con PyInstaller.

    El detalle importante: cuando PyInstaller usa --onefile, al abrir el .exe
    todo se descomprime en una carpeta TEMPORAL del sistema, y __file__ apunta
    ahi. Esa carpeta la borra Windows al cerrar, asi que si guardaramos la
    config ahi, se perderia. Por eso, si detectamos que corremos como .exe
    (sys.frozen), usamos la carpeta REAL donde esta el ejecutable
    (sys.executable), que es la que el usuario ve y que no se borra.
    """
    if getattr(sys, "frozen", False):
        # Estamos corriendo como .exe: la carpeta real es la del ejecutable.
        return Path(sys.executable).resolve().parent
    # Estamos corriendo como script normal .py: la carpeta de este archivo.
    return Path(__file__).resolve().parent


def _carpeta_config() -> Path:
    """Devuelve la carpeta 'config' junto al ejecutable/script, creandola."""
    # Colgamos la subcarpeta 'config' de la carpeta base (que ya distingue
    # entre correr como .py o como .exe). exist_ok=True evita que truene si ya
    # existia. Nota: esa carpeta debe tener permiso de escritura, asi que el
    # .exe no debe ponerse en un lugar de solo lectura (ej. Archivos de programa).
    base = _carpeta_base()
    carpeta = base / "config"
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta


ARCHIVO_CONFIG = "campos.json"


# ---------------------------------------------------------------------------
# Configuracion POR DEFECTO. Es la que se crea la primera vez, y la que se
# restaura si el archivo se pierde o se daña. Refleja el formato actual del
# PDF de Alerta Nacional.
# ---------------------------------------------------------------------------
def configuracion_por_defecto() -> dict:
    """Devuelve la configuracion inicial con todos los campos que hoy funcionan."""
    # Cada campo lleva 4 datos clave: la 'clave' (nombre interno que usa el
    # codigo), las 'etiquetas_pdf' (como aparece en el PDF), los
    # 'encabezados_excel' (nombre de la columna destino) y el 'tipo'.
    #
    # El 'tipo' es lo importante para entender que puede tocar el usuario:
    #   - sistema:  lo genera el programa (expediente, fecha de hoy). No se lee
    #               del PDF, por eso 'etiquetas_pdf' va vacio.
    #   - especial: se lee del PDF pero con logica propia en el codigo (nombre,
    #               fechas, caracteristicas fisicas...). El usuario no toca su
    #               extraccion, solo a que columna va.
    #   - simple:   el clasico "Etiqueta: valor". Totalmente configurable.
    campos = [
        # --- Campos del SISTEMA (los genera el programa, no vienen del PDF) ---
        {"clave": "expediente", "etiquetas_pdf": [],
         "encabezados_excel": ["EXPEDIENTE"], "tipo": "sistema",
         "obligatorio": False},
        {"clave": "fecha_registro", "etiquetas_pdf": [],
         "encabezados_excel": ["FECHA DE REGISTRO"], "tipo": "sistema",
         "obligatorio": False},

        # --- Campos ESPECIALES (logica propia en el codigo) ---
        {"clave": "nombre", "etiquetas_pdf": ["Nombre Social:"],
         "encabezados_excel": ["NOMBRE"], "tipo": "especial",
         "obligatorio": True},
        {"clave": "fub", "etiquetas_pdf": ["Folio Unico de Identificacion"],
         "encabezados_excel": ["FUB", "Folio Unico de Identificacion"],
         "tipo": "especial", "obligatorio": True},
        {"clave": "lugar_hechos", "etiquetas_pdf": ["Lugar de los hechos:"],
         "encabezados_excel": ["LUGAR DE HECHOS", "Lugar de los hechos"],
         "tipo": "especial", "obligatorio": False},
        {"clave": "municipio", "etiquetas_pdf": ["MUNICIPIO:"],
         "encabezados_excel": ["MUNICIPIO"], "tipo": "especial",
         "obligatorio": False},
        {"clave": "fecha_hecho", "etiquetas_pdf": ["Fecha de hechos:"],
         "encabezados_excel": ["FECHA DE HECHO", "Fecha de hechos"],
         "tipo": "especial", "obligatorio": False},
        {"clave": "fecha_percato", "etiquetas_pdf": ["Fecha de percato:"],
         "encabezados_excel": ["FECHA DE PERCATO"], "tipo": "especial",
         "obligatorio": False},
        {"clave": "caracteristicas_fisicas",
         "etiquetas_pdf": ["Caracteristicas fisicas:"],
         "encabezados_excel": ["Caracteristicas fisicas"], "tipo": "especial",
         "obligatorio": False},
        {"clave": "senas_particulares",
         "etiquetas_pdf": ["Senas particulares:"],
         "encabezados_excel": ["Senas particulares"], "tipo": "especial",
         "obligatorio": False},
        {"clave": "prendas_vestir", "etiquetas_pdf": ["Prendas de vestir:"],
         "encabezados_excel": ["Prendas de vestir"], "tipo": "especial",
         "obligatorio": False},

        # --- Campos SIMPLES (etiqueta: valor, totalmente configurables) ---
        {"clave": "edad_desaparecer",
         "etiquetas_pdf": ["Edad al momento de la desaparicion:"],
         "encabezados_excel": ["EDAD AL DESAPARECER",
                               "Edad al momento de la desaparicion"],
         "tipo": "simple", "obligatorio": False},
        {"clave": "edad_actual", "etiquetas_pdf": ["Edad Actual:"],
         "encabezados_excel": ["EDAD ACTUAL"], "tipo": "simple",
         "obligatorio": False},
        {"clave": "sexo", "etiquetas_pdf": ["Sexo:"],
         "encabezados_excel": ["SEXO"], "tipo": "simple", "obligatorio": False},
        {"clave": "genero", "etiquetas_pdf": ["Genero:"],
         "encabezados_excel": ["GENERO"], "tipo": "simple",
         "obligatorio": False},
        {"clave": "nacionalidad", "etiquetas_pdf": ["Nacionalidad:"],
         "encabezados_excel": ["NACIONALIDAD"], "tipo": "simple",
         "obligatorio": False},
        {"clave": "lugar_nacimiento", "etiquetas_pdf": ["Lugar de nacimiento:"],
         "encabezados_excel": ["Lugar de nacimiento"], "tipo": "simple",
         "obligatorio": False},
        {"clave": "habla_espanol", "etiquetas_pdf": ["¿Habla español?:"],
         "encabezados_excel": ["Habla espanol"], "tipo": "simple",
         "obligatorio": False},
        {"clave": "lengua_indigena",
         "etiquetas_pdf": ["Idioma o lengua indígena:"],
         "encabezados_excel": ["Idioma o lengua indigena", "Lengua indigena"],
         "tipo": "simple", "obligatorio": False},
        {"clave": "discapacidad", "etiquetas_pdf": ["Discapacidad:"],
         "encabezados_excel": ["Discapacidad"], "tipo": "simple",
         "obligatorio": False},
        {"clave": "hora_hecho", "etiquetas_pdf": ["Hora de hechos:"],
         "encabezados_excel": ["Hora de hechos", "Hora de hecho"],
         "tipo": "simple", "obligatorio": False},
        {"clave": "hora_percato", "etiquetas_pdf": ["Hora de percato:"],
         "encabezados_excel": ["Hora de percato"], "tipo": "simple",
         "obligatorio": False},
        {"clave": "carpeta_investigacion",
         "etiquetas_pdf": ["Carpeta de investigación"],
         "encabezados_excel": ["Carpeta de investigacion"], "tipo": "simple",
         "obligatorio": False},
        {"clave": "autoridad", "etiquetas_pdf": ["Autoridades Competentes:"],
         "encabezados_excel": ["AUTORIDAD QUE INGRESO EL REPORTE",
                               "Autoridades Competentes", "Autoridad"],
         "tipo": "simple", "obligatorio": False},
    ]
    # Ademas de los campos, guardamos una version (por si algun dia cambia el
    # formato y hay que migrar) y la fecha de la ultima actualizacion.
    return {
        "version": 1,
        "actualizado": datetime.now().isoformat(timespec="seconds"),
        "campos": campos,
    }


# ---------------------------------------------------------------------------
# Validacion: garantiza que una configuracion tenga la forma correcta ANTES
# de usarla o guardarla. Devuelve (True, "") si es valida, o (False, motivo).
# ---------------------------------------------------------------------------
def validar_configuracion(config) -> tuple:
    """Revisa la estructura de una configuracion. Devuelve (valida, motivo)."""
    # Esta funcion es el portero: nadie usa ni guarda una config sin pasar por
    # aqui. Devolvemos un motivo en texto para poder decirle al usuario EXACTO
    # que esta mal, en vez de un "error" generico.
    if not isinstance(config, dict):
        return False, "la configuracion no tiene el formato esperado"
    campos = config.get("campos")
    if not isinstance(campos, list) or not campos:
        return False, "la configuracion no tiene lista de campos"

    # Recorremos campo por campo revisando lo minimo indispensable. De paso
    # llevamos un set de claves ya vistas para cachar duplicados.
    claves_vistas = set()
    for i, campo in enumerate(campos, 1):
        if not isinstance(campo, dict):
            return False, f"el campo #{i} no tiene el formato correcto"
        clave = campo.get("clave", "")
        if not isinstance(clave, str) or not clave.strip():
            return False, f"el campo #{i} no tiene 'clave'"
        if clave in claves_vistas:
            return False, f"la clave '{clave}' esta repetida"
        claves_vistas.add(clave)

        # Un campo sin columna de destino no sirve: no sabriamos donde
        # escribirlo. Las etiquetas SI pueden ir vacias (sistema/especial).
        encs = campo.get("encabezados_excel")
        if not isinstance(encs, list) or not any(str(e).strip() for e in encs):
            return False, (f"el campo '{clave}' no tiene ningun "
                           "encabezado de Excel")
        etqs = campo.get("etiquetas_pdf")
        if not isinstance(etqs, list):
            return False, f"el campo '{clave}' tiene etiquetas mal formadas"

    # Sin nombre no sabemos donde va la fila; sin fub no hay como detectar
    # duplicados. Estos dos no son negociables, aunque el usuario los borre.
    if "nombre" not in claves_vistas:
        return False, "falta el campo obligatorio 'nombre'"
    if "fub" not in claves_vistas:
        return False, "falta el campo obligatorio 'fub'"

    return True, ""


# ---------------------------------------------------------------------------
# Guardar de forma segura: valida primero; si es valida, respalda el archivo
# bueno anterior y escribe el nuevo. Nunca pisa lo bueno con algo malo.
# ---------------------------------------------------------------------------
def guardar_configuracion(config) -> tuple:
    """
    Guarda la configuracion de forma segura. Devuelve (exito, mensaje).
    - Valida ANTES de escribir. Si no es valida, NO guarda nada.
    - Respalda el archivo bueno anterior en campos.bak.json.
    """
    # Regla de oro: validar ANTES de tocar nada. Si la config nueva viene mal,
    # cortamos aqui y el archivo bueno de siempre ni se entera.
    valida, motivo = validar_configuracion(config)
    if not valida:
        return False, f"No se guardo: {motivo}."

    carpeta = _carpeta_config()
    ruta = carpeta / ARCHIVO_CONFIG
    respaldo = carpeta / (ARCHIVO_CONFIG.replace(".json", ".bak.json"))

    # Antes de pisar el archivo actual, guardamos una copia .bak. Asi, si el
    # guardado nuevo saliera mal por lo que sea, todavia tenemos de donde tirar.
    if ruta.exists():
        try:
            shutil.copy(ruta, respaldo)
        except Exception:
            pass  # si falla el respaldo, aun asi intentamos guardar

    config["actualizado"] = datetime.now().isoformat(timespec="seconds")
    try:
        # Escritura atomica: escribimos primero a un .tmp y solo al final lo
        # renombramos sobre el archivo real. Si la maquina se apaga a medias, el
        # archivo bueno queda intacto y lo unico dañado es el .tmp, que no
        # importa. Renombrar es una operacion instantanea, escribir no.
        temp = carpeta / (ARCHIVO_CONFIG + ".tmp")
        with open(temp, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        temp.replace(ruta)
        return True, "Configuracion guardada correctamente."
    except Exception as e:
        return False, f"Error al guardar la configuracion: {e}"


# ---------------------------------------------------------------------------
# Cargar: lee el archivo. Si no existe, lo crea por defecto. Si esta dañado,
# intenta el respaldo; si tampoco sirve, regenera por defecto. Nunca truena.
# ---------------------------------------------------------------------------
def cargar_configuracion(log=lambda m: None) -> dict:
    """
    Carga la configuracion, protegida en varias capas:
      1) si no existe, la crea por defecto
      2) si esta dañada, intenta el respaldo
      3) si el respaldo tampoco sirve, regenera por defecto
    'log' recibe mensajes de lo que fue pasando (opcional).
    Siempre devuelve una configuracion valida y usable.
    """
    carpeta = _carpeta_config()
    ruta = carpeta / ARCHIVO_CONFIG
    respaldo = carpeta / (ARCHIVO_CONFIG.replace(".json", ".bak.json"))

    # La idea de esta funcion es que PASE LO QUE PASE, siempre devuelve una
    # config usable. Vamos probando de la mejor opcion a la peor, en cascada.

    # Capa 1 - no hay archivo (primera vez, o alguien lo borro): lo creamos de
    # cero con la config de fabrica.
    if not ruta.exists():
        log("No habia configuracion. Se creo una por defecto.")
        config = configuracion_por_defecto()
        guardar_configuracion(config)
        return config

    # Capa 2 - el archivo existe: intentamos leerlo y que pase la validacion.
    # Si todo bien, este es el camino normal y feliz.
    config = _intentar_leer(ruta)
    if config is not None:
        valida, motivo = validar_configuracion(config)
        if valida:
            return config
        log(f"La configuracion tenia un problema ({motivo}).")

    # Capa 3 - el archivo principal estaba dañado o corrupto: tiramos del
    # respaldo .bak. Si sirve, lo usamos y de paso rehacemos el principal.
    if respaldo.exists():
        config = _intentar_leer(respaldo)
        if config is not None:
            valida, _ = validar_configuracion(config)
            if valida:
                log("Se restauro la configuracion desde el respaldo.")
                guardar_configuracion(config)  # rehacer el archivo principal
                return config

    # Capa 4 - ni el principal ni el respaldo sirvieron. Ultimo recurso:
    # volvemos a la config de fabrica. Nunca dejamos al usuario sin nada.
    log("No se pudo recuperar la configuracion. Se regenero por defecto.")
    config = configuracion_por_defecto()
    guardar_configuracion(config)
    return config


def _intentar_leer(ruta: Path):
    """Intenta leer y parsear un JSON. Devuelve el dict o None si falla."""
    # Un helper chiquito para no repetir el try/except cada vez. Si el archivo
    # no abre o el JSON esta roto, devolvemos None y quien llama decide que hacer.
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Adaptadores: convierten la configuracion en las estructuras que el motor
# ya sabe usar (el mapa de encabezados para mapear_columnas).
# ---------------------------------------------------------------------------
# El motor no sabe nada de nuestro formato de config con "tipo", "obligatorio",
# etc. Estas dos funciones son el traductor: le entregan justo lo que espera,
# ni mas ni menos. Asi el motor y la config quedan desacoplados.
def mapa_encabezados_desde_config(config) -> dict:
    """
    Convierte la configuracion en el formato {clave: [encabezados]} que usa
    la funcion mapear_columnas del motor.
    """
    # Esto le sirve al motor para saber, por cada campo, en que columna del
    # Excel escribirlo (buscando cualquiera de esos encabezados posibles).
    mapa = {}
    for campo in config.get("campos", []):
        clave = campo.get("clave")
        encs = [str(e).strip() for e in campo.get("encabezados_excel", [])
                if str(e).strip()]
        if clave and encs:
            mapa[clave] = encs
    return mapa


def campos_simples_desde_config(config) -> dict:
    """
    Devuelve {clave: [etiquetas_pdf]} solo de los campos tipo 'simple'.
    El motor los usa para extraer esos campos con el patron etiqueta: valor.
    """
    # Ojo el filtro: solo los 'simple'. Los especiales tienen su propia logica
    # en el motor, asi que sus etiquetas no las tocamos desde aqui.
    simples = {}
    for campo in config.get("campos", []):
        if campo.get("tipo") == "simple":
            clave = campo.get("clave")
            etqs = [str(e).strip() for e in campo.get("etiquetas_pdf", [])
                    if str(e).strip()]
            if clave and etqs:
                simples[clave] = etqs
    return simples