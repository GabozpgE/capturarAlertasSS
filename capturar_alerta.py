# -*- coding: utf-8 -*-
"""

Uso:
    python capturar_alerta.py "ruta/al/flyer.pdf"
    python capturar_alerta.py "ruta/al/flyer.pdf" --excel "ruta/al/base.xlsx"

Requisitos:
    pip install pdfplumber openpyxl
"""

import re
import sys
import shutil
import unicodedata
import datetime as dt
from pathlib import Path

import pdfplumber
import openpyxl
from copy import copy

# ----------------- CONFIGURACION -----------------
EXCEL_DEFAULT = "BASE ALERTAS NACIONAL 2026 (JOAO).xlsx"
HOJA_DEFAULT = "Hoja1"          # se usa solo si no se especifica otra
ANIO_EXPEDIENTE = 2026          # ano que va en AN/{id}/{anio}
COL_ID = 1                       # columna A
PRIMERA_FILA_DATOS = 2           # los datos empiezan en la fila 2

# Mapa de columnas (numero de columna en el Excel)
COLS = {
    "expediente": 2,    # B
    "fecha_registro": 3,  # C
    "nombre": 4,        # D
    "edad_desaparecer": 5,  # E
    "edad_actual": 6,   # F
    "sexo": 7,          # G
    "genero": 8,        # H
    "nacionalidad": 9,  # I
    "fub": 10,          # J
    "fecha_hecho": 11,  # K
    "fecha_percato": 12,  # L
    "lugar_hechos": 13,  # M
    "autoridad": 14,    # N
}
# -------------------------------------------------


class ErrorPDF(Exception):
    """Error con mensaje claro al procesar un PDF (para mostrar al usuario)."""
    pass


def leer_texto_pdf(ruta_pdf: str) -> str:
    """Extrae todo el texto del PDF en una sola cadena.

    Lanza ErrorPDF con un mensaje entendible si el archivo esta dañado,
    protegido con contraseña, o no es un PDF valido.
    """
    try:
        partes = []
        with pdfplumber.open(ruta_pdf) as pdf:
            if len(pdf.pages) == 0:
                raise ErrorPDF("el PDF no tiene paginas")
            for pagina in pdf.pages:
                txt = pagina.extract_text() or ""
                partes.append(txt)
        texto = "\n".join(partes)
        if not texto.strip():
            raise ErrorPDF("el PDF no contiene texto legible "
                           "(podria ser una imagen escaneada)")
        return texto
    except ErrorPDF:
        raise
    except Exception as e:
        # pdfplumber/pdfminer lanzan varios tipos; los unificamos
        raise ErrorPDF(f"no se pudo leer el PDF (archivo dañado o "
                       f"protegido): {type(e).__name__}")


def buscar(patron: str, texto: str, grupo: int = 1, flags=re.IGNORECASE) -> str:
    """Aplica una regex y devuelve el grupo limpio, o '' si no encuentra."""
    m = re.search(patron, texto, flags)
    if not m:
        return ""
    return m.group(grupo).strip()


def fecha_a_objeto(fecha_str: str):
    """
    Convierte 'DD/MM/YYYY' (o 'DD-MM-YYYY') a un objeto date real de Python.

    Devolver una fecha real (no texto) es importante: la celda del Excel ya
    tiene su propio formato de fecha (DD-MM-AA), asi que Excel la mostrara
    correctamente y NO marcara el triangulo verde de "fecha como texto".

    Si el texto no tiene forma de fecha, devuelve la cadena original en
    mayusculas (para no perder informacion en casos raros).
    """
    fecha_str = (fecha_str or "").strip()
    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", fecha_str)
    if not m:
        return fecha_str.upper()
    d, mth, y = m.groups()
    y = int(y)
    if y < 100:               # año de 2 digitos -> asumir 2000+
        y += 2000
    try:
        return dt.date(y, int(mth), int(d))
    except ValueError:
        # fecha imposible (ej. 31/02); devuelve el texto para no romper
        return fecha_str.upper()


def normalizar_lugar(lugar_raw: str) -> str:
    """
    Del PDF 'ESTADO: CIUDAD DE MEXICO , MUNICIPIO: IZTAPALAPA'
    extrae solo el estado.
    """
    m = re.search(r"ESTADO:\s*(.+?)\s*,\s*MUNICIPIO", lugar_raw, re.IGNORECASE)
    if m:
        return m.group(1).strip().upper()
    return lugar_raw.strip().upper()


def extraer_datos(texto: str) -> dict:
    """Extrae cada campo del flyer usando las etiquetas como anclas."""
    d = {}

    # Toma el nombre del titulo de la ficha (suele venir bien escrito).
    # Esta entre 'FICHA DE BUSQUEDA DE PERSONA DESAPARECIDA' y
    # 'Folio Unico de Identificacion'.
    d["nombre"] = buscar(
        r"FICHA DE B[UÚ]SQUEDA DE PERSONA DESAPARECIDA\s*(.+?)\s*"
        r"Folio [UÚ]nico de Identificaci[oó]n",
        texto, flags=re.IGNORECASE | re.DOTALL)
    # Respaldo: si el titulo no se encuentra, usa el campo Nombre Social.
    if not d["nombre"]:
        d["nombre"] = buscar(
            r"Nombre Social:\s*(.+?)\s*(?:Nacionalidad:|\n)", texto)

    d["edad_desaparecer"] = buscar(
        r"Edad al momento de la desaparici[oó]n:\s*(\d+)", texto)
    d["edad_actual"] = buscar(r"Edad Actual:\s*(\d+)", texto)

    d["sexo"] = buscar(r"Sexo:\s*(.+?)\s*(?:Genero:|G[eé]nero:)", texto)
    d["genero"] = buscar(r"G[eé]nero:\s*(.+?)\s*(?:Nombre Social:|\n)", texto)
    d["nacionalidad"] = buscar(
        r"Nacionalidad:\s*(.+?)\s*(?:¿Habla|Habla|\n)", texto)

    d["fub"] = buscar(
        r"Folio [UÚ]nico de Identificaci[oó]n\s*([A-Z0-9\-]+)", texto)

    fh = buscar(r"Fecha de hechos:\s*([\d/]+)", texto)
    fp = buscar(r"Fecha de percato:\s*([\d/]+)", texto)
    d["fecha_hecho"] = fecha_a_objeto(fh)
    d["fecha_percato"] = fecha_a_objeto(fp)

    lugar = buscar(r"Lugar de los hechos:\s*(.+?)\s*(?:Carpeta|\n)", texto)
    d["lugar_hechos"] = normalizar_lugar(lugar)

    aut = buscar(r"Autoridades Competentes:\s*(.+?)\s*(?:Si tienes|\n)", texto)
    d["autoridad"] = aut.upper()

    return d


def validar_datos(datos: dict):
    """
    Revisa que el PDF tenga los datos minimos indispensables.
    Lanza ErrorPDF con un mensaje claro si falta algo critico.
    Esto evita meter registros vacios o incompletos al Excel.
    """
    faltantes = []
    if not (datos.get("nombre") or "").strip():
        faltantes.append("NOMBRE")
    if not (datos.get("fub") or "").strip():
        faltantes.append("FUB")
    if faltantes:
        raise ErrorPDF("no parece una ficha de Alerta valida; "
                       f"falta(n): {', '.join(faltantes)}")


def quitar_acentos(texto: str) -> str:
    """
    Quita acentos de las vocales (Á->A, É->E, etc.) pero CONSERVA la enie (Ñ),
    porque en nombres mexicanos la enie es una letra distinta, no un acento.
    Tambien conserva la dieresis de la u (Ü) si apareciera.
    """
    if not isinstance(texto, str):
        return texto
    # Protege la enie y la U con dieresis con marcadores temporales
    protegidos = (texto
                  .replace("Ñ", "\x00").replace("ñ", "\x01")
                  .replace("Ü", "\x02").replace("ü", "\x03"))
    # Descompone los caracteres acentuados y elimina las marcas de acento
    sin_acentos = unicodedata.normalize("NFD", protegidos)
    sin_acentos = "".join(c for c in sin_acentos
                          if unicodedata.category(c) != "Mn")
    # Restaura la enie y la dieresis
    return (sin_acentos
            .replace("\x00", "Ñ").replace("\x01", "ñ")
            .replace("\x02", "Ü").replace("\x03", "ü"))


def limpiar_para_sql(texto: str) -> str:
    """
    Deja solo caracteres compatibles con la base de datos destino:
    letras (incluida Ñ/ñ), numeros, espacio, guion '-' y diagonal '/'.
    Elimina cualquier otro simbolo (puntos, comas, parentesis, etc.).
    Tambien colapsa espacios multiples y recorta los extremos.

    Se asume que el texto ya viene SIN ACENTOS (quitar_acentos se aplica
    antes), por eso aqui solo permitimos A-Z, a-z, Ñ y ñ entre las letras.
    """
    if not isinstance(texto, str):
        return texto
    # Conserva solo lo permitido. \w incluiria '_' y acentos, por eso
    # definimos el conjunto a mano de forma explicita.
    permitido = re.sub(r"[^A-Za-z0-9Ññ /\-]", "", texto)
    # Colapsa espacios repetidos que pudieran quedar al quitar simbolos
    permitido = re.sub(r"\s+", " ", permitido).strip()
    return permitido


def a_mayusculas(datos: dict) -> dict:
    """
    Limpia los campos de texto para que sean aptos para Excel y SQL:
      1) mayusculas
      2) sin acentos (conservando Ñ)
      3) sin caracteres extraños (solo letras, numeros, espacio, - y /)
    No toca las fechas ya formateadas.
    """
    no_tocar = {"fecha_hecho", "fecha_percato"}
    salida = {}
    for k, v in datos.items():
        if isinstance(v, str) and k not in no_tocar:
            salida[k] = limpiar_para_sql(quitar_acentos(v.upper()))
        else:
            salida[k] = v
    return salida


def fila_para_nuevo_registro(ws):
    """
    Encuentra la fila donde escribir el nuevo registro.

    Importante: en este Excel las columnas ID y EXPEDIENTE ya vienen
    pre-llenadas hasta muchas filas adelante, aunque la fila este vacia.
    Por eso NO podemos guiarnos por el ID. El indicador de un registro
    real es que la columna NOMBRE (D) tenga texto.

    Devuelve la primera fila cuyo NOMBRE este vacio (= primer hueco real).
    """
    col_nombre = COLS["nombre"]
    fila = PRIMERA_FILA_DATOS
    while True:
        val = ws.cell(row=fila, column=col_nombre).value
        if val is None or str(val).strip() == "":
            return fila
        fila += 1


def escribir_fila(ws, datos: dict):
    """
    Escribe un registro en la primera fila sin NOMBRE de un worksheet
    YA ABIERTO. No guarda el archivo (eso lo hace quien llama, una sola
    vez al final, lo cual es mas rapido y seguro al procesar varios PDFs).
    Devuelve (nuevo_id, fila, valores).
    """
    fila = fila_para_nuevo_registro(ws)
    fila_modelo = fila - 1

    # --- Copiar el formato de la fila anterior a la fila nueva ---
    # openpyxl no hereda estilos en celdas vacias, asi que clonamos
    # el formato (fuente, color, bordes, relleno, alineacion, formato
    # de numero) de la fila de arriba que ya tiene el estilo correcto.
    if fila_modelo >= PRIMERA_FILA_DATOS:
        for col in range(1, ws.max_column + 1):
            origen = ws.cell(row=fila_modelo, column=col)
            destino = ws.cell(row=fila, column=col)
            if origen.has_style:
                destino.font = copy(origen.font)
                destino.fill = copy(origen.fill)
                destino.border = copy(origen.border)
                destino.alignment = copy(origen.alignment)
                destino.number_format = origen.number_format
                destino.protection = copy(origen.protection)
        if fila_modelo in ws.row_dimensions and ws.row_dimensions[fila_modelo].height:
            ws.row_dimensions[fila].height = ws.row_dimensions[fila_modelo].height

    # El ID ya viene pre-llenado en la columna A. Lo respetamos.
    # Si por algun motivo estuviera vacio, lo calculamos a partir del de arriba.
    id_actual = ws.cell(row=fila, column=COL_ID).value
    if id_actual is None or str(id_actual).strip() == "":
        id_arriba = ws.cell(row=fila_modelo, column=COL_ID).value
        nuevo_id = int(id_arriba) + 1 if id_arriba not in (None, "") else 1
        ws.cell(row=fila, column=COL_ID, value=nuevo_id)
    else:
        nuevo_id = int(id_actual)

    # El EXPEDIENTE tambien suele venir pre-llenado. Solo lo ponemos
    # si esta vacio, para no pisar lo que ya tiene el Excel.
    exp_actual = ws.cell(row=fila, column=COLS["expediente"]).value
    if exp_actual is None or str(exp_actual).strip() == "":
        ws.cell(row=fila, column=COLS["expediente"],
                value=f"AN/{nuevo_id}/{ANIO_EXPEDIENTE}")
    expediente = ws.cell(row=fila, column=COLS["expediente"]).value

    # Fecha de registro = hoy, como fecha real (no texto) para que la
    # celda con formato DD-MM-AA la muestre bien sin triangulo verde.
    fecha_registro = dt.date.today()

    valores = {
        "expediente": expediente,
        "fecha_registro": fecha_registro,
        "nombre": datos["nombre"],
        "edad_desaparecer": datos["edad_desaparecer"],
        "edad_actual": datos["edad_actual"],
        "sexo": datos["sexo"],
        "genero": datos["genero"],
        "nacionalidad": datos["nacionalidad"],
        "fub": datos["fub"],
        "fecha_hecho": datos["fecha_hecho"],
        "fecha_percato": datos["fecha_percato"],
        "lugar_hechos": datos["lugar_hechos"],
        "autoridad": datos["autoridad"],
    }

    for clave, col in COLS.items():
        valor = valores[clave]
        # Las edades deben quedar como NUMERO, no como texto,
        # para que Excel no marque el triangulo verde.
        if clave in ("edad_desaparecer", "edad_actual"):
            try:
                valor = int(valor)
            except (ValueError, TypeError):
                pass  # si por algo no es numero, lo deja como viene
        ws.cell(row=fila, column=col, value=valor)

    return nuevo_id, fila, valores


def leer_fubs_existentes(ws) -> set:
    """
    Lee todos los FUB que ya estan registrados (columna J) y los devuelve
    en un conjunto normalizado (mayusculas, sin espacios) para comparar.
    Sirve para detectar duplicados.
    """
    col_fub = COLS["fub"]
    col_nombre = COLS["nombre"]
    fubs = set()
    fila = PRIMERA_FILA_DATOS
    while True:
        nombre = ws.cell(row=fila, column=col_nombre).value
        if nombre is None or str(nombre).strip() == "":
            break  # se acabaron los registros reales
        fub = ws.cell(row=fila, column=col_fub).value
        if fub is not None and str(fub).strip():
            fubs.add(str(fub).strip().upper())
        fila += 1
    return fubs


def listar_hojas(ruta_excel) -> list:
    """
    Devuelve la lista de nombres de hojas de un archivo Excel.
    La interfaz la usa para llenar el menu desplegable de seleccion de hoja.
    Si el archivo no se puede abrir, devuelve lista vacia.
    """
    try:
        wb = openpyxl.load_workbook(ruta_excel, read_only=True)
        nombres = list(wb.sheetnames)
        wb.close()
        return nombres
    except Exception:
        return []


def procesar_carpeta(carpeta_pdfs, ruta_excel, carpeta_procesados,
                     hoja=None, log=print):
    """
    Procesa todos los PDF de 'carpeta_pdfs':
      - extrae y valida datos de cada uno
      - los nuevos se escriben en el Excel y el PDF se mueve a 'procesados'
      - los duplicados (FUB ya registrado) se mueven a 'procesados/duplicados'
      - los que fallan se mueven a 'procesados/errores'
      - las subcarpetas 'duplicados' y 'errores' se crean SOLO si hace falta

    'log' es una funcion para reportar el avance.

    Devuelve un resumen: dict con listas 'agregados', 'duplicados', y
    'errores' (esta ultima con tuplas (nombre_pdf, motivo)).
    """
    carpeta_pdfs = Path(carpeta_pdfs)
    ruta_excel = Path(ruta_excel)
    carpeta_procesados = Path(carpeta_procesados)

    resumen = {"agregados": [], "duplicados": [], "errores": []}

    # --- Validaciones basicas de entrada ---
    if not carpeta_pdfs.is_dir():
        log(f"ERROR: la carpeta de PDFs no existe: {carpeta_pdfs}")
        return resumen
    if not ruta_excel.is_file():
        log(f"ERROR: no encuentro el Excel: {ruta_excel}")
        return resumen

    pdfs = sorted(carpeta_pdfs.glob("*.pdf"))
    if not pdfs:
        log("No hay archivos PDF en la carpeta.")
        return resumen

    # --- Respaldo de seguridad del Excel ---
    try:
        respaldo = ruta_excel.with_suffix(".bak.xlsx")
        shutil.copy(ruta_excel, respaldo)
        log(f"Respaldo creado: {respaldo.name}")
    except Exception as e:
        log(f"ERROR: no se pudo crear el respaldo del Excel ({e}). "
            "Se cancela el proceso por seguridad.")
        return resumen

    # --- Abrir el Excel una sola vez ---
    try:
        wb = openpyxl.load_workbook(ruta_excel)
    except PermissionError:
        log("ERROR: el Excel esta abierto. Cierralo y vuelve a intentar.")
        return resumen
    except Exception as e:
        log(f"ERROR: no se pudo abrir el Excel ({type(e).__name__}: {e}).")
        return resumen

    # --- Determinar la hoja a usar ---
    hoja_usar = hoja if hoja else HOJA_DEFAULT

    # --- Verificar que exista la hoja indicada ---
    if hoja_usar not in wb.sheetnames:
        log(f"ERROR: el Excel no tiene la hoja '{hoja_usar}'. "
            f"Hojas disponibles: {', '.join(wb.sheetnames)}.")
        return resumen
    ws = wb[hoja_usar]

    # --- Asegurar la carpeta de procesados (la principal si se crea) ---
    try:
        carpeta_procesados.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        log(f"ERROR: no se pudo crear la carpeta de procesados ({e}).")
        return resumen

    # Subcarpetas para duplicados y errores. NO se crean todavia:
    # solo cuando aparezca el primer duplicado / primer error.
    carpeta_duplicados = carpeta_procesados / "duplicados"
    carpeta_errores = carpeta_procesados / "errores"

    # FUBs ya existentes + los que agreguemos en esta tanda
    fubs_vistos = leer_fubs_existentes(ws)
    log(f"Registros existentes detectados: {len(fubs_vistos)} FUB unicos.")
    log("-" * 50)

    hubo_cambios = False

    for pdf in pdfs:
        try:
            texto = leer_texto_pdf(str(pdf))
            datos = a_mayusculas(extraer_datos(texto))
            validar_datos(datos)  # lanza ErrorPDF si faltan nombre o FUB

            fub = (datos.get("fub") or "").strip().upper()

            # Duplicado -> subcarpeta 'duplicados'
            if fub in fubs_vistos:
                log(f"  [=] {pdf.name}: duplicado (FUB ya registrado). Se omite.")
                resumen["duplicados"].append(pdf.name)
                _asegurar_carpeta(carpeta_duplicados)
                _mover(pdf, carpeta_duplicados, log)
                continue

            # Nuevo -> escribir y mover a 'procesados'
            nuevo_id, fila, _ = escribir_fila(ws, datos)
            fubs_vistos.add(fub)
            hubo_cambios = True
            log(f"  [+] {pdf.name}: agregado. ID {nuevo_id}, fila {fila} "
                f"({datos.get('nombre','')}).")
            resumen["agregados"].append(pdf.name)
            _mover(pdf, carpeta_procesados, log)

        except ErrorPDF as e:
            # Error "esperado" con mensaje claro -> subcarpeta 'errores'
            log(f"  [x] {pdf.name}: ERROR -> {e}")
            resumen["errores"].append((pdf.name, str(e)))
            _asegurar_carpeta(carpeta_errores)
            _mover(pdf, carpeta_errores, log)

        except Exception as e:
            # Error inesperado -> tambien a 'errores', con el tipo de error
            motivo = f"error inesperado ({type(e).__name__}: {e})"
            log(f"  [x] {pdf.name}: {motivo}")
            resumen["errores"].append((pdf.name, motivo))
            _asegurar_carpeta(carpeta_errores)
            _mover(pdf, carpeta_errores, log)

    # --- Guardar el Excel una sola vez al final ---
    if hubo_cambios:
        try:
            wb.save(ruta_excel)
        except PermissionError:
            log("ERROR al guardar: el Excel esta abierto. "
                "Cierralo; los cambios de esta tanda NO se guardaron.")
            return resumen
        except Exception as e:
            log(f"ERROR al guardar el Excel ({type(e).__name__}: {e}). "
                "Los cambios podrian no haberse guardado.")
            return resumen

    # --- Resumen final ---
    log("-" * 50)
    log(f"Listo. Agregados con exito: {len(resumen['agregados'])} | "
        f"Duplicados: {len(resumen['duplicados'])} | "
        f"Con error: {len(resumen['errores'])}.")

    # Detalle de los que fallaron, para revisarlos a mano
    if resumen["errores"]:
        log("")
        log("Archivos que NO se pudieron procesar:")
        for nombre, motivo in resumen["errores"]:
            log(f"   - {nombre}: {motivo}")

    return resumen


def _asegurar_carpeta(carpeta: Path):
    """Crea la carpeta solo si no existe. Se llama justo antes de necesitarla,
    para que las subcarpetas de duplicados/errores aparezcan unicamente
    cuando de verdad hay algo que mover ahi."""
    if not carpeta.exists():
        carpeta.mkdir(parents=True, exist_ok=True)


def _mover(pdf: Path, destino: Path, log):
    """Mueve un PDF a la carpeta destino. Si ya existe alli, le agrega
    un sufijo para no sobrescribir."""
    objetivo = destino / pdf.name
    if objetivo.exists():
        # Evitar choque de nombres: agrega _1, _2, etc.
        i = 1
        while True:
            cand = destino / f"{pdf.stem}_{i}{pdf.suffix}"
            if not cand.exists():
                objetivo = cand
                break
            i += 1
    try:
        shutil.move(str(pdf), str(objetivo))
    except Exception as e:
        log(f"      (no se pudo mover {pdf.name}: {e})")


def main():
    args = sys.argv[1:]

    # Modo carpeta:
    #   python capturar_alerta.py --carpeta "ruta/pdfs" --excel "base.xlsx"
    #                             --procesados "ruta/procesados"
    if "--carpeta" in args:
        carpeta = args[args.index("--carpeta") + 1]
        ruta_excel = EXCEL_DEFAULT
        if "--excel" in args:
            ruta_excel = args[args.index("--excel") + 1]
        if "--procesados" in args:
            procesados = args[args.index("--procesados") + 1]
        else:
            procesados = str(Path(carpeta) / "procesados")
        hoja = None
        if "--hoja" in args:
            hoja = args[args.index("--hoja") + 1]
        procesar_carpeta(carpeta, ruta_excel, procesados, hoja=hoja)
        return

    # Modo un solo PDF (compatibilidad con lo anterior):
    #   python capturar_alerta.py "flyer.pdf" [--excel "base.xlsx"]
    if not args:
        print('Uso:')
        print('  Un PDF:   python capturar_alerta.py "flyer.pdf" [--excel "base.xlsx"]')
        print('  Carpeta:  python capturar_alerta.py --carpeta "pdfs" '
              '--excel "base.xlsx" --procesados "procesados" [--hoja "Hoja1"]')
        sys.exit(1)

    ruta_pdf = args[0]
    ruta_excel = EXCEL_DEFAULT
    if "--excel" in args:
        ruta_excel = args[args.index("--excel") + 1]
    hoja_cli = args[args.index("--hoja") + 1] if "--hoja" in args else HOJA_DEFAULT

    if not Path(ruta_pdf).exists():
        print(f"No encuentro el PDF: {ruta_pdf}")
        sys.exit(1)
    if not Path(ruta_excel).exists():
        print(f"No encuentro el Excel: {ruta_excel}")
        sys.exit(1)

    respaldo = Path(ruta_excel).with_suffix(".bak.xlsx")
    shutil.copy(ruta_excel, respaldo)

    try:
        wb = openpyxl.load_workbook(ruta_excel)
    except PermissionError:
        print("ERROR: el Excel esta abierto. Cierralo y vuelve a intentar.")
        sys.exit(1)
    if hoja_cli not in wb.sheetnames:
        print(f"ERROR: el Excel no tiene la hoja '{hoja_cli}'. "
              f"Hojas disponibles: {', '.join(wb.sheetnames)}.")
        sys.exit(1)
    ws = wb[hoja_cli]

    try:
        texto = leer_texto_pdf(ruta_pdf)
        datos = a_mayusculas(extraer_datos(texto))
        validar_datos(datos)
    except ErrorPDF as e:
        print(f"\nERROR: {e}")
        return

    # Aviso de duplicado tambien en modo un PDF
    fub = (datos.get("fub") or "").strip().upper()
    if fub and fub in leer_fubs_existentes(ws):
        print(f"\n[=] Duplicado: el FUB {fub} ya esta registrado. No se agrega.")
        return

    nuevo_id, fila, valores = escribir_fila(ws, datos)
    wb.save(ruta_excel)

    print(f"\nRegistro agregado. ID {nuevo_id} en la fila {fila}.")
    print(f"Respaldo previo guardado en: {respaldo.name}\n")
    print("Datos capturados:")
    for k, v in valores.items():
        print(f"  {k:18s}: {v}")
    print("\nRevisa siempre el registro antes de darlo por bueno.")


if __name__ == "__main__":
    main()