"""
Renombra PDFs de renuncias comparando datos extraidos por OCR contra un Excel.

Para cada PDF:
    1. Confirma que sea una renuncia.
    2. Lee la fecha del encabezado en la primera linea del documento.
    3. Lee el nombre de quien renuncia desde el bloque final:
       - despues de "Nombre:"; o
       - arriba de "Nombre y firma".
    4. Busca en el Excel:
       - Columna B: nombre de quien renuncia.
       - Columna E: fecha de baja.
       - Columna F: nuevo nombre del PDF.
    5. Si nombre y fecha coinciden, renombra el PDF con el valor de la
       columna F y extension .pdf.

Requisitos:
    python3 -m pip install pymupdf pandas openpyxl tqdm
    Tener tesseract instalado en el PATH o indicar su ruta con --tesseract.

Uso:
    python3 renombrar_renuncias_basico.py
    python3 renombrar_renuncias_basico.py /ruta/a/carpeta --excel "Archivo.xlsx"
    python3 renombrar_renuncias_basico.py --dry-run
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import shutil
import subprocess
import sys
import unicodedata
from dataclasses import dataclass
from datetime import date
from difflib import SequenceMatcher, get_close_matches
from pathlib import Path

try:
    import fitz  # PyMuPDF
    import pandas as pd
except ImportError as e:
    print(
        "Falta instalar una dependencia. Ejecuta:\n"
        "python3 -m pip install pymupdf pandas openpyxl tqdm\n"
        f"\nDetalle: {e}",
        file=sys.stderr,
    )
    sys.exit(1)

try:
    from tqdm import tqdm
except ImportError:
    def tqdm(iterable=None, **kwargs):
        return iterable if iterable is not None else []


OCR_DPI = 260
UMBRAL_COINCIDENCIA_NOMBRE = 0.74
REPORTE_DEFAULT = Path("reporte_renombrado_renuncias_basico.xlsx")
TESSERACT_EXE = "tesseract"

MESES = {
    "ENERO": 1,
    "FEBRERO": 2,
    "MARZO": 3,
    "ABRIL": 4,
    "MAYO": 5,
    "JUNIO": 6,
    "JULIO": 7,
    "AGOSTO": 8,
    "SEPTIEMBRE": 9,
    "SETIEMBRE": 9,
    "OCTUBRE": 10,
    "NOVIEMBRE": 11,
    "DICIEMBRE": 12,
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


@dataclass
class FilaExcel:
    nombre: str
    nombre_norm: str
    fecha_baja: date
    nuevo_nombre: str


def normalizar_texto(valor: object) -> str:
    texto = str(valor or "").upper()
    texto = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    texto = re.sub(r"[^A-Z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def normalizar_nombre(valor: object) -> str:
    texto = normalizar_texto(valor)
    palabras_basura = {
        "C", "CC", "LA", "EL", "DEL", "DE", "Y", "POR", "NOMBRE",
        "FIRMA", "RFC", "R", "F", "OO", "ATENTAMENTE", "PRESENTE",
    }
    palabras = [
        palabra
        for palabra in texto.split()
        if len(palabra) > 1 and palabra not in palabras_basura
    ]
    return " ".join(palabras)


def es_renuncia(texto_hoja: str) -> bool:
    texto = normalizar_texto(texto_hoja)[:2500]
    if "RENUNCIA" in texto:
        return True

    # Tolerancia para errores OCR como RENUNCTA / RENUNCIA pegada a otra palabra.
    for palabra in texto.split():
        if SequenceMatcher(None, palabra, "RENUNCIA").ratio() >= 0.82:
            return True
    return False


def corregir_numero_ocr(valor: str) -> int:
    valor = (
        valor.upper()
        .replace("O", "0")
        .replace("I", "1")
        .replace("L", "1")
        .replace("|", "1")
        .replace("Z", "3")
    )
    return int(valor)


def corregir_mes_ocr(valor: str) -> int | None:
    mes = normalizar_texto(valor)
    correcciones = {
        "AGOSTC": "AGOSTO",
        "DTCLEMBRE": "DICIEMBRE",
        "DICLEMBRE": "DICIEMBRE",
        "DICJEMBRE": "DICIEMBRE",
        "NOVLEMBRE": "NOVIEMBRE",
        "SEPTIEMBRE": "SEPTIEMBRE",
    }
    mes = correcciones.get(mes, mes)
    if mes in MESES:
        return MESES[mes]

    cercano = get_close_matches(mes, MESES.keys(), n=1, cutoff=0.60)
    return MESES[cercano[0]] if cercano else None


def extraer_fecha_renuncia(texto_hoja: str) -> date | None:
    texto = normalizar_texto(texto_hoja)
    encabezado = texto[:900]

    patrones = [
        # Xalapa, Ver., a 31 de mayo de 2025
        re.compile(
            r"(?:^|\s)A\s+([0-3ZO]?[0-9OIL])\s+D[A-Z]{0,4}\s+"
            r"([A-Z]{4,12})\s+D[A-Z]{0,4}\s+([12][0-9OIL]{3})"
        ),
        # 31 de mayo de 2025, por si OCR pierde la 'a'
        re.compile(
            r"([0-3ZO]?[0-9OIL])\s+D[A-Z]{0,4}\s+"
            r"([A-Z]{4,12})\s+D[A-Z]{0,4}\s+([12][0-9OIL]{3})"
        ),
    ]

    for patron in patrones:
        for match in patron.finditer(encabezado):
            mes = corregir_mes_ocr(match.group(2))
            if not mes:
                continue
            try:
                return date(
                    corregir_numero_ocr(match.group(3)),
                    mes,
                    corregir_numero_ocr(match.group(1)),
                )
            except ValueError:
                continue

    # Fechas numericas por si algun encabezado viene como 31/05/2025.
    match = re.search(r"([0-3]?[0-9])[/\-]([01]?[0-9])[/\-]([12][0-9]{3})", encabezado)
    if match:
        try:
            return date(int(match.group(3)), int(match.group(2)), int(match.group(1)))
        except ValueError:
            return None
    return None


def limpiar_nombre_extraido(nombre: str) -> str:
    nombre = normalizar_nombre(nombre)
    palabras = nombre.split()
    cortado: list[str] = []
    cortes = {"RFC", "R", "F", "PUESTO", "ADSCRITO", "SIN", "ATENTAMENTE"}
    for palabra in palabras:
        if palabra in cortes:
            break
        cortado.append(palabra)
    return " ".join(cortado[:6])


def _lineas_limpias(texto: str) -> list[str]:
    return [linea.strip() for linea in texto.splitlines() if linea.strip()]


def _es_linea_nombre_probable(linea: str) -> bool:
    nombre = limpiar_nombre_extraido(linea)
    palabras = nombre.split()
    if len(palabras) < 3:
        return False
    prohibidas = {
        "DIRECTOR", "REGIONAL",
        "PRESENTE", "PUESTO", "RENUNCIA", "CARACTER", "IRREVOCABLE",
        "ENCUESTA", "COORDINACION", "ESTATAL", "VERACRUZ",
        "CONSIDERACION", "ATENTA", "DISTINGUIDA",
    }
    return not any(palabra in prohibidas for palabra in palabras)


def extraer_nombres_renuncia(texto_hoja: str) -> list[str]:
    nombres: list[str] = []
    lineas = _lineas_limpias(texto_hoja)

    # Caso: NOMBRE: PERSONA EJEMPLO
    for linea in lineas:
        match = re.search(r"\bNOMBRE\s*[:\-]?\s*(.+)$", linea, flags=re.IGNORECASE)
        if not match:
            continue
        candidato = limpiar_nombre_extraido(match.group(1))
        if candidato and candidato not in nombres:
            nombres.append(candidato)

    # Caso: nombre de la persona / Nombre y firma
    for indice, linea in enumerate(lineas):
        if "NOMBRE" not in normalizar_texto(linea) or "FIRMA" not in normalizar_texto(linea):
            continue
        for previa in reversed(lineas[max(0, indice - 5):indice]):
            candidato = limpiar_nombre_extraido(previa)
            if candidato and _es_linea_nombre_probable(candidato) and candidato not in nombres:
                nombres.append(candidato)
                break

    # Respaldo: buscar en las ultimas lineas, porque el nombre esta al final.
    for linea in reversed(lineas[-10:]):
        candidato = limpiar_nombre_extraido(linea)
        if candidato and _es_linea_nombre_probable(candidato) and candidato not in nombres:
            nombres.append(candidato)

    return nombres


def configurar_tesseract(tesseract_arg: str | None) -> bool:
    global TESSERACT_EXE

    candidatos: list[Path | str] = []
    if tesseract_arg:
        candidatos.append(Path(tesseract_arg))

    env_tesseract = os.environ.get("TESSERACT_CMD")
    if env_tesseract:
        candidatos.append(Path(env_tesseract))

    encontrado = shutil.which("tesseract") or shutil.which("tesseract.exe")
    if encontrado:
        candidatos.append(encontrado)

    candidatos.extend([
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    ])

    for candidato in candidatos:
        path_candidato = Path(candidato)
        if path_candidato.exists() or shutil.which(str(candidato)):
            TESSERACT_EXE = str(path_candidato if path_candidato.exists() else candidato)
            tessdata = path_candidato.parent / "tessdata" if path_candidato.exists() else None
            if tessdata and tessdata.exists():
                os.environ.setdefault("TESSDATA_PREFIX", str(tessdata))
            return True

    return False


def ocr_pagina(pagina: "fitz.Page") -> str:
    pix = pagina.get_pixmap(dpi=OCR_DPI, alpha=False)
    try:
        proceso = subprocess.run(
            [TESSERACT_EXE, "stdin", "stdout", "--psm", "6"],
            input=pix.tobytes("png"),
            capture_output=True,
            check=False,
            timeout=60,
        )
    except subprocess.TimeoutExpired:
        return ""
    return proceso.stdout.decode("utf-8", errors="ignore") if proceso.stdout else ""


def cargar_excel(path: Path) -> list[FilaExcel]:
    df = pd.read_excel(path)
    if df.shape[1] < 6:
        raise ValueError("El Excel debe tener al menos 6 columnas; se usan B, E y F.")

    col_nombre = df.columns[1]
    col_fecha = df.columns[4]
    col_nuevo_nombre = df.columns[5]

    filas: list[FilaExcel] = []
    for _, row in df.iterrows():
        nombre = str(row[col_nombre] or "").strip()
        fecha = pd.to_datetime(row[col_fecha], dayfirst=True, errors="coerce")
        nuevo_nombre = str(row[col_nuevo_nombre] or "").strip()
        if not nombre or pd.isna(fecha) or not nuevo_nombre:
            continue
        filas.append(
            FilaExcel(
                nombre=nombre,
                nombre_norm=normalizar_nombre(nombre),
                fecha_baja=fecha.date(),
                nuevo_nombre=nuevo_nombre,
            )
        )
    return filas


def similitud_nombre(nombre_pdf: str, nombre_excel: str) -> float:
    nombre_pdf_norm = normalizar_nombre(nombre_pdf)
    nombre_excel_norm = normalizar_nombre(nombre_excel)

    if not nombre_pdf_norm or not nombre_excel_norm:
        return 0.0

    if nombre_pdf_norm == nombre_excel_norm:
        return 1.0

    palabras_pdf = set(nombre_pdf_norm.split())
    palabras_excel = set(nombre_excel_norm.split())
    coincidencia_palabras = len(palabras_pdf & palabras_excel) / max(
        len(palabras_pdf),
        len(palabras_excel),
        1,
    )
    secuencia = SequenceMatcher(None, nombre_pdf_norm, nombre_excel_norm).ratio()
    return max(coincidencia_palabras, secuencia)


def buscar_coincidencia(
    filas_excel: list[FilaExcel],
    fecha_pdf: date | None,
    nombres_pdf: list[str],
) -> tuple[FilaExcel | None, float, str]:
    if not fecha_pdf:
        return None, 0.0, ""

    candidatas = [fila for fila in filas_excel if fila.fecha_baja == fecha_pdf]
    mejor_fila: FilaExcel | None = None
    mejor_score = 0.0
    mejor_nombre_pdf = ""
    for fila in candidatas:
        for nombre_pdf in nombres_pdf:
            score = similitud_nombre(nombre_pdf, fila.nombre)
            if score > mejor_score:
                mejor_fila = fila
                mejor_score = score
                mejor_nombre_pdf = nombre_pdf

    if mejor_fila and mejor_score >= UMBRAL_COINCIDENCIA_NOMBRE:
        return mejor_fila, mejor_score, mejor_nombre_pdf
    return None, mejor_score, mejor_nombre_pdf


def nombre_pdf_destino(valor_columna_f: str) -> str:
    nombre = Path(valor_columna_f).name.strip()
    if not nombre.lower().endswith(".pdf"):
        nombre = f"{nombre}.pdf"
    return nombre


def encontrar_excel(carpeta: Path, excel_arg: str | None) -> Path:
    if excel_arg:
        excel_path = Path(excel_arg)
        if not excel_path.exists():
            raise FileNotFoundError(
                f"No encontre el Excel indicado: {excel_path}. "
                "Revisa la ruta o usa --excel con un archivo .xlsx existente."
            )
        if not excel_path.is_file():
            raise ValueError(f"La ruta indicada en --excel no es un archivo: {excel_path}")
        return excel_path

    if not carpeta.exists():
        raise FileNotFoundError(f"No encontre la carpeta indicada: {carpeta}")
    if not carpeta.is_dir():
        raise ValueError(f"La ruta indicada como carpeta no es una carpeta: {carpeta}")

    excels = sorted(
        path for path in carpeta.glob("*.xlsx")
        if not path.name.startswith("~$")
        and not path.stem.lower().startswith("reporte_")
    )
    if len(excels) == 1:
        return excels[0]
    if not excels:
        raise FileNotFoundError("No encontre ningun archivo .xlsx en la carpeta.")
    raise ValueError("Hay varios Excel. Indica cual usar con --excel.")


def procesar_pdf(pdf_path: Path, filas_excel: list[FilaExcel], dry_run: bool) -> dict:
    resultado = {
        "archivo_original": pdf_path.name,
        "fecha_pdf": "",
        "nombre_pdf": "",
        "nombre_excel": "",
        "nuevo_nombre": "",
        "score_nombre": 0.0,
        "estado": "",
    }

    try:
        with fitz.open(pdf_path) as doc:
            if doc.page_count < 1:
                resultado["estado"] = "sin paginas"
                return resultado
            texto_hoja = ocr_pagina(doc[0])
    except Exception as e:  # noqa: BLE001 - se reporta en el resumen
        resultado["estado"] = f"error leyendo PDF: {e}"
        return resultado

    if not es_renuncia(texto_hoja):
        resultado["estado"] = "omitido: no es renuncia"
        return resultado

    fecha_pdf = extraer_fecha_renuncia(texto_hoja)
    nombres_pdf = extraer_nombres_renuncia(texto_hoja)
    fila, score, nombre_usado = buscar_coincidencia(filas_excel, fecha_pdf, nombres_pdf)

    resultado["fecha_pdf"] = fecha_pdf.strftime("%d/%m/%Y") if fecha_pdf else ""
    resultado["nombre_pdf"] = nombre_usado or " | ".join(nombres_pdf)
    resultado["score_nombre"] = round(score, 3)

    if not fecha_pdf:
        resultado["estado"] = "no se pudo extraer fecha"
        return resultado
    if not nombres_pdf:
        resultado["estado"] = "no se pudo extraer nombre"
        return resultado
    if not fila:
        resultado["estado"] = "sin coincidencia en Excel"
        return resultado

    destino = pdf_path.with_name(nombre_pdf_destino(fila.nuevo_nombre))
    resultado["nombre_excel"] = fila.nombre
    resultado["nuevo_nombre"] = destino.name

    if destino.exists() and destino.resolve() != pdf_path.resolve():
        resultado["estado"] = "destino ya existe"
        return resultado

    if dry_run:
        resultado["estado"] = "coincide (simulacion)"
        return resultado

    pdf_path.rename(destino)
    resultado["estado"] = "renombrado"
    return resultado


def parsear_argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Renombra PDFs de renuncias usando fecha/nombre extraidos por OCR y un Excel."
    )
    parser.add_argument(
        "carpeta",
        nargs="?",
        default=".",
        help="Carpeta con PDFs y Excel. Por defecto: carpeta actual.",
    )
    parser.add_argument(
        "--excel",
        help="Ruta del Excel. Si no se indica, usa el unico .xlsx de la carpeta.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="No renombra archivos; solo genera/reporta lo que haria.",
    )
    parser.add_argument(
        "--reporte",
        default=str(REPORTE_DEFAULT),
        help=f"Excel de reporte. Por defecto: {REPORTE_DEFAULT}",
    )
    parser.add_argument(
        "--tesseract",
        help="Ruta al programa tesseract si no esta en el PATH.",
    )
    return parser.parse_args()


def main() -> None:
    args = parsear_argumentos()
    if not configurar_tesseract(args.tesseract):
        logger.error(
            "No encontre tesseract. Instala Tesseract OCR, usa --tesseract, "
            "o agrega tesseract al PATH."
        )
        sys.exit(1)

    carpeta = Path(args.carpeta)
    try:
        excel_path = encontrar_excel(carpeta, args.excel)
    except (FileNotFoundError, ValueError) as e:
        logger.error("%s", e)
        sys.exit(1)
    reporte_path = Path(args.reporte)

    try:
        filas_excel = cargar_excel(excel_path)
    except Exception as e:  # noqa: BLE001 - se muestra como error de entrada
        logger.error("No pude leer el Excel '%s': %s", excel_path, e)
        sys.exit(1)
    pdfs = sorted(carpeta.glob("*.pdf"))
    if not pdfs:
        logger.warning("No se encontraron PDFs en '%s'.", carpeta)
        return

    logger.info("Excel: %s", excel_path)
    logger.info("PDFs encontrados: %s", len(pdfs))

    resultados = [
        procesar_pdf(pdf_path, filas_excel, dry_run=args.dry_run)
        for pdf_path in tqdm(pdfs, desc="Procesando renuncias")
    ]

    df_reporte = pd.DataFrame(resultados)
    df_reporte.to_excel(reporte_path, index=False)

    renombrados = sum(1 for item in resultados if item["estado"] == "renombrado")
    simulados = sum(1 for item in resultados if item["estado"] == "coincide (simulacion)")
    omitidos = sum(1 for item in resultados if str(item["estado"]).startswith("omitido"))
    logger.info(
        "Listo. Renombrados: %s. Coincidencias en simulacion: %s. Omitidos: %s. Reporte: %s",
        renombrados,
        simulados,
        omitidos,
        reporte_path,
    )


if __name__ == "__main__":
    main()
