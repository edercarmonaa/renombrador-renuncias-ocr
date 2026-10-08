"""
Renombra PDFs de renuncias con mejora para nombres en orden invertido.

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

Reglas de precision (version mejorada):
    - La fecha SIEMPRE debe coincidir exactamente con el Excel, incluso si
      hay un RFC que coincide. No se usa el RFC para saltarse la fecha.
    - El score de nombre combina "palabras en comun" y "similitud de
      caracteres" con peso, en vez de tomar el maximo de los dos. Si solo
      comparten 0 o 1 palabra, el score se limita para evitar falsos
      positivos por apellidos comunes o parecido superficial.
    - Si dos candidatos del mismo dia quedan con score muy parecido, el PDF
      se marca como "ambiguo" y NO se renombra; el reporte incluye ambas
      opciones para revision manual.
    - Fechas con anio fuera de un rango razonable se descartan (probable
      error de OCR) en vez de usarse para intentar un match.

Requisitos:
    python3 -m pip install pymupdf pandas openpyxl tqdm
    Tener tesseract instalado en el PATH o indicar su ruta con --tesseract.

Uso:
    python3 renombrar_renuncias.py
    python3 renombrar_renuncias.py /ruta/a/carpeta --excel "Archivo.xlsx"
    python3 renombrar_renuncias.py --dry-run
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
from dataclasses import dataclass, field
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
MARGEN_AMBIGUEDAD = 0.06  # si el 1er y 2do lugar quedan a menos de esto, es ambiguo
ANIO_MINIMO = 2000
REPORTE_DEFAULT = Path("reporte_renombrado_renuncias.xlsx")
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
    clave: str
    clave10: str
    nombre: str
    nombre_norm: str
    fecha_baja: date
    nuevo_nombre: str


@dataclass
class ResultadoCoincidencia:
    fila: FilaExcel | None
    score: float
    nombre_usado: str
    ambiguo: bool = False
    alternativas: list[str] = field(default_factory=list)


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


def normalizar_clave(valor: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(valor or "").upper())


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


def _fecha_valida(fecha: date) -> bool:
    """Descarta fechas fuera de un rango razonable (probable error de OCR)."""
    hoy = date.today()
    return date(ANIO_MINIMO, 1, 1) <= fecha <= date(hoy.year + 1, 12, 31)


def extraer_fecha_renuncia(texto_hoja: str) -> date | None:
    lineas = _lineas_limpias(texto_hoja)
    encabezado = normalizar_texto(" ".join(lineas[:4]))

    patrones = [
        # Xalapa, Ver., a 31 de mayo de 2025
        re.compile(
            r"(?:^|\s)A\s*([0-3ZO]?[0-9OIL])\s+D[A-Z]{0,4}\s+"
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
                fecha = date(
                    corregir_numero_ocr(match.group(3)),
                    mes,
                    corregir_numero_ocr(match.group(1)),
                )
            except ValueError:
                continue
            if _fecha_valida(fecha):
                return fecha

    # Caso OCR pegado: a31demayode 2025.
    compacto = re.sub(r"\s+", "", encabezado)
    match = re.search(
        r"([0-3ZO]?[0-9OIL])D[A-Z]{0,4}([A-Z]{4,12})D[A-Z]{0,4}([12][0-9OIL]{3})",
        compacto,
    )
    if match:
        mes = corregir_mes_ocr(match.group(2))
        if mes:
            try:
                fecha = date(
                    corregir_numero_ocr(match.group(3)),
                    mes,
                    corregir_numero_ocr(match.group(1)),
                )
                if _fecha_valida(fecha):
                    return fecha
            except ValueError:
                pass

    # Fechas numericas por si algun encabezado viene como 31/05/2025.
    match = re.search(r"([0-3]?[0-9])[/\-]([01]?[0-9])[/\-]([12][0-9]{3})", encabezado)
    if match:
        try:
            fecha = date(int(match.group(3)), int(match.group(2)), int(match.group(1)))
            if _fecha_valida(fecha):
                return fecha
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


def _es_linea_separador_firma(linea: str) -> bool:
    linea_raw = str(linea or "").strip()
    linea_norm = normalizar_texto(linea_raw)
    if not linea_raw or not linea_norm:
        return True
    if linea_norm in {"FIRMA", "ATENTAMENTE", "NOMBRE FIRMA", "NOMBRE Y FIRMA"}:
        return True

    # Lineas horizontales, trazos o ruido de firma que quedan entre el nombre
    # y la etiqueta "Nombre y firma".
    sin_espacios = re.sub(r"\s+", "", linea_raw)
    if len(sin_espacios) >= 3 and re.fullmatch(r"[-_=~.·—_]+", sin_espacios):
        return True
    letras = re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]", linea_raw)
    return len(linea_raw) >= 4 and len(letras) <= 1


def _es_linea_nombre_probable(linea: str) -> bool:
    if _es_linea_separador_firma(linea):
        return False

    nombre = limpiar_nombre_extraido(linea)
    palabras = nombre.split()
    if len(palabras) < 2:
        return False
    if len(palabras) == 2 and max(len(palabra) for palabra in palabras) < 7:
        return False
    prohibidas = {
        "DIRECTOR", "REGIONAL",
        "PRESENTE", "PUESTO", "RENUNCIA", "CARACTER", "IRREVOCABLE",
        "ENCUESTA", "COORDINACION", "ESTATAL", "VERACRUZ",
        "CONSIDERACION", "ATENTA", "DISTINGUIDA", "FIRMA", "CLAVE",
        "PLAZA", "EVENTUALES", "PROGRAMA", "DESEMPEFIANDO", "DESEMPENANDO",
        "SEGURIDAD", "PUBLICA", "JUSTICIA", "INFORMACION", "GOBIERNOS",
        "MUNICIPALES", "DEMARCACIONES", "TERRITORIALES", "CIUDAD", "MEXICO",
        "CNGMD",
    }
    return not any(palabra in prohibidas for palabra in palabras)


def _buscar_nombre_encima_de_firma(lineas: list[str], indice_etiqueta: int) -> str:
    for previa in reversed(lineas[max(0, indice_etiqueta - 10):indice_etiqueta]):
        if _es_linea_separador_firma(previa):
            continue
        candidato = limpiar_nombre_extraido(previa)
        if candidato and _es_linea_nombre_probable(candidato):
            return candidato
    return ""


def normalizar_rfc10(valor: object) -> str:
    clave = normalizar_clave(valor)
    if len(clave) < 10:
        return ""

    letras = clave[:4].replace("0", "O").replace("1", "I")
    fecha = (
        clave[4:10]
        .replace("O", "0")
        .replace("I", "1")
        .replace("L", "1")
        .replace("Z", "2")
    )
    if not re.fullmatch(r"[A-Z&]{3,4}", letras):
        return ""
    if not re.fullmatch(r"[0-9]{6}", fecha):
        return ""
    return f"{letras}{fecha}"


def extraer_rfcs(texto_hoja: str) -> list[str]:
    candidatos: list[str] = []
    lineas = _lineas_limpias(texto_hoja)
    texto_norm = normalizar_texto(texto_hoja)

    def agregar(valor: object) -> None:
        rfc10 = normalizar_rfc10(valor)
        if rfc10 and rfc10 not in candidatos:
            candidatos.append(rfc10)

    for match in re.finditer(r"\b[A-ZÑ&]{3,4}[0-9OILZ]{6}[A-Z0-9]{0,3}\b", texto_norm):
        agregar(match.group(0))

    for indice, linea in enumerate(lineas):
        linea_norm = normalizar_texto(linea)
        if "RFC" not in linea_norm and "R F C" not in linea_norm:
            continue
        bloque = linea_norm
        if indice + 1 < len(lineas):
            bloque = f"{bloque} {normalizar_texto(lineas[indice + 1])}"
        bloque = re.sub(r".*?R\s*F\s*C", "", bloque)
        for match in re.finditer(r"[A-ZÑ&0-9OILZ]{10,13}", bloque):
            agregar(match.group(0))

    return candidatos


def extraer_nombres_renuncia(texto_hoja: str) -> list[str]:
    nombres: list[str] = []
    lineas = _lineas_limpias(texto_hoja)

    # Caso: NOMBRE: PERSONA EJEMPLO
    for linea in lineas:
        linea_norm = normalizar_texto(linea)
        if "FIRMA" in linea_norm:
            continue
        match = re.search(r"\bNOMBRE\s*[:\-]\s*(.+)$", linea, flags=re.IGNORECASE)
        if not match:
            continue
        candidato = limpiar_nombre_extraido(match.group(1))
        if candidato and _es_linea_nombre_probable(candidato) and candidato not in nombres:
            nombres.append(candidato)

    # Caso: nombre de la persona / linea horizontal / Nombre y firma
    for indice, linea in enumerate(lineas):
        linea_norm = normalizar_texto(linea)
        if "NOMBRE" not in linea_norm or "FIRMA" not in linea_norm:
            continue
        candidato = _buscar_nombre_encima_de_firma(lineas, indice)
        if candidato and candidato not in nombres:
            nombres.append(candidato)

    # Caso mas confiable: el nombre del trabajador esta justo debajo de
    # "Atentamente", normalmente en la penultima linea del documento.
    for indice, linea in enumerate(lineas):
        if "ATENTAMENTE" not in normalizar_texto(linea):
            continue
        for candidata_linea in lineas[indice + 1:indice + 8]:
            if _es_linea_separador_firma(candidata_linea):
                continue
            if "NOMBRE" in normalizar_texto(candidata_linea) and "FIRMA" in normalizar_texto(candidata_linea):
                continue
            candidato = limpiar_nombre_extraido(candidata_linea)
            if candidato and _es_linea_nombre_probable(candidato):
                if candidato in nombres:
                    nombres.remove(candidato)
                nombres.insert(0, candidato)
                break
        break

    # Respaldo directo: buscar desde el final, saltando lineas horizontales,
    # etiquetas y trazos de firma.
    if not nombres:
        for linea in reversed(lineas[-10:]):
            if "NOMBRE" in normalizar_texto(linea) and "FIRMA" in normalizar_texto(linea):
                continue
            if _es_linea_separador_firma(linea):
                continue
            candidato = limpiar_nombre_extraido(linea)
            if candidato and _es_linea_nombre_probable(candidato):
                nombres.insert(0, candidato)
                break

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

    col_clave = df.columns[0]
    col_nombre = df.columns[1]
    col_fecha = df.columns[4]
    col_nuevo_nombre = df.columns[5]

    filas: list[FilaExcel] = []
    for _, row in df.iterrows():
        clave = normalizar_clave(row[col_clave])
        nombre = str(row[col_nombre] or "").strip()
        fecha = pd.to_datetime(row[col_fecha], dayfirst=True, errors="coerce")
        nuevo_nombre = str(row[col_nuevo_nombre] or "").strip()
        if not nombre or pd.isna(fecha) or not nuevo_nombre:
            continue
        filas.append(
            FilaExcel(
                clave=clave,
                clave10=normalizar_rfc10(clave),
                nombre=nombre,
                nombre_norm=normalizar_nombre(nombre),
                fecha_baja=fecha.date(),
                nuevo_nombre=nuevo_nombre,
            )
        )
    return filas


def _tokens_nombre(nombre: str) -> list[str]:
    return normalizar_nombre(nombre).split()


def _rotaciones(tokens: list[str]) -> list[str]:
    variantes = []
    for indice in range(len(tokens)):
        variante = " ".join(tokens[indice:] + tokens[:indice])
        if variante and variante not in variantes:
            variantes.append(variante)
    return variantes


def variantes_nombre(nombre: str) -> list[str]:
    tokens = _tokens_nombre(nombre)
    variantes = []
    base = " ".join(tokens)
    if base:
        variantes.append(base)

    # Para casos donde el PDF trae APELLIDOS + NOMBRE(S) y el Excel trae
    # NOMBRE(S) + APELLIDOS, o viceversa.
    for variante in _rotaciones(tokens):
        if variante not in variantes:
            variantes.append(variante)

    ordenado = " ".join(sorted(tokens))
    if ordenado and ordenado not in variantes:
        variantes.append(ordenado)

    return variantes


def variantes_apellidos_primero(nombre_excel: str) -> list[str]:
    tokens = _tokens_nombre(nombre_excel)
    variantes = []
    if len(tokens) >= 4:
        # NOMBRE1 NOMBRE2 APELLIDO1 APELLIDO2 -> APELLIDO1 APELLIDO2 NOMBRE1 NOMBRE2
        variantes.append(" ".join(tokens[-2:] + tokens[:-2]))
    if len(tokens) >= 3:
        # NOMBRE APELLIDO1 APELLIDO2 -> APELLIDO1 APELLIDO2 NOMBRE
        variantes.append(" ".join(tokens[-2:] + tokens[:-2]))
        # NOMBRE1 NOMBRE2 APELLIDO -> APELLIDO NOMBRE1 NOMBRE2
        variantes.append(" ".join(tokens[-1:] + tokens[:-1]))
    return [variante for variante in variantes if variante]


def similitud_nombre(nombre_pdf: str, nombre_excel: str) -> float:
    """Combina palabras en comun y similitud de caracteres con peso.

    Antes se usaba max(coincidencia_palabras, secuencia), lo que permitia que
    un parecido puramente de caracteres (sin compartir palabras reales)
    disparara un falso positivo. Ahora se combinan ambas metricas y, si solo
    comparten 0 o 1 palabra, se limita el score para exigir mas evidencia
    antes de aceptar la coincidencia.
    """
    variantes_pdf = variantes_nombre(nombre_pdf)
    variantes_excel = variantes_nombre(nombre_excel)

    if not variantes_pdf or not variantes_excel:
        return 0.0

    mejor = 0.0
    for pdf_variante in variantes_pdf:
        for excel_variante in variantes_excel:
            if pdf_variante == excel_variante:
                return 1.0
            tokens_pdf = set(pdf_variante.split())
            tokens_excel = set(excel_variante.split())
            comunes = tokens_pdf & tokens_excel
            coincidencia_palabras = len(comunes) / max(
                len(tokens_pdf),
                len(tokens_excel),
                1,
            )
            secuencia = SequenceMatcher(None, pdf_variante, excel_variante).ratio()
            combinado = (coincidencia_palabras * 0.65) + (secuencia * 0.35)

            if len(comunes) < 2:
                # Con 0 o 1 palabra en comun no hay suficiente evidencia;
                # se limita el score para que no cruce el umbral de aceptacion.
                combinado = min(combinado, 0.5)

            mejor = max(mejor, combinado)
    return mejor


def score_nombre_en_texto_ocr(texto_ocr: str, nombre_excel: str) -> tuple[float, str]:
    texto_norm = normalizar_texto(texto_ocr)
    variantes = [*variantes_nombre(nombre_excel), *variantes_apellidos_primero(nombre_excel)]
    for variante in variantes:
        if variante and variante in texto_norm:
            return 1.0, variante
    return 0.0, ""


def buscar_coincidencia(
    filas_excel: list[FilaExcel],
    fecha_pdf: date | None,
    nombres_pdf: list[str],
    texto_ocr: str = "",
    rfcs_pdf: list[str] | None = None,
) -> ResultadoCoincidencia:
    """Busca la fila del Excel que corresponde a un PDF.

    Reglas:
      - La fecha SIEMPRE debe coincidir (RFC no la reemplaza).
      - Si dos o mas candidatas del mismo dia quedan con score muy cercano
        (dentro de MARGEN_AMBIGUEDAD) por nombre, se marca como ambiguo y no
        se elige ninguna, para evitar un renombrado incorrecto.
      - Un match exacto por RFC se considera no ambiguo (es un identificador
        unico), salvo que otra fila del mismo dia tambien tenga ese RFC.
    """
    if not fecha_pdf:
        return ResultadoCoincidencia(None, 0.0, "")

    rfcs_pdf = rfcs_pdf or []
    candidatas = [fila for fila in filas_excel if fila.fecha_baja == fecha_pdf]
    if not candidatas:
        return ResultadoCoincidencia(None, 0.0, "")

    # (score, fila, nombre_usado, es_rfc)
    puntuadas: list[tuple[float, FilaExcel, str, bool]] = []

    for fila in candidatas:
        if fila.clave10 and fila.clave10 in rfcs_pdf:
            puntuadas.append((1.0, fila, f"RFC:{fila.clave10}", True))
            continue

        mejor_local = 0.0
        nombre_local = ""
        for nombre_pdf in nombres_pdf:
            score = similitud_nombre(nombre_pdf, fila.nombre)
            if score > mejor_local:
                mejor_local, nombre_local = score, nombre_pdf

        score_ocr, nombre_ocr = score_nombre_en_texto_ocr(texto_ocr, fila.nombre)
        if score_ocr > mejor_local:
            mejor_local, nombre_local = score_ocr, nombre_ocr

        puntuadas.append((mejor_local, fila, nombre_local, False))

    puntuadas.sort(key=lambda item: item[0], reverse=True)
    mejor_score, mejor_fila, mejor_nombre_pdf, es_rfc = puntuadas[0]

    if mejor_score < UMBRAL_COINCIDENCIA_NOMBRE:
        return ResultadoCoincidencia(None, mejor_score, mejor_nombre_pdf)

    if not es_rfc and len(puntuadas) > 1:
        segundo_score = puntuadas[1][0]
        if (mejor_score - segundo_score) < MARGEN_AMBIGUEDAD and segundo_score >= (
            UMBRAL_COINCIDENCIA_NOMBRE - 0.05
        ):
            alternativas = [
                f"{fila.nombre} -> {fila.nuevo_nombre} (score {score:.3f})"
                for score, fila, _, _ in puntuadas[:3]
                if score >= (UMBRAL_COINCIDENCIA_NOMBRE - 0.05)
            ]
            return ResultadoCoincidencia(
                None, mejor_score, mejor_nombre_pdf, ambiguo=True, alternativas=alternativas
            )

    return ResultadoCoincidencia(mejor_fila, mejor_score, mejor_nombre_pdf)


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
        "rfc_pdf": "",
        "clave_excel": "",
        "alternativas_ambiguas": "",
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
    rfcs_pdf = extraer_rfcs(texto_hoja)
    match = buscar_coincidencia(filas_excel, fecha_pdf, nombres_pdf, texto_hoja, rfcs_pdf)

    resultado["fecha_pdf"] = fecha_pdf.strftime("%d/%m/%Y") if fecha_pdf else ""
    resultado["nombre_pdf"] = match.nombre_usado or " | ".join(nombres_pdf)
    resultado["score_nombre"] = round(match.score, 3)
    resultado["rfc_pdf"] = " | ".join(rfcs_pdf)

    if not fecha_pdf:
        resultado["estado"] = "no se pudo extraer fecha"
        return resultado
    if not nombres_pdf and not rfcs_pdf:
        resultado["estado"] = "no se pudo extraer nombre ni RFC"
        return resultado

    if match.ambiguo:
        resultado["alternativas_ambiguas"] = " || ".join(match.alternativas)
        resultado["estado"] = "ambiguo: revisar manualmente"
        return resultado

    if not match.fila:
        resultado["estado"] = "sin coincidencia en Excel"
        return resultado

    destino = pdf_path.with_name(nombre_pdf_destino(match.fila.nuevo_nombre))
    resultado["nombre_excel"] = match.fila.nombre
    resultado["clave_excel"] = match.fila.clave
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
        description="Renombra PDFs de renuncias, incluyendo nombres en orden invertido, usando OCR y un Excel."
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
    ambiguos = sum(1 for item in resultados if str(item["estado"]).startswith("ambiguo"))
    logger.info(
        "Listo. Renombrados: %s. Coincidencias en simulacion: %s. Omitidos: %s. "
        "Ambiguos (revisar manual): %s. Reporte: %s",
        renombrados,
        simulados,
        omitidos,
        ambiguos,
        reporte_path,
    )


if __name__ == "__main__":
    main()
