# renombrador-renuncias-ocr

Herramienta en Python para renombrar lotes de PDF de renuncias usando OCR y una hoja Excel de control. El problema que resuelve es operativo: cuando llegan decenas o cientos de renuncias escaneadas con nombres genericos, revisar cada archivo manualmente para asignarle la clave correcta consume mucho tiempo y aumenta el riesgo de errores.

El script extrae texto de la primera pagina del PDF, confirma que el documento sea una renuncia, identifica fecha, nombre y, en la version mejorada, RFC parcial. Luego compara esos datos contra un Excel y renombra el PDF con el nombre indicado en la columna de destino. Tambien genera un reporte para revisar coincidencias, omisiones y casos ambiguos.

## Estado del proyecto

Este repositorio esta preparado para publicarse sin documentos reales. Los PDF, Excel de trabajo, reportes y archivos comprimidos estan excluidos por `.gitignore` porque pueden contener datos personales.

## Scripts incluidos

- `renombrar_renuncias.py`: version recomendada. Mejora la comparacion de nombres, maneja nombres en orden invertido, detecta ambiguedades y usa RFC parcial solo como apoyo sin saltarse la validacion de fecha.
- `renombrar_renuncias_basico.py`: version anterior y mas simple, conservada como referencia funcional.

## Requisitos

- Python 3.10 o superior.
- Tesseract OCR instalado y disponible en el `PATH`.
- Dependencias Python listadas en `requirements.txt`.

Instalacion:

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

En Windows:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

Si Tesseract no esta en el `PATH`, define `TESSERACT_CMD` o usa `--tesseract`.

## Formato esperado del Excel

El archivo Excel debe tener al menos seis columnas. Los scripts usan estas posiciones:

- Columna A: clave o RFC completo/parcial, usada por `renombrar_renuncias.py`.
- Columna B: nombre de la persona.
- Columna E: fecha de baja.
- Columna F: nombre final que debe recibir el PDF.

Los reportes generados se excluyen del repositorio por defecto.

## Uso

Reemplaza `/ruta/a/carpeta` y `/ruta/a/archivo.xlsx` por tus rutas reales.

Si los PDF y el Excel estan en la carpeta actual:

```bash
python3 renombrar_renuncias.py . --excel "archivo_control.xlsx" --dry-run
```

Simular el proceso sin renombrar archivos:

```bash
python3 renombrar_renuncias.py /ruta/a/carpeta --excel "/ruta/a/archivo.xlsx" --dry-run
```

Renombrar archivos:

```bash
python3 renombrar_renuncias.py /ruta/a/carpeta --excel "/ruta/a/archivo.xlsx"
```

Indicar Tesseract manualmente:

```bash
python3 renombrar_renuncias.py /ruta/a/carpeta --excel "/ruta/a/archivo.xlsx" --tesseract "/ruta/a/tesseract"
```

Cambiar el nombre del reporte:

```bash
python3 renombrar_renuncias.py /ruta/a/carpeta --excel "/ruta/a/archivo.xlsx" --reporte reporte_revision.xlsx
```

## Flujo recomendado

1. Coloca los PDF y el Excel en una carpeta local fuera del control de versiones.
2. Ejecuta primero con `--dry-run`.
3. Revisa el reporte generado.
4. Si las coincidencias son correctas, ejecuta sin `--dry-run`.
5. Revisa manualmente cualquier estado `ambiguo: revisar manualmente`.

## Privacidad

No publiques documentos de renuncia reales, hojas Excel de control, reportes generados ni archivos comprimidos de trabajo. Este proyecto esta pensado para publicar el codigo y las instrucciones, no los datos procesados.

Antes de crear el primer commit, verifica:

```bash
git status --short
```

Solo deberian agregarse archivos fuente y documentacion.

## Limitaciones

- La calidad del resultado depende del escaneo y del OCR.
- Los documentos con mala resolucion pueden requerir revision manual.
- El script no modifica el contenido de los PDF; solo decide si debe renombrarlos.
- Los casos ambiguos se bloquean intencionalmente para evitar renombrados incorrectos.
