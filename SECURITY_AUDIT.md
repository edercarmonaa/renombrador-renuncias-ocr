# Auditoria de seguridad y publicacion

Fecha de revision: 2026-10-08

## Alcance revisado

- Scripts Python: `renombrar_renuncias.py`, `renombrar_renuncias_basico.py`.
- Archivos ocultos y metadatos locales.
- Paquetes comprimidos, reportes generados y hojas de calculo.
- PDFs de renuncias presentes en la carpeta de trabajo.
- Estado Git local.

## Hallazgos

### ALTO - Documentos con datos personales

La carpeta contiene PDFs de renuncias y hojas Excel de trabajo. Por la naturaleza del proyecto, estos archivos pueden incluir nombres, claves, RFC u otros datos personales. No deben publicarse en GitHub.

Accion aplicada:

- Se agregaron patrones para excluir `*.pdf`, `*.xlsx`, `*.xls`, `*.csv` y reportes generados en `.gitignore`.
- Los documentos locales no se copiaron al README ni a ejemplos.

### MEDIO - Archivos generados y comprimidos

Se detectaron reportes `.xlsx`, un archivo `.zip`, carpeta `hechos/`, `__pycache__/` y metadatos `.DS_Store`. Estos archivos son salidas locales o empaquetados regenerables.

Accion aplicada:

- Se eliminaron cachés Python y `.DS_Store`.
- Se agregaron reglas de `.gitignore` para salidas, comprimidos, cachés y temporales.

### INFORMATIVO - Secretos y credenciales

No se encontraron llaves API, tokens, claves privadas, cadenas de conexion ni credenciales hardcodeadas en los archivos de texto revisados.

Accion aplicada:

- Se creo `.env.example` con la unica configuracion opcional usada por los scripts: `TESSERACT_CMD`.
- Se agregaron reglas preventivas para credenciales y certificados.

### INFORMATIVO - Historial Git

La carpeta revisada no es un repositorio Git. No existe historial local que auditar con `git log`, `git diff` o `git ls-files`.

Accion pendiente antes de publicar:

- Inicializar Git despues de revisar que solo se agreguen archivos publicables.
- Si estos archivos provienen de otro repositorio con historial previo, auditar ese historial antes de hacerlo publico.

## Recomendacion antes del primer commit

Ejecutar:

```bash
git init
git status --short
git add .gitignore .env.example requirements.txt README.md SECURITY_AUDIT.md renombrar_renuncias.py renombrar_renuncias_basico.py
git status --short
```

No agregar PDFs, hojas Excel reales, reportes, zips ni carpetas de salida.
