# AGENTS.md

Greenfield: no app code yet. Stack decided: Python + Flask in `app.py` (not yet installed — see Environment). Repo is `main`, `README.md` is a one-line stub.

## Layout

- `Insumos/` — read-only source of truth. Never edit in place; write outputs elsewhere.
  - `Maestro_Estudiantes.xlsx` — sheet `Estudiantes`, `A1:E25` + header, frozen header row. Columns: `Identificacion | Nombre_Completo | Correo | Programa | Cohorte`. 24 students, all `Cohorte=2026-A`. Emails are `@ejemplo.cr` (synthetic).
  - `Registro_Evaluaciones.xlsx` — sheet `Evaluaciones`, `A1:F89` + header, frozen header row. Columns: `Identificacion | Programa | Modulo | Nota | Asistencia_Pct | Fecha_Cierre`. 88 data rows. `Nota`/`Asistencia_Pct` are numeric cells (`t="n"`); everything else is `inlineStr` (no `sharedStrings.xml` in either file). Dates are `YYYY-MM-DD` text.

## Data gotchas (verified from sheet XML)

- Join key is `Identificacion`: keep as string, never cast to int.
- Module count differs by program: `Técnico en IA Aplicada` has Módulo 1–4, `Excel Avanzado para Negocios` has Módulo 1–3. Do not assume 4 rows per student.
- Known mismatches (handle explicitly, don't silently drop):
  - `304560321` (Espinoza León, Técnico en IA Aplicada in Maestro) has 0 evaluation rows.
  - `999880777` has 4 IA evaluation rows but no Maestro entry (orphan).
- `Programa` is in both files; on conflict prefer Maestro and flag it.

## Arquitectura (Academia Horizonte — app web de certificados)

- Bodega (datos): `Insumos/` (nota: el usuario la llama `insumos/` en minúsculas). NO se modifica.
- Cocina (backend): Python con Flask en `app.py`. Toda la lógica va aquí, NUNCA en la interfaz.
- Salón (frontend): páginas HTML servidas por Flask. Solo muestran información; no calculan ni deciden.
- Librerías: `openpyxl` para leer Excel (simple, deja ver bucles y condicionales; se descarta pandas por pesada y por esconder la lógica), `Flask` para el servidor (+ Jinja, que viene con Flask, solo para mostrar). Ver detalle y orden en `PLAN.md`.
- Estructura: `app.py`, `requirements.txt` (`Flask`, `openpyxl`), `templates/base.html`, `templates/index.html`, `templates/estudiante.html`, `static/style.css`. Rutas mínimas: `/` y `/estudiante/<identificacion>`.
- Huérfano `999880777`: se marca "revisar", no genera certificado oficial (el Maestro manda).

## Reglas de negocio

- Agrupar por `Identificacion` + `Programa`. Un certificado por estudiante y programa.
- `Promedio` = suma de `Nota` / cantidad de módulos cursados. `Asistencia` = promedio de `Asistencia_Pct`.
- Aprobación si `Promedio >= 70` y `Asistencia >= 80`. Participación si `Promedio < 70` y `Asistencia >= 80`. Sin certificado si `Asistencia < 80`. Los límites INCLUYEN el valor.

## Estilo de trabajo

- Código simple y comentado en español, señalando los ladrillos: variables, tipos, condicionales, bucles y funciones.
- Diseño web en azul marino y dorado.
- Explicar en español sencillo lo que se vaya haciendo.

## Environment

- Windows + PowerShell (`shell` tool runs PowerShell, not cmd/bash).
- No working Python/Node runtimes (`python` is a Microsoft Store stub, `node` missing); only `git` works. Stack decided (Python + Flask) but not installed — install it before building `app.py`.
- PowerShell: use `Get-ChildItem -Force` with the `workdir` param instead of `cd` / `dir /a` (breaks on paths with spaces).
- To inspect sheets as XML: copy `.xlsx` to `*.zip` first (`Expand-Archive` rejects `.xlsx`), then `Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::ExtractToDirectory(src, dst)`.
