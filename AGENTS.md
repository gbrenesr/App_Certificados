# AGENTS.md

Greenfield workspace: no code, no README, no manifests, no git repo yet. Only data inputs exist.

## Layout

- `Insumos/` — read-only source of truth. Never edit in place; write outputs elsewhere.
  - `Maestro_Estudiantes.xlsx` — sheet `Estudiantes`, range `A1:E25` + header, frozen header row. Columns: `Identificacion | Nombre_Completo | Correo | Programa | Cohorte`. 24 students, all `Cohorte=2026-A`. Emails are `@ejemplo.cr` (synthetic).
  - `Registro_Evaluaciones.xlsx` — sheet `Evaluaciones`, range `A1:F89` + header. Columns: `Identificacion | Programa | Modulo | Nota | Asistencia_Pct | Fecha_Cierre`. 88 rows. `Nota`/`Asistencia_Pct` are numeric; rest stored as inline strings (no `sharedStrings.xml` in the .xlsx zip). Dates are `YYYY-MM-DD` text.

## Data gotchas (verified from sheet XML)

- Join key is `Identificacion` (text, preserve leading digits as string, do not cast to int).
- Module count differs by program: `Técnico en IA Aplicada` has Módulo 1–4, `Excel Avanzado para Negocios` has Módulo 1–3 only. Do not assume 4 rows per student.
- Known mismatches (handle explicitly, don't silently drop):
  - `304560321` (Espinoza León, Excel) is in Maestro but has 0 evaluation rows.
  - `999880777` has 4 IA evaluation rows but no Maestro entry (orphan).
- `Programa` appears in both files; if they conflict, prefer Maestro and flag it.

## Environment

- Windows + PowerShell (`shell` tool runs PowerShell, not cmd/bash).
- No Python/Node runtimes installed; only `git` is available. Install or declare a toolchain before building any app (first choice becomes the documented stack).
- PowerShell quoting: `dir /a "path with spaces"` fails (parsed as `C:\a`). Use `Get-ChildItem -Force` with `workdir` param instead of `cd`.
- `Expand-Archive` rejects `.xlsx` extension; to inspect sheets as XML, copy to `*.zip` first, then `[IO.Compression.ZipFile]::ExtractToDirectory(src, dst)` (2-arg form).
