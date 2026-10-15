# PLAN.md — App de Certificados (Academia Horizonte)

Construcción por etapas: primero validar el cruce de datos, luego el backend con Flask y al final la interfaz. Sin código todavía.

## Etapa 0 — Preparar la mesa (requisito antes de todo)

- Qué: instalar Python real (el actual es un acceso directo a Microsoft Store) y crear entorno + dependencias.
- Archivos: `requirements.txt` con `Flask` y `openpyxl` (Jinja viene con Flask).
- Validación: `python --version` funciona, `pip install -r requirements.txt` termina sin errores y `python -c "import flask, openpyxl"` no falla.
- No seguir si: Python sigue siendo el stub de la tienda o falta alguna librería.

## Etapa 1 — Validar el cruce de datos (Bodega)

- Qué: leer los 2 Excel solo para comprobar, sin decidir certificados todavía. Revisar: llave `Identificacion` como texto, columnas esperadas, Módulo 1–4 en IA y 1–3 en Excel, casos `304560321` (sin filas) y `999880777` (huérfano), y que no haya conflictos de `Programa` entre archivos.
- Salida esperada: reporte en pantalla o archivo temporal con el conteo (24 estudiantes, 88 filas, 24 IDs en evaluaciones) y la lista de casos raros.
- Validación: los números coinciden con `AGENTS.md` (24 / 88 / huérfano / sin registros / 0 conflictos de programa). Si algo no cuadra, se corrige el plan antes de programar la lógica.
- No seguir si: el reporte no cuadra con lo verificado en el XML.

## Etapa 2 — Backend con Flask, Cocina en `app.py`

- Qué: toda la lógica en `app.py`, nada en los HTML. Funciones simples y comentadas en español (variables, tipos, condicionales, bucles y funciones señalados): cargar Excel con openpyxl, agrupar por `Identificacion + Programa`, calcular `Promedio` (suma de `Nota` / módulos cursados) y `Asistencia` (promedio de `Asistencia_Pct`), y decidir con `>=`: Aprobación (≥70 y ≥80), Participación (<70 y ≥80), Sin certificado (<80). El huérfano `999880777` se marca "revisar" y no genera certificado oficial (el Maestro manda).
- Archivos: `app.py`, `requirements.txt`.
- Rutas mínimas: `/` (lista con resultado), `/estudiante/<identificacion>` (detalle por módulo). La de certificado imprimible (`/certificado/<id>`) solo si la piden después.
- Validación: con los datos actuales la app debe dar **22 oficiales (19 Aprobación + 3 Participación)**, más 1 sin certificado (`111230984`), 1 sin registros (`304560321`) y 1 en revisión (`999880777`). Se comprueba contra esos 5 IDs a mano antes de seguir.
- No seguir si: algún total difiere o hay cálculo en una plantilla HTML.

## Etapa 3 — Interfaz, Salón (HTML + CSS)

- Qué: plantillas que solo muestran lo que `app.py` ya calculó. Diseño azul marino y dorado.
- Archivos: `templates/base.html`, `templates/index.html`, `templates/estudiante.html`, `static/style.css`.
- Validación: abrir `/` y un detalle en el navegador; revisar que un estudiante de cada tipo se vea bien (ej.: Aprobación `101230456`, Participación `116890234`, Sin certificado `111230984`, Sin registros `304560321`). Sin errores en consola y sin lógica en los HTML.
- No seguir si: falta algún caso en pantalla o hay promedios calculados en Jinja.

## Orden y regla general

Orden: 0 → 1 → 2 → 3. Cada etapa se valida y solo entonces se pasa a la siguiente. Si una validación falla, se arregla esa etapa, no se salta.
