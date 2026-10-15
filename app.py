# app.py — LA COCINA (Academia Horizonte)
# Toda la lógica vive aquí. Las páginas HTML solo muestran lo ya calculado.
# Ladrillos señalados en español: variables, tipos, funciones, bucles y condicionales.

from pathlib import Path  # Ladrillo: tipo para manejar rutas de archivos
from collections import Counter  # Ladrillo: tipo contador para los contadores de arriba
import openpyxl  # Ladrillo: librería para leer los Excel (Bodega)
from flask import Flask, render_template, request  # Ladrillo: Flask = servidor web

# Ladrillo: variables (texto con las rutas; la Bodega NO se modifica)
BASE = Path(__file__).resolve().parent
RUTA_MAESTRO = BASE / "Insumos" / "Maestro_Estudiantes.xlsx"
RUTA_EVALS = BASE / "Insumos" / "Registro_Evaluaciones.xlsx"

# Ladrillo: creación de la aplicación Flask (tipo Flask)
app = Flask(__name__)


# Ladrillo: FUNCIÓN que calcula el promedio con un BUCLE (suma / cantidad)
def calcular_promedio(valores):
    # Ladrillo: variable de tipo número, empieza en 0
    total = 0.0
    # Ladrillo: BUCLE que recorre cada valor de la lista
    for v in valores:  # v es de tipo número (float)
        total = total + float(v)  # suma acumulada
    cantidad = len(valores)  # Ladrillo: variable entera = módulos cursados
    # Ladrillo: CONDICIONAL que evita dividir entre cero
    if cantidad == 0:
        return 0.0
    return total / cantidad


# Ladrillo: FUNCIÓN que decide el certificado con CONDICIONALES (límites INCLUYEN el valor)
def decidir_certificado(promedio, asistencia):
    # Ladrillo: variables de tipo número (float)
    prom = float(promedio)
    asis = float(asistencia)
    # Ladrillo: CONDICIONALES en orden (primero el caso sin certificado)
    if asis < 80:  # si la asistencia es menor a 80, no hay certificado
        return "Sin certificado"
    if prom >= 70:  # asistencia >= 80 y promedio >= 70
        return "Aprobación"
    return "Participación"  # asistencia >= 80 pero promedio < 70


def cargar_datos():
    """Lee los Excel y prepara resultados + contadores + inconsistencias. Todo el cálculo aquí."""
    # --- Leer Maestro: diccionario {identificacion (texto): datos} ---
    wb_m = openpyxl.load_workbook(RUTA_MAESTRO, data_only=True)
    ws_m = wb_m["Estudiantes"]
    cab_m = [c.value for c in ws_m[1]]
    im = {nombre: i for i, nombre in enumerate(cab_m)}  # Ladrillo: diccionario columna->posición
    maestro = {}  # Ladrillo: variable de tipo diccionario
    # Ladrillo: BUCLE por las filas del Maestro (desde la fila 2, la 1 es el encabezado)
    for fila in ws_m.iter_rows(min_row=2, values_only=True):
        ident = str(fila[im["Identificacion"]]).strip()  # Ladrillo: texto, nunca número
        maestro[ident] = {
            "nombre": fila[im["Nombre_Completo"]],
            "programa": fila[im["Programa"]],
        }

    # --- Leer Evaluaciones: agrupar por (identificacion + programa) ---
    wb_e = openpyxl.load_workbook(RUTA_EVALS, data_only=True)
    ws_e = wb_e["Evaluaciones"]
    cab_e = [c.value for c in ws_e[1]]
    ie = {nombre: i for i, nombre in enumerate(cab_e)}
    grupos = {}  # Ladrillo: diccionario {(id, programa): lista de filas}
    # Ladrillo: BUCLE por las filas de Evaluaciones
    for fila in ws_e.iter_rows(min_row=2, values_only=True):
        ident = str(fila[ie["Identificacion"]]).strip()
        prog = fila[ie["Programa"]]
        clave = (ident, prog)  # Ladrillo: tupla = un certificado por estudiante y programa
        # Ladrillo: CONDICIONAL para crear la lista la primera vez
        if clave not in grupos:
            grupos[clave] = []
        grupos[clave].append(fila)

    # --- Cruce: un resultado por estudiante del Maestro ---
    resultados = []  # Ladrillo: lista de diccionarios (tipo lista)
    # Ladrillo: BUCLE por cada estudiante del Maestro
    for ident, datos in maestro.items():
        clave = (ident, datos["programa"])
        evs = grupos.get(clave, [])
        # Ladrillo: CONDICIONAL sin evaluaciones
        if len(evs) == 0:
            resultados.append({
                "identificacion": ident,
                "nombre": datos["nombre"],
                "programa": datos["programa"],
                "modulos": 0,
                "promedio": 0.0,
                "asistencia": 0.0,
                "certificado": "Sin registros",
            })
        else:
            notas = [float(f[ie["Nota"]]) for f in evs]  # Ladrillo: lista de números
            asists = [float(f[ie["Asistencia_Pct"]]) for f in evs]
            prom = round(calcular_promedio(notas), 2)  # Ladrillo: uso de la FUNCIÓN
            asis = round(calcular_promedio(asists), 2)
            tipo = decidir_certificado(prom, asis)  # Ladrillo: uso de la FUNCIÓN
            resultados.append({
                "identificacion": ident,
                "nombre": datos["nombre"],
                "programa": datos["programa"],
                "modulos": len(evs),
                "promedio": prom,
                "asistencia": asis,
                "certificado": tipo,
            })
    # Ladrillo: ordenar por programa y nombre (no calcula, solo ordena)
    resultados.sort(key=lambda r: (r["programa"], r["nombre"]))

    # --- Contadores para la parte de arriba (solo certificados reales) ---
    conteo = Counter(r["certificado"] for r in resultados)  # Ladrillo: contador por tipo
    contadores = {
        "aprobacion": conteo.get("Aprobación", 0),
        "participacion": conteo.get("Participación", 0),
        "sin_certificado": conteo.get("Sin certificado", 0),
    }

    # --- Inconsistencias: las 3 detectadas ---
    inconsistencias = []
    # 1) Del Maestro SIN ninguna evaluación
    # Ladrillo: BUCLE + CONDICIONAL buscando al sin registros
    for ident, datos in maestro.items():
        if (ident, datos["programa"]) not in grupos:
            inconsistencias.append({
                "identificacion": ident,
                "nombre": datos["nombre"],
                "programa": datos["programa"],
                "detalle": "Está en el Maestro pero no tiene ninguna evaluación.",
            })
    # 2) En Evaluaciones pero NO en el Maestro (huérfano: el Maestro manda, queda en revisar)
    for (ident, prog), evs in sorted(grupos.items()):
        # Ladrillo: CONDICIONAL huérfano
        if ident not in maestro:
            notas = [float(f[ie["Nota"]]) for f in evs]
            asists = [float(f[ie["Asistencia_Pct"]]) for f in evs]
            inconsistencias.append({
                "identificacion": ident,
                "nombre": "— (no está en el Maestro)",
                "programa": prog,
                "detalle": f"Tiene {len(evs)} evaluaciones (prom {round(calcular_promedio(notas), 2)}, "
                           f"asis {round(calcular_promedio(asists), 2)}), pero no existe en el Maestro. Queda en revisar.",
            })
    # 3) Sin certificado por asistencia (caso con ceros: posible abandono)
    for r in resultados:
        # Ladrillo: CONDICIONAL para el caso de asistencia menor a 80
        if r["certificado"] == "Sin certificado":
            inconsistencias.append({
                "identificacion": r["identificacion"],
                "nombre": r["nombre"],
                "programa": r["programa"],
                "detalle": f"Promedio {r['promedio']} y asistencia {r['asistencia']}. Sin certificado por asistencia < 80.",
            })

    programas = sorted({r["programa"] for r in resultados})  # Ladrillo: lista para el filtro
    return resultados, contadores, inconsistencias, programas


# Ladrillo: ruta principal (página en http://localhost:5000)
@app.route("/")
def index():
    resultados, contadores, inconsistencias, programas = cargar_datos()
    # Ladrillo: variable texto con el filtro que eligió la persona (o "Todos")
    filtro = request.args.get("programa", "Todos")
    # Ladrillo: CONDICIONAL para filtrar en el backend (el HTML no decide)
    if filtro in programas:
        visibles = [r for r in resultados if r["programa"] == filtro]
    else:
        visibles = resultados
        filtro = "Todos"
    # El HTML solo muestra lo que aquí ya se calculó.
    return render_template("index.html", resultados=visibles, contadores=contadores,
                           inconsistencias=inconsistencias, programas=programas, filtro=filtro)


# Ladrillo: ruta de detalle por estudiante (solo muestra, no calcula nada nuevo)
@app.route("/estudiante/<identificacion>")
def estudiante(identificacion):
    resultados, _, _, _ = cargar_datos()
    # Ladrillo: BUCLE para buscar a la persona por su identificación (texto)
    for r in resultados:
        # Ladrillo: CONDICIONAL que compara textos
        if r["identificacion"] == str(identificacion).strip():
            return render_template("estudiante.html", r=r)
    return render_template("estudiante.html", r=None)


# Ladrillo: arranque de la app (solo se ejecuta al abrir este archivo directo)
if __name__ == "__main__":
    app.run(debug=True)
