# -*- coding: utf-8 -*-
"""
Optimiza la hoja 'CUADRO DE MANDO T' del Cuadro de Mando de SIMCARDS.

Qué hace
--------
1. Corrige los subtotales mal rangeados (filas 31 y 44 del archivo original).
2. Sustituye los SUM() de rango fijo por SUMIFS() sobre una columna de clave,
   de modo que los totales nunca se descuadren al insertar o mover artículos.
3. Agrega un segundo nivel de separación (SUBGRUPO = familia de producto)
   manteniendo los 6 grupos/equipos originales.
4. Parametriza el divisor del promedio diario (antes constante 20) en la
   celda nombrada DIAS_PERIODO, con detección automática del período real.
5. Agrega fila de TOTAL GENERAL y un bloque de control de cuadratura.
6. Limpia espacios no separables (NBSP) que rompen los SUMIFS en silencio.
7. Cierra el hueco del formato condicional en el valor exacto 15.

El resto del libro (PMP, REFERENCIA, RESUMEN *, INVENTARIO, DESPACHOS 20 DIAS,
tablas dinámicas, imágenes y estilos) se conserva sin cambios.

Uso:  python3 optimizar_cuadro.py <entrada.xlsx> <salida.xlsx>
"""
import re
import sys
from copy import copy

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

HOJA = "CUADRO DE MANDO T"
DESP = "DESPACHOS 20 DIAS"
INV = "INVENTARIO"

FILA_DET_INI = 13          # primera fila de detalle
COL_MIN, COL_MAX = 1, 18   # A..R

# ---------------------------------------------------------------- taxonomía
# clave de subgrupo -> (clave de grupo, etiqueta del subgrupo)
SUBGRUPOS = {
    "KITE.ESTANDAR":  ("KITE",     "SIM / HALFSIM KITE M2M ESTÁNDAR"),
    "KITE.DEVCTRL":   ("KITE",     "HALFSIM KITE M2M DEVICE CONTROL"),
    "M2MLTE.LTE":     ("M2MLTE",   "SIMCARD / HALFSIM M2M LTE"),
    "MOVISTAR.HALFSIM": ("MOVISTAR", "HALFSIM MOVISTAR (FÍSICA)"),
    "MOVISTAR.ESIM":  ("MOVISTAR", "ESIM MOVISTAR (DIGITAL)"),
    "TIGO.HALFSIM":   ("TIGO",     "HALFSIM TIGO (COMERCIAL)"),
    "TIGO.CHIP3":     ("TIGO",     "HALFSIM TIGO CHIP PREPAGO $3 (NUMERADA)"),
    "TIGO.ESIM":      ("TIGO",     "ESIM TIGO DIGITAL"),
    "TIGO.M2M":       ("TIGO",     "HALFSIM TIGO M2M"),
    "AKI.HALFSIM":    ("AKI",      "HALFSIM AKIMOVIL (FÍSICA)"),
    "AKI.NUMERADA":   ("AKI",      "HALFSIM AKIMOVIL NUMERADA"),
    "AKI.ESIM":       ("AKI",      "ESIM AKIMOVIL (DIGITAL)"),
    "MAXI.HALFSIM":   ("MAXI",     "HALFSIM MAXIPLUS (FÍSICA)"),
    "MAXI.NUMERADA":  ("MAXI",     "HALFSIM MAXIPLUS NUMERADA"),
    "MAXI.ESIM":      ("MAXI",     "ESIM MAXIPLUS (DIGITAL)"),
}

GRUPOS = [  # orden de presentación: (clave, etiqueta, [subgrupos en orden])
    ("KITE",     "KITE M2M",  ["KITE.ESTANDAR", "KITE.DEVCTRL"]),
    ("M2MLTE",   "M2M LTE",   ["M2MLTE.LTE"]),
    ("MOVISTAR", "MOVISTAR",  ["MOVISTAR.HALFSIM", "MOVISTAR.ESIM"]),
    ("TIGO",     "TIGO",      ["TIGO.HALFSIM", "TIGO.CHIP3", "TIGO.ESIM", "TIGO.M2M"]),
    ("AKI",      "AKIMOVIL",  ["AKI.HALFSIM", "AKI.NUMERADA", "AKI.ESIM"]),
    ("MAXI",     "MAXIPLUS",  ["MAXI.HALFSIM", "MAXI.NUMERADA", "MAXI.ESIM"]),
]

# (codigo, descripcion_normalizada) -> clave de subgrupo
def clasificar(codigo, desc):
    d = desc.upper()
    if codigo in ("162082", "162086", "162439"):
        return "KITE.ESTANDAR"
    if "KITE" in d:
        return "KITE.DEVCTRL"
    if codigo in ("162115", "162145", "162148"):
        return "M2MLTE.LTE"
    if "MOVISTAR" in d:
        return "MOVISTAR.ESIM" if d.startswith("ESIM") else "MOVISTAR.HALFSIM"
    if "TIGO" in d:
        if d.startswith("ESIM"):
            return "TIGO.ESIM"
        if "CHIP PREPAGO" in d:
            return "TIGO.CHIP3"
        if "M2M" in d:
            return "TIGO.M2M"
        return "TIGO.HALFSIM"
    if "AKIMOVIL" in d:
        if d.startswith("ESIM"):
            return "AKI.ESIM"
        return "AKI.NUMERADA" if "NUMERADA" in d else "AKI.HALFSIM"
    if "MAXIPLUS" in d:
        if d.startswith("ESIM"):
            return "MAXI.ESIM"
        return "MAXI.NUMERADA" if "NUMERADA" in d else "MAXI.HALFSIM"
    raise ValueError("Artículo sin clasificar: %r / %r" % (codigo, desc))


def limpiar(texto):
    """Quita NBSP y espacios sobrantes sin alterar el contenido visible."""
    if not isinstance(texto, str):
        return texto
    return " ".join(texto.replace("\xa0", " ").split())


# --------------------------------------------------------------- fórmulas
def f_despacho(fila, canal):
    return ("=SUMIFS('{d}'!F:F,'{d}'!D:D,$C{r},'{d}'!E:E,$D{r},'{d}'!N:N,\"{c}\")"
            .format(d=DESP, r=fila, c=canal))


def f_inventario(fila, almacen):
    return ("=SUMIFS({i}!E:E,{i}!D:D,$C{r},{i}!B:B,$D{r},{i}!A:A,\"{a}\")"
            .format(i=INV, r=fila, a=almacen))


def f_dias(fila, col_stock):
    """Misma lógica de negocio del archivo original, sin cambios de criterio."""
    r = fila
    s = col_stock
    return ('=IF(E{r}="Numerada",'
            'IF(AND(F{r}=0,G{r}=0,H{r}=0,I{r}=0),">15",'
            'IF(ISERROR({s}{r}/I{r})=FALSE,{s}{r}/I{r},0)),'
            'IF(AND(F{r}=0,G{r}=0,H{r}=0,I{r}=0),">150",'
            'IF(ISERROR({s}{r}/I{r})=FALSE,{s}{r}/I{r},0)))'
            ).format(r=r, s=s)


def f_banda(col, criterio, fila_ini, fila_fin):
    return '=SUMIFS({c}${a}:{c}${b},$A${a}:$A${b},"{k}")'.format(
        c=col, a=fila_ini, b=fila_fin, k=criterio)


# ------------------------------------------------------------------ main
def main(entrada, salida):
    wb = openpyxl.load_workbook(entrada)
    ws = wb[HOJA]

    # ---- 1. leer el detalle original (filas 13..58, saltando subtotales) --
    articulos = []
    for r in range(FILA_DET_INI, 60):
        cod, desc, tipo = ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value
        if not cod or not desc:
            continue
        cod, desc = str(cod).strip(), limpiar(str(desc))
        articulos.append({"cod": cod, "desc": desc, "tipo": tipo,
                          "sub": clasificar(cod, desc), "orden": r})
    assert len(articulos) == 41, "Se esperaban 41 artículos, se leyeron %d" % len(articulos)

    # orden explícito dentro de los subgrupos numerados (7D, 15D, 30D)
    def clave_orden(a):
        d = a["desc"].upper()
        for n, peso in (("NUMERADA 7D", 1), ("NUMERADA 15D", 2), ("NUMERADA 30D", 3)):
            if n in d:
                return (peso, a["orden"])
        return (0, a["orden"])

    # ---- 2. plantillas de estilo ------------------------------------------
    est_detalle = {c: copy(ws.cell(FILA_DET_INI, c)._style) for c in range(COL_MIN, COL_MAX + 1)}
    est_banda = {c: copy(ws.cell(21, c)._style) for c in range(COL_MIN, COL_MAX + 1)}

    FILL_GRUPO = PatternFill("solid", fgColor="FF0070C0")   # azul original
    FILL_SUB = PatternFill("solid", fgColor="FF9DC3E6")     # azul claro (nivel 2)
    FILL_TOTAL = PatternFill("solid", fgColor="FF1F4E79")   # azul oscuro (total)
    FILL_CLAVE = PatternFill("solid", fgColor="FFF2F2F2")
    FONT_SUB = Font(name="Calibri", size=8, bold=True, color="FF1F3864")
    FONT_GRUPO = Font(name="Calibri", size=8, bold=True, color="FFFFFFFF")
    FONT_CLAVE = Font(name="Calibri", size=7, italic=True, color="FF808080")

    # ---- 3. limpiar la zona de detalle ------------------------------------
    for r in range(FILA_DET_INI, 1002):
        for c in range(COL_MIN, COL_MAX + 1):
            ws.cell(r, c).value = None

    # ---- 4. calcular el layout --------------------------------------------
    plan = []          # (tipo_fila, payload)
    for gkey, glabel, subs in GRUPOS:
        subs_reales = [s for s in subs if any(a["sub"] == s for a in articulos)]
        for skey in subs_reales:
            items = sorted([a for a in articulos if a["sub"] == skey], key=clave_orden)
            for a in items:
                plan.append(("det", a))
            if len(subs_reales) > 1:      # banda de subgrupo solo si aporta
                plan.append(("sub", (skey, SUBGRUPOS[skey][1])))
        plan.append(("grp", (gkey, glabel)))

    fila_ini = FILA_DET_INI
    fila_fin = fila_ini + len(plan) - 1
    fila_total = fila_fin + 1

    # ---- 5. escribir ------------------------------------------------------
    no = 0
    r = fila_ini
    for tipo_fila, payload in plan:
        for c in range(COL_MIN, COL_MAX + 1):
            ws.cell(r, c)._style = copy(
                est_detalle[c] if tipo_fila == "det" else est_banda[c])

        if tipo_fila == "det":
            a = payload
            no += 1
            ws.cell(r, 1, a["sub"])                     # A  clave de subgrupo
            ws.cell(r, 1).fill = FILL_CLAVE
            ws.cell(r, 1).font = FONT_CLAVE
            ws.cell(r, 1).alignment = Alignment(horizontal="left")
            ws.cell(r, 2, no)                           # B  No.
            ws.cell(r, 3, a["cod"])                     # C  CÓDIGO JANUS
            ws.cell(r, 4, a["desc"])                    # D  ARTÍCULO JANUS
            ws.cell(r, 5, a["tipo"])                    # E  Tipo
            ws.cell(r, 6, f_despacho(r, "CANAL INDIRECTO"))
            ws.cell(r, 7, f_despacho(r, "CANAL DIRECTO"))
            ws.cell(r, 8, "=F{r}+G{r}".format(r=r))
            ws.cell(r, 9, "=IFERROR(H{r}/DIAS_PERIODO,0)".format(r=r))
            ws.cell(r, 10, 0)                           # J  en proceso de liberación
            ws.cell(r, 11, f_inventario(r, "PRODUCCION"))
            ws.cell(r, 12, f_inventario(r, "MATERIA PRIMA"))
            ws.cell(r, 13, f_dias(r, "L"))
            ws.cell(r, 14, "={a}+{b}".format(a=f_inventario(r, "BODEGA CENTRAL")[1:],
                                             b=f_inventario(r, "DELIVERY")[1:]))
            ws.cell(r, 15, f_dias(r, "N"))
            ws.cell(r, 16, "=L{r}+N{r}+K{r}".format(r=r))
            ws.cell(r, 17, f_dias(r, "P"))
        else:
            key, label = payload
            criterio = key if tipo_fila == "sub" else key + ".*"
            es_grupo = tipo_fila == "grp"
            fill = FILL_GRUPO if es_grupo else FILL_SUB
            font = FONT_GRUPO if es_grupo else FONT_SUB
            for c in range(COL_MIN + 1, COL_MAX + 1):   # B..R
                ws.cell(r, c).fill = fill
                ws.cell(r, c).font = font
            ws.cell(r, 1).fill = PatternFill()
            ws.cell(r, 4, ("TOTAL " if es_grupo else "Subtotal ") + label)
            ws.cell(r, 4).alignment = Alignment(horizontal="left", vertical="center")
            for col in ("F", "G", "H", "J", "K", "L", "N", "P"):
                ws.cell(r, openpyxl.utils.column_index_from_string(col),
                        f_banda(col, criterio, fila_ini, fila_fin))
            ws.cell(r, 9, "=IFERROR(H{r}/DIAS_PERIODO,0)".format(r=r))
            # tipo homogéneo -> el semáforo del grupo se colorea solo
            tipos = {a["tipo"] for a in articulos
                     if (a["sub"] == key if tipo_fila == "sub"
                         else SUBGRUPOS[a["sub"]][0] == key)}
            ws.cell(r, 5, tipos.pop() if len(tipos) == 1 else None)
            ws.cell(r, 13, f_dias(r, "L"))
            ws.cell(r, 15, f_dias(r, "N"))
            ws.cell(r, 17, f_dias(r, "P"))
        r += 1

    # ---- 6. fila de TOTAL GENERAL -----------------------------------------
    r = fila_total
    for c in range(COL_MIN, COL_MAX + 1):
        ws.cell(r, c)._style = copy(est_banda[c])
        if c > COL_MIN:
            ws.cell(r, c).fill = FILL_TOTAL
            ws.cell(r, c).font = Font(name="Calibri", size=8, bold=True, color="FFFFFFFF")
    ws.cell(r, 1).fill = PatternFill()
    ws.cell(r, 4, "TOTAL GENERAL")
    ws.cell(r, 4).alignment = Alignment(horizontal="left", vertical="center")
    for col in ("F", "G", "H", "J", "K", "L", "N", "P"):
        ws.cell(r, openpyxl.utils.column_index_from_string(col),
                f_banda(col, "?*", fila_ini, fila_fin))
    ws.cell(r, 9, "=IFERROR(H{r}/DIAS_PERIODO,0)".format(r=r))
    ws.cell(r, 13, "=IFERROR(L{r}/I{r},0)".format(r=r))
    ws.cell(r, 15, "=IFERROR(N{r}/I{r},0)".format(r=r))
    ws.cell(r, 17, "=IFERROR(P{r}/I{r},0)".format(r=r))

    # ---- 6bis. formatos numéricos homogéneos ------------------------------
    # Las bandas heredaban formato "General" y desbordaban con decimales.
    # El promedio diario pasa a 0.00: con formato entero, 0,05/día se leía "0".
    for rr in range(FILA_DET_INI, fila_total + 1):
        for c in (6, 7, 8, 10, 11, 12, 14, 16):      # cantidades
            ws.cell(rr, c).number_format = "#,##0"
        ws.cell(rr, 9).number_format = "#,##0.00"    # promedio diario
        for c in (13, 15, 17):                       # días de stock
            ws.cell(rr, c).number_format = "#,##0"
    for rr in (fila_total + 1 + 1, fila_total + 1 + 2):
        for c in (7, 8, 9):
            ws.cell(rr, c).number_format = "#,##0"

    # ---- 7. encabezado: clave + parámetro del período ---------------------
    ws.cell(11, 1, "CLAVE")
    ws.cell(11, 1)._style = copy(ws.cell(11, 2)._style)
    ws.cell(12, 1)._style = copy(ws.cell(12, 2)._style)
    ws.column_dimensions["A"].width = 15.5

    rng = ("OFFSET('{d}'!$A$2,0,0,MAX(1,COUNTA('{d}'!$A:$A)-1),1)".format(d=DESP))
    ws.cell(7, 5, "DÍAS DEL PERÍODO (divisor del promedio diario):")
    ws.cell(7, 9, 20)
    ws.cell(8, 5, "Días con despacho detectados en la hoja origen:")
    ws.cell(8, 9, "=IFERROR(SUMPRODUCT(1/COUNTIF({r},{r})),0)".format(r=rng))
    ws.cell(9, 5, '=IF($I$8=$I$7,"✔ Divisor alineado con el período cargado",'
                  '"⚠ REVISAR: el período cargado tiene "&$I$8&" días y el divisor es "&$I$7)')
    for fila in (7, 8, 9):
        ws.merge_cells(start_row=fila, start_column=5, end_row=fila, end_column=8)
        ws.cell(fila, 5).font = Font(name="Calibri", size=8, bold=(fila == 9))
        ws.cell(fila, 5).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(fila, 9).font = Font(name="Calibri", size=8, bold=True)
        ws.cell(fila, 9).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(7, 9).fill = PatternFill("solid", fgColor="FFFFF2CC")
    ws.cell(9, 5).alignment = Alignment(horizontal="left", vertical="center")

    wb.defined_names.add(DefinedName("DIAS_PERIODO",
                                     attr_text="'{h}'!$I$7".format(h=HOJA)))

    # ---- 8. bloque de control de cuadratura -------------------------------
    base = fila_total + 2
    ws.cell(base, 3, "CONTROL DE CUADRATURA (debe dar 0)")
    ws.cell(base, 3).font = Font(name="Calibri", size=9, bold=True, color="FF1F4E79")
    filas_ctl = [
        ("Stock total del cuadro vs. hoja INVENTARIO",
         "=P{t}".format(t=fila_total),
         "=SUM({i}!E:E)".format(i=INV)),
        ("Despacho total del cuadro vs. hoja DESPACHOS",
         "=H{t}".format(t=fila_total),
         "=SUM('{d}'!F:F)".format(d=DESP)),
    ]
    for k, (etiqueta, f_cuadro, f_origen) in enumerate(filas_ctl):
        rr = base + 1 + k
        ws.cell(rr, 3, etiqueta)
        ws.cell(rr, 7, f_cuadro)
        ws.cell(rr, 8, f_origen)
        ws.cell(rr, 9, "=G{r}-H{r}".format(r=rr))
        ws.cell(rr, 10, '=IF(I{r}=0,"OK","REVISAR: hay artículos en la hoja origen '
                        'que no están en el cuadro")'.format(r=rr))
        for c in (3, 7, 8, 9, 10):
            ws.cell(rr, c).font = Font(name="Calibri", size=8,
                                       bold=(c == 10), color="FF404040")
    ws.cell(base + 3, 3,
            "Las bandas se calculan con SUMIFS sobre la columna A (CLAVE). "
            "Para agregar un artículo: inserte la fila dentro de su bloque, "
            "copie las fórmulas de la fila de arriba y escriba la misma CLAVE en la columna A.")
    ws.cell(base + 3, 3).font = Font(name="Calibri", size=8, italic=True, color="FF808080")

    # ---- 9. formato condicional: cerrar el hueco en el valor 15 -----------
    for rango, reglas in ws.conditional_formatting._cf_rules.items():
        for regla in reglas:
            if regla.type == "expression" and regla.formula:
                # ojo: '>15' no debe tocar '>150' (umbral de las SIM Blanca)
                regla.formula = [re.sub(r'([MOQ]13)>15(?![0-9])', r'\1>=15', f)
                                 for f in regla.formula]

    # ---- 10. limpiar NBSP en la hoja INVENTARIO ---------------------------
    wsi = wb[INV]
    for fila in wsi.iter_rows(min_row=2, max_row=wsi.max_row, min_col=1, max_col=4):
        for celda in fila:
            if isinstance(celda.value, str):
                nuevo = limpiar(celda.value)
                if nuevo != celda.value:
                    celda.value = nuevo

    wb.save(salida)
    print("OK -> %s" % salida)
    print("   detalle: filas %d..%d   TOTAL GENERAL: fila %d" % (fila_ini, fila_fin, fila_total))
    print("   artículos: %d   bandas: %d" % (no, len(plan) - no))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
