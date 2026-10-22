# -*- coding: utf-8 -*-
"""
Optimiza la hoja 'CUADRO DE MANDO T' del Cuadro de Mando de SIMCARDS.

Principio: no se adultera ningún dato de origen. Las hojas PMP, REFERENCIA,
RESUMEN INVENTARIO T, RESUMEN DESPACHO y DESPACHOS 20 DIAS se restauran al
final con sus valores numéricos originales bit a bit (paso 12), porque la
reescritura del formato XLSX trunca los flotantes a 15 dígitos. La única celda
de origen que cambia es el NBSP de INVENTARIO!B46, que estaba roto.

Correcciones aplicadas a la hoja del cuadro:
  1. Subtotales mal rangeados (filas 31 y 44 del original).
  2. Bandas con SUMIFS sobre columna de clave: no se descuadran nunca.
  3. Segundo nivel de separación (familia de producto) con sus cifras.
  4. Divisor del promedio diario: AUTO (período real) o fijo, parametrizado.
  5. Días de stock numéricos de verdad (antes texto ">150"/">15"), con el
     mismo aspecto gracias a un formato numérico condicional.
  6. Columna ESTADO: distingue SIN ROTACIÓN de OK (antes ambos en verde).
  7. Columna ALERTA DE PRODUCCIÓN: detecta cobertura total sana con producto
     terminado en rojo.
  8. Fila TOTAL GENERAL y bloque de control de cuadratura.
  9. Formato condicional reconstruido sobre valores numéricos.
 10. Limpieza de NBSP que rompían un SUMIFS en silencio.

Uso:  python3 optimizar_cuadro.py <entrada.xlsx> <salida.xlsx>
"""
import re
import shutil
import sys
import zipfile
from copy import copy

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string
from openpyxl.workbook.defined_name import DefinedName

HOJA = "CUADRO DE MANDO T"
DESP = "DESPACHOS 20 DIAS"
INV = "INVENTARIO"
REF = "REFERENCIA"

FILA_DET_INI = 13
COL_MIN, COL_MAX = 1, 19          # A..S

C_EST, C_ALE = 18, 19             # R = ESTADO, S = ALERTA DE PRODUCCIÓN

# Hojas cuyos valores se restauran tal cual del original.
HOJAS_INTACTAS = ["PMP", REF, "RESUMEN INVENTARIO T", "RESUMEN DESPACHO", DESP]

# ---------------------------------------------------------------- taxonomía
SUBGRUPOS = {
    "KITE.ESTANDAR":    ("KITE",     "SIM / HALFSIM KITE M2M ESTÁNDAR"),
    "KITE.DEVCTRL":     ("KITE",     "HALFSIM KITE M2M DEVICE CONTROL"),
    "M2MLTE.LTE":       ("M2MLTE",   "SIMCARD / HALFSIM M2M LTE"),
    "MOVISTAR.HALFSIM": ("MOVISTAR", "HALFSIM MOVISTAR (FÍSICA)"),
    "MOVISTAR.ESIM":    ("MOVISTAR", "ESIM MOVISTAR (DIGITAL)"),
    "TIGO.HALFSIM":     ("TIGO",     "HALFSIM TIGO (COMERCIAL)"),
    "TIGO.CHIP3":       ("TIGO",     "HALFSIM TIGO CHIP PREPAGO $3 (NUMERADA)"),
    "TIGO.ESIM":        ("TIGO",     "ESIM TIGO DIGITAL"),
    "TIGO.M2M":         ("TIGO",     "HALFSIM TIGO M2M"),
    "AKI.HALFSIM":      ("AKI",      "HALFSIM AKIMOVIL (FÍSICA)"),
    "AKI.NUMERADA":     ("AKI",      "HALFSIM AKIMOVIL NUMERADA"),
    "AKI.ESIM":         ("AKI",      "ESIM AKIMOVIL (DIGITAL)"),
    "MAXI.HALFSIM":     ("MAXI",     "HALFSIM MAXIPLUS (FÍSICA)"),
    "MAXI.NUMERADA":    ("MAXI",     "HALFSIM MAXIPLUS NUMERADA"),
    "MAXI.ESIM":        ("MAXI",     "ESIM MAXIPLUS (DIGITAL)"),
}

GRUPOS = [
    ("KITE",     "KITE M2M", ["KITE.ESTANDAR", "KITE.DEVCTRL"]),
    ("M2MLTE",   "M2M LTE",  ["M2MLTE.LTE"]),
    ("MOVISTAR", "MOVISTAR", ["MOVISTAR.HALFSIM", "MOVISTAR.ESIM"]),
    ("TIGO",     "TIGO",     ["TIGO.HALFSIM", "TIGO.CHIP3", "TIGO.ESIM", "TIGO.M2M"]),
    ("AKI",      "AKIMOVIL", ["AKI.HALFSIM", "AKI.NUMERADA", "AKI.ESIM"]),
    ("MAXI",     "MAXIPLUS", ["MAXI.HALFSIM", "MAXI.NUMERADA", "MAXI.ESIM"]),
]

# Umbrales de negocio (sin cambios respecto del original)
U_NUM_ROJO, U_NUM_AMBAR = 8, 15      # Numerada: <8 rojo, 8-14 ámbar, >=15 verde
U_BLA_ROJO, U_BLA_AMBAR = 90, 150    # Blanca:  <=90 rojo, 91-150 ámbar, >150 verde

FMT_CANT = "#,##0"
FMT_PROM = "#,##0.00"
FMT_DIAS_BLANCA = '[>150]">150";#,##0'
FMT_DIAS_NUMERADA = '[>15]">15";#,##0'


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
        return "TIGO.M2M" if "M2M" in d else "TIGO.HALFSIM"
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
    if not isinstance(texto, str):
        return texto
    return " ".join(texto.replace("\xa0", " ").split())


# --------------------------------------------------------------- fórmulas
def f_despacho(r, canal):
    return ("=SUMIFS('{d}'!F:F,'{d}'!D:D,$C{r},'{d}'!E:E,$D{r},'{d}'!N:N,\"{c}\")"
            .format(d=DESP, r=r, c=canal))


def f_inventario(r, almacen):
    return ("SUMIFS({i}!E:E,{i}!D:D,$C{r},{i}!B:B,$D{r},{i}!A:A,\"{a}\")"
            .format(i=INV, r=r, a=almacen))


def f_dias(r, col_stock):
    """Días de stock NUMÉRICOS. Vacío cuando no hay rotación (sin denominador).

    Antes devolvía el texto ">150"/">15", que impedía ordenar, promediar o
    graficar la columna. El rótulo se conserva mediante formato numérico
    condicional, así que el aspecto es el mismo pero la celda es un número.
    """
    return '=IF($I{r}=0,"",{s}{r}/$I{r})'.format(r=r, s=col_stock)


def f_estado(r):
    return ('=IF($D{r}="","",'
            'IF($I{r}=0,"SIN ROTACIÓN",'
            'IF($E{r}="","",'
            'IF($E{r}="Numerada",'
            'IF($Q{r}<{nr},"CRÍTICO",IF($Q{r}<{na},"ALERTA","OK")),'
            'IF($Q{r}<={br},"CRÍTICO",IF($Q{r}<={ba},"ALERTA","OK"))))))'
            ).format(r=r, nr=U_NUM_ROJO, na=U_NUM_AMBAR,
                     br=U_BLA_ROJO, ba=U_BLA_AMBAR)


def f_alerta_produccion(r):
    """Cobertura total sana pero producto terminado en rojo.

    Es el riesgo que el tablero no mostraba: HALFSIM TIGO tiene cobertura
    total verde y apenas unos días de producto terminado. No es un problema
    de abastecimiento sino de capacidad de producción.
    """
    # Solo en filas de artículo ($C con código): la acción es sobre un SKU.
    return ('=IF(OR($C{r}="",$E{r}="",$I{r}=0),"",'
            'IF($E{r}="Numerada",'
            'IF(AND($O{r}<{nr},$Q{r}>={nr}),"PRODUCIR - "&TEXT($O{r},"0")&" d terminado",""),'
            'IF(AND($O{r}<={br},$Q{r}>{br}),"PRODUCIR - "&TEXT($O{r},"0")&" d terminado","")))'
            ).format(r=r, nr=U_NUM_ROJO, br=U_BLA_ROJO)


def f_banda(col, criterio, ini, fin):
    return '=SUMIFS({c}${a}:{c}${b},$A${a}:$A${b},"{k}")'.format(
        c=col, a=ini, b=fin, k=criterio)


def rango_dinamico(hoja, col):
    """Rango del alto exacto de los datos; evita COUNTIF sobre 1M de filas."""
    return ("OFFSET('{h}'!${c}$2,0,0,MAX(1,COUNTA('{h}'!${c}:${c})-1),1)"
            .format(h=hoja, c=col))


# ------------------------------------------------------------------ main
def main(entrada, salida):
    wb = openpyxl.load_workbook(entrada)
    ws = wb[HOJA]

    # ---- 1. leer el detalle original --------------------------------------
    articulos = []
    for r in range(FILA_DET_INI, 60):
        cod, desc, tipo = ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value
        if not cod or not desc:
            continue
        cod, desc = str(cod).strip(), limpiar(str(desc))
        articulos.append({"cod": cod, "desc": desc, "tipo": tipo,
                          "sub": clasificar(cod, desc), "orden": r})
    assert len(articulos) == 41, "Se esperaban 41 artículos, se leyeron %d" % len(articulos)

    def clave_orden(a):
        d = a["desc"].upper()
        for n, peso in (("NUMERADA 7D", 1), ("NUMERADA 15D", 2), ("NUMERADA 30D", 3)):
            if n in d:
                return (peso, a["orden"])
        return (0, a["orden"])

    # ---- 2. plantillas de estilo ------------------------------------------
    est_detalle = {c: copy(ws.cell(FILA_DET_INI, min(c, 17))._style) for c in range(COL_MIN, COL_MAX + 1)}
    est_banda = {c: copy(ws.cell(21, min(c, 17))._style) for c in range(COL_MIN, COL_MAX + 1)}

    FILL_GRUPO = PatternFill("solid", fgColor="FF0070C0")
    FILL_SUB = PatternFill("solid", fgColor="FF9DC3E6")
    FILL_TOTAL = PatternFill("solid", fgColor="FF1F4E79")
    FILL_CLAVE = PatternFill("solid", fgColor="FFF2F2F2")
    FONT_SUB = Font(name="Calibri", size=8, bold=True, color="FF1F3864")
    FONT_BLANCO = Font(name="Calibri", size=8, bold=True, color="FFFFFFFF")
    FONT_CLAVE = Font(name="Calibri", size=7, italic=True, color="FF808080")

    # ---- 3. limpiar la zona de detalle ------------------------------------
    for r in range(FILA_DET_INI, 1002):
        for c in range(COL_MIN, COL_MAX + 1):
            ws.cell(r, c).value = None

    # ---- 4. layout ---------------------------------------------------------
    plan = []
    for gkey, glabel, subs in GRUPOS:
        subs_reales = [s for s in subs if any(a["sub"] == s for a in articulos)]
        for skey in subs_reales:
            for a in sorted([x for x in articulos if x["sub"] == skey], key=clave_orden):
                plan.append(("det", a))
            if len(subs_reales) > 1:
                plan.append(("sub", (skey, SUBGRUPOS[skey][1])))
        plan.append(("grp", (gkey, glabel)))

    ini = FILA_DET_INI
    fin = ini + len(plan) - 1
    total = fin + 1

    # ---- 5. escribir -------------------------------------------------------
    no = 0
    r = ini
    for kind, payload in plan:
        for c in range(COL_MIN, COL_MAX + 1):
            ws.cell(r, c)._style = copy(est_detalle[c] if kind == "det" else est_banda[c])

        if kind == "det":
            a = payload
            no += 1
            ws.cell(r, 1, a["sub"])
            ws.cell(r, 1).fill = FILL_CLAVE
            ws.cell(r, 1).font = FONT_CLAVE
            ws.cell(r, 1).alignment = Alignment(horizontal="left")
            ws.cell(r, 2, no)
            ws.cell(r, 3, a["cod"])
            ws.cell(r, 4, a["desc"])
            ws.cell(r, 5, a["tipo"])
            ws.cell(r, 6, f_despacho(r, "CANAL INDIRECTO"))
            ws.cell(r, 7, f_despacho(r, "CANAL DIRECTO"))
            ws.cell(r, 8, "=F{r}+G{r}".format(r=r))
            ws.cell(r, 9, "=IFERROR(H{r}/DIAS_PERIODO,0)".format(r=r))
            ws.cell(r, 10, 0)
            ws.cell(r, 11, "=" + f_inventario(r, "PRODUCCION"))
            ws.cell(r, 12, "=" + f_inventario(r, "MATERIA PRIMA"))
            ws.cell(r, 14, "=" + f_inventario(r, "BODEGA CENTRAL")
                            + "+" + f_inventario(r, "DELIVERY"))
            ws.cell(r, 16, "=L{r}+N{r}+K{r}".format(r=r))
            tipo_fila = a["tipo"]
        else:
            key, label = payload
            es_grupo = kind == "grp"
            criterio = key + ".*" if es_grupo else key
            for c in range(COL_MIN + 1, COL_MAX + 1):
                ws.cell(r, c).fill = FILL_GRUPO if es_grupo else FILL_SUB
                ws.cell(r, c).font = FONT_BLANCO if es_grupo else FONT_SUB
            ws.cell(r, 1).fill = PatternFill()
            ws.cell(r, 4, ("TOTAL " if es_grupo else "Subtotal ") + label)
            ws.cell(r, 4).alignment = Alignment(horizontal="left", vertical="center")
            for col in ("F", "G", "H", "J", "K", "L", "N", "P"):
                ws.cell(r, column_index_from_string(col), f_banda(col, criterio, ini, fin))
            ws.cell(r, 9, "=IFERROR(H{r}/DIAS_PERIODO,0)".format(r=r))
            tipos = {a["tipo"] for a in articulos
                     if (a["sub"] == key if kind == "sub" else SUBGRUPOS[a["sub"]][0] == key)}
            tipo_fila = tipos.pop() if len(tipos) == 1 else None
            ws.cell(r, 5, tipo_fila)

        # comunes a detalle y banda
        for col_dias, col_stock in ((13, "L"), (15, "N"), (17, "P")):
            ws.cell(r, col_dias, f_dias(r, col_stock))
        ws.cell(r, C_EST, f_estado(r))
        ws.cell(r, C_ALE, f_alerta_produccion(r))
        fmt_dias = FMT_DIAS_NUMERADA if tipo_fila == "Numerada" else FMT_DIAS_BLANCA
        for c in (6, 7, 8, 10, 11, 12, 14, 16):
            ws.cell(r, c).number_format = FMT_CANT
        ws.cell(r, 9).number_format = FMT_PROM
        for c in (13, 15, 17):
            ws.cell(r, c).number_format = fmt_dias
        for c in (C_EST, C_ALE):
            ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="center")
        r += 1

    # ---- 6. TOTAL GENERAL --------------------------------------------------
    r = total
    for c in range(COL_MIN, COL_MAX + 1):
        ws.cell(r, c)._style = copy(est_banda[c])
        if c > COL_MIN:
            ws.cell(r, c).fill = FILL_TOTAL
            ws.cell(r, c).font = FONT_BLANCO
    ws.cell(r, 1).fill = PatternFill()
    ws.cell(r, 4, "TOTAL GENERAL")
    ws.cell(r, 4).alignment = Alignment(horizontal="left", vertical="center")
    for col in ("F", "G", "H", "J", "K", "L", "N", "P"):
        ws.cell(r, column_index_from_string(col), f_banda(col, "?*", ini, fin))
    ws.cell(r, 9, "=IFERROR(H{r}/DIAS_PERIODO,0)".format(r=r))
    for col_dias, col_stock in ((13, "L"), (15, "N"), (17, "P")):
        ws.cell(r, col_dias, f_dias(r, col_stock))
    for c in (6, 7, 8, 10, 11, 12, 14, 16):
        ws.cell(r, c).number_format = FMT_CANT
    ws.cell(r, 9).number_format = FMT_PROM
    for c in (13, 15, 17):
        ws.cell(r, c).number_format = FMT_DIAS_BLANCA

    # ---- 7. encabezados de las columnas nuevas ----------------------------
    ws.cell(11, 1, "CLAVE")
    ws.cell(11, 1)._style = copy(ws.cell(11, 2)._style)
    ws.cell(12, 1)._style = copy(ws.cell(12, 2)._style)
    for col, titulo, ancho in ((C_EST, "ESTADO", 15.5), (C_ALE, "ALERTA DE PRODUCCIÓN", 26.0)):
        for fila in (11, 12):
            ws.cell(fila, col)._style = copy(ws.cell(fila, 17)._style)
        ws.merge_cells(start_row=11, start_column=col, end_row=12, end_column=col)
        ws.cell(11, col, titulo)
        ws.cell(11, col).alignment = Alignment(horizontal="center", vertical="center",
                                               wrap_text=True)
        ws.column_dimensions[ws.cell(1, col).column_letter].width = ancho
    ws.column_dimensions["A"].width = 15.5

    # ---- 8. parámetro del período -----------------------------------------
    rng = rango_dinamico(DESP, "A")
    etiquetas = [
        (7, "DÍAS DEL PERÍODO — escriba AUTO o un número fijo:", "AUTO"),
        (8, "Días con despacho detectados en la hoja origen:",
            "=IFERROR(SUMPRODUCT(1/COUNTIF({r},{r})),0)".format(r=rng)),
        (9, "Divisor aplicado al promedio diario:",
            '=IF(ISNUMBER($I$7),$I$7,IF($I$8>0,$I$8,20))'),
    ]
    for fila, texto, valor in etiquetas:
        ws.merge_cells(start_row=fila, start_column=5, end_row=fila, end_column=8)
        ws.cell(fila, 5, texto)
        ws.cell(fila, 5).font = Font(name="Calibri", size=8, bold=(fila == 9))
        ws.cell(fila, 5).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(fila, 9, valor)
        ws.cell(fila, 9).font = Font(name="Calibri", size=8, bold=True)
        ws.cell(fila, 9).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(7, 9).fill = PatternFill("solid", fgColor="FFFFF2CC")
    ws.cell(9, 9).fill = PatternFill("solid", fgColor="FFDDEBF7")
    wb.defined_names.add(DefinedName("DIAS_PERIODO", attr_text="'{h}'!$I$9".format(h=HOJA)))

    # ---- 9. control de cuadratura -----------------------------------------
    base = total + 2
    ws.cell(base, 3, "CONTROL DE CUADRATURA (la diferencia debe ser 0)")
    ws.cell(base, 3).font = Font(name="Calibri", size=9, bold=True, color="FF1F4E79")
    controles = [
        ("Stock total del cuadro vs. hoja INVENTARIO",
         "=P{t}".format(t=total), "=SUM({i}!E:E)".format(i=INV),
         "hay artículos en INVENTARIO que no están dados de alta en el cuadro"),
        ("Despacho total del cuadro vs. hoja DESPACHOS",
         "=H{t}".format(t=total), "=SUM('{d}'!F:F)".format(d=DESP),
         "hay artículos o canales en DESPACHOS que el cuadro no está sumando"),
    ]
    for k, (etiqueta, f_cua, f_ori, motivo) in enumerate(controles):
        rr = base + 1 + k
        ws.cell(rr, 3, etiqueta)
        ws.cell(rr, 7, f_cua)
        ws.cell(rr, 8, f_ori)
        ws.cell(rr, 9, "=G{r}-H{r}".format(r=rr))
        ws.cell(rr, 10, '=IF(I{r}=0,"OK","REVISAR: {m}")'.format(r=rr, m=motivo))
        for c in (7, 8, 9):
            ws.cell(rr, c).number_format = FMT_CANT

    # canales de DESPACHOS que no existen en REFERENCIA
    rr = base + 3
    rng_canal = rango_dinamico(DESP, "M")
    ws.cell(rr, 3, "Líneas de DESPACHOS cuyo CANAL no existe en REFERENCIA")
    ws.cell(rr, 9, "=IFERROR(SUMPRODUCT(--(COUNTIF({ref}!$F$2:$F$200,{rc})=0)),0)"
                   .format(ref=REF, rc=rng_canal))
    ws.cell(rr, 9).number_format = FMT_CANT
    ws.cell(rr, 10, '=IF(I{r}=0,"OK","REVISAR: complete esos canales en la hoja '
                    'REFERENCIA antes de la próxima carga")'.format(r=rr))

    for fila in range(base + 1, base + 4):
        for c in (3, 7, 8, 9, 10):
            ws.cell(fila, c).font = Font(name="Calibri", size=8,
                                         bold=(c == 10), color="FF404040")

    ws.cell(base + 5, 3,
            "Las bandas se calculan con SUMIFS sobre la columna A (CLAVE). "
            "Para agregar un artículo: inserte la fila dentro de su bloque, copie las "
            "fórmulas de la fila de arriba y escriba la misma CLAVE en la columna A.")
    ws.cell(base + 5, 3).font = Font(name="Calibri", size=8, italic=True, color="FF808080")

    # ---- 10. formato condicional sobre valores numéricos -------------------
    rango_dias = "M{a}:M{b} O{a}:O{b} Q{a}:Q{b}".format(a=ini, b=total)
    rango_estado = "R{a}:R{b}".format(a=ini, b=total)
    rango_alerta = "S{a}:S{b}".format(a=ini, b=total)

    dxf_rojo = dxf_ambar = dxf_verde = None
    for _, reglas in ws.conditional_formatting._cf_rules.items():
        for regla in reglas:                     # reutiliza los colores del archivo
            if regla.type == "expression" and regla.formula and regla.dxf:
                f = regla.formula[0]
                if "M13<8" in f:
                    dxf_rojo = regla.dxf
                elif "M13>=8" in f:
                    dxf_ambar = regla.dxf
                elif 'M13>150' in f:
                    dxf_verde = regla.dxf
    # Se eliminan SOLO las reglas del bloque de días (M/O/Q). Las de la leyenda
    # (L7, L8, L9) se conservan: dan color a los rótulos del encabezado.
    conservar = {cf: reglas for cf, reglas in ws.conditional_formatting._cf_rules.items()
                 if not any(c in str(cf.sqref) for c in ("M13", "O13", "Q13"))}
    ws.conditional_formatting._cf_rules.clear()
    ws.conditional_formatting._cf_rules.update(conservar)

    FILL_ROJO = PatternFill("solid", fgColor="FFFFC7CE")
    FILL_AMBAR = PatternFill("solid", fgColor="FFFFEB9C")
    FILL_VERDE = PatternFill("solid", fgColor="FFC6EFCE")
    FILL_GRIS = PatternFill("solid", fgColor="FFE7E6E6")
    FUENTE_GRIS = Font(color="FF7F7F7F", italic=True)
    # Las bandas usan texto blanco; sobre un semáforo claro sería ilegible.
    FUENTE_ROJO = Font(color="FF9C0006", bold=True)
    FUENTE_AMBAR = Font(color="FF9C6500", bold=True)
    FUENTE_VERDE = Font(color="FF006100", bold=True)

    def regla(formula, fill, font=None):
        r = FormulaRule(formula=[formula], stopIfTrue=False)
        r.dxf.fill = copy(fill)
        if font:
            r.dxf.font = copy(font)
        return r

    for f, fill, fnt in [
        ('AND($D13<>"",$I13=0)', FILL_GRIS, FUENTE_GRIS),                       # sin rotación
        ('AND($E13="Numerada",$I13>0,M13<%d)' % U_NUM_ROJO, FILL_ROJO, FUENTE_ROJO),
        ('AND($E13="Numerada",$I13>0,M13>=%d,M13<%d)' % (U_NUM_ROJO, U_NUM_AMBAR), FILL_AMBAR, FUENTE_AMBAR),
        ('AND($E13="Numerada",$I13>0,M13>=%d)' % U_NUM_AMBAR, FILL_VERDE, FUENTE_VERDE),
        ('AND($E13="Blanca",$I13>0,M13<=%d)' % U_BLA_ROJO, FILL_ROJO, FUENTE_ROJO),
        ('AND($E13="Blanca",$I13>0,M13>%d,M13<=%d)' % (U_BLA_ROJO, U_BLA_AMBAR), FILL_AMBAR, FUENTE_AMBAR),
        ('AND($E13="Blanca",$I13>0,M13>%d)' % U_BLA_AMBAR, FILL_VERDE, FUENTE_VERDE),
    ]:
        ws.conditional_formatting.add(rango_dias, regla(f, fill, fnt))

    for f, fill, fnt in [
        ('$R13="SIN ROTACIÓN"', FILL_GRIS, FUENTE_GRIS),
        ('$R13="CRÍTICO"', FILL_ROJO, FUENTE_ROJO),
        ('$R13="ALERTA"', FILL_AMBAR, FUENTE_AMBAR),
        ('$R13="OK"', FILL_VERDE, FUENTE_VERDE),
    ]:
        ws.conditional_formatting.add(rango_estado, regla(f, fill, fnt))
    ws.conditional_formatting.add(
        rango_alerta,
        regla('LEFT($S13,8)="PRODUCIR"',
              PatternFill("solid", fgColor="FFF8CBAD"),
              Font(color="FF833C00", bold=True)))

    # leyenda: aclarar el nuevo estado
    ws.cell(10, 10, "Sin rotación en el período = gris (antes se pintaba de verde)")
    ws.cell(10, 10).font = Font(name="Calibri", size=7, italic=True, color="FF808080")

    # ---- 11. NBSP en INVENTARIO -------------------------------------------
    wsi = wb[INV]
    for fila in wsi.iter_rows(min_row=2, max_row=wsi.max_row, min_col=1, max_col=4):
        for celda in fila:
            if isinstance(celda.value, str):
                nuevo = limpiar(celda.value)
                if nuevo != celda.value:
                    celda.value = nuevo

    wb.save(salida)

    # ---- 12. restaurar los valores numéricos originales -------------------
    n = restaurar_numeros(entrada, salida, wb.sheetnames, HOJAS_INTACTAS)

    print("OK -> %s" % salida)
    print("   detalle: filas %d..%d   TOTAL GENERAL: fila %d" % (ini, fin, total))
    print("   artículos: %d   bandas: %d" % (no, len(plan) - no))
    print("   valores numéricos restaurados al original: %d" % n)


# --------------------------------------------- restauración bit a bit
RE_CELDA = re.compile(r'<c\b[^>]*\br="([A-Z]+[0-9]+)"[^>]*/?>(?:(?!<c\b).)*?(?:</c>|(?<=/>))',
                      re.DOTALL)
RE_V = re.compile(r'<v>([^<]*)</v>')


RE_T = re.compile(r'\bt="([^"]+)"')


def _mapa_numericos(xml):
    """{ref: texto del <v>} de las celdas numéricas.

    Se excluyen las de texto (t="s"/"str"/"inlineStr"), booleanos y errores:
    su <v> es un índice a sharedStrings, que openpyxl regenera con otro orden.
    """
    out = {}
    for m in RE_CELDA.finditer(xml):
        frag = m.group(0)
        etiqueta = frag.split(">", 1)[0]
        t = RE_T.search(etiqueta)
        if t and t.group(1) != "n":
            continue
        v = RE_V.search(frag)
        if v and v.group(1):
            out[m.group(1)] = v.group(1)
    return out


def _mismo_numero(a, b):
    """True si las dos cadenas representan el mismo número (salvo truncamiento)."""
    try:
        fa, fb = float(a), float(b)
    except ValueError:
        return False
    if fa == fb:
        return True
    escala = max(abs(fa), abs(fb), 1.0)
    return abs(fa - fb) <= 1e-9 * escala


def restaurar_numeros(entrada, salida, nombres_hojas, hojas_intactas):
    """Devuelve a su texto original los números de las hojas no intervenidas.

    openpyxl serializa los flotantes con 15 dígitos significativos, lo que
    produce un double distinto (40,709999999999994 -> 40,70999999999999).
    Se restaura el texto exacto del archivo de entrada.
    """
    indices = [nombres_hojas.index(h) + 1 for h in hojas_intactas]
    cambios = 0
    with zipfile.ZipFile(entrada) as zin:
        originales = {i: zin.read("xl/worksheets/sheet%d.xml" % i).decode("utf-8")
                      for i in indices}

    tmp = salida + ".tmp"
    with zipfile.ZipFile(salida) as zsal, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zsal.infolist():
            datos = zsal.read(item.filename)
            m = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", item.filename)
            if m and int(m.group(1)) in originales:
                orig = _mapa_numericos(originales[int(m.group(1))])
                xml = datos.decode("utf-8")
                nuevos = _mapa_numericos(xml)
                reemplazos = {}
                for ref, val in nuevos.items():
                    previo = orig.get(ref)
                    if previo is None or previo == val:
                        continue
                    if _mismo_numero(previo, val):
                        reemplazos[ref] = previo
                if reemplazos:
                    def _sub(mm):
                        ref = mm.group(1)
                        if ref not in reemplazos:
                            return mm.group(0)
                        return RE_V.sub("<v>%s</v>" % reemplazos[ref], mm.group(0), count=1)
                    xml = RE_CELDA.sub(_sub, xml)
                    cambios += len(reemplazos)
                datos = xml.encode("utf-8")
            zout.writestr(item, datos)
    shutil.move(tmp, salida)
    return cambios


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
