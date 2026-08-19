"""
build_excel.py — Crea XML_Manager.xlsm usando Excel COM (pywin32).

Genera el archivo Excel completo con:
  - Hojas formateadas: Inicio, Config, FC_Recibidas, CR_Recibidas
  - Codigo VBA desde modMain.bas
  - Ribbon personalizado (customUI.xml) inyectado en el ZIP

Uso:
  python build_excel.py
  python build_excel.py --output C:/ruta/XML_Manager.xlsm
  python build_excel.py --no-abrir
"""
import argparse
import shutil
import sys
import zipfile
from pathlib import Path

# ─── Constantes de colores Excel (BGR interno de Excel) ─────────────────────
# Excel usa BGR invertido para .Color → 0xBBGGRR
C_AZUL_OSCURO  = 0x79_4E_1F   # RGB(31,78,121)   headers FC
C_CAFE_OSCURO  = 0x00_3C_83   # RGB(131,60,0)    headers CR
C_VERDE_OSCURO = 0x23_56_37   # RGB(55,86,35)    tab Config
C_BLANCO       = 0xFF_FF_FF
C_GRIS_CLARO   = 0xF2_F2_F2
C_GRIS_TEXTO   = 0x88_88_88
C_AZUL_LINK    = 0xFF_66_00   # naranja => usado para tips
C_TEXTO_DARK   = 0x59_59_59

# Alineaciones Excel
XL_LEFT   = -4131
XL_CENTER = -4108
XL_RIGHT  = -4152
XL_TOP    = -4160
XL_VCENTER = -4108

# Constantes Excel COM
XL_XLSM           = 52   # xlOpenXMLWorkbookMacroEnabled
VBA_STD_MODULE    = 1    # vbext_ct_StdModule

SCRIPT_DIR  = Path(__file__).parent
PROJECT_DIR = SCRIPT_DIR.parent
VBA_DIR     = SCRIPT_DIR / "vba"
CUI_DIR     = SCRIPT_DIR / "customUI"

VERSION = "1.0.0"


# ══════════════════════════════════════════════════════════════════════════════
# Helpers COM
# ══════════════════════════════════════════════════════════════════════════════

def rgb(r, g, b) -> int:
    """Convierte RGB a formato de color Excel (BGR)."""
    return r + (g * 256) + (b * 65536)


def set_style(cell, bold=False, italic=False, size=10,
              fg=None, bg=None, h_align=XL_LEFT, v_align=XL_VCENTER,
              wrap=False, border=False):
    cell.Font.Bold   = bold
    cell.Font.Italic = italic
    cell.Font.Size   = size
    if fg is not None:
        cell.Font.Color = fg
    if bg is not None:
        cell.Interior.Color   = bg
        cell.Interior.Pattern = 1  # xlSolid
    cell.HorizontalAlignment = h_align
    cell.VerticalAlignment   = v_align
    cell.WrapText            = wrap
    if border:
        cell.BorderAround(1, 2, 1)  # xlContinuous, xlThin, black


def header_row(ws, col_defs: list[tuple], row: int = 1,
               bg_color: int = C_AZUL_OSCURO, row_height: int = 32):
    """Escribe encabezados formateados y ajusta ancho de columnas."""
    for i, (nombre, ancho) in enumerate(col_defs, 1):
        c = ws.Cells(row, i)
        c.Value = nombre
        set_style(c, bold=True, size=9, fg=C_BLANCO, bg=bg_color,
                  h_align=XL_CENTER, wrap=True)
        ws.Columns(i).ColumnWidth = ancho
    ws.Rows(row).RowHeight = row_height


# ══════════════════════════════════════════════════════════════════════════════
# Hojas
# ══════════════════════════════════════════════════════════════════════════════

def crear_hoja_inicio(xl, wb):
    ws = wb.Sheets(1)
    ws.Name = "Inicio"
    ws.Tab.Color = C_AZUL_OSCURO

    xl.ActiveWindow.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 8
    ws.Columns("C").ColumnWidth = 70

    # Titulo
    c = ws.Range("B2")
    c.Value = "  SRI XML Manager"
    set_style(c, bold=True, size=22, fg=C_AZUL_OSCURO)
    ws.Rows(2).RowHeight = 40

    # Subtitulo
    c = ws.Range("B3")
    c.Value = "  Descarga automatica de comprobantes electronicos del SRI Ecuador"
    set_style(c, italic=True, size=11, fg=C_TEXTO_DARK)
    ws.Rows(3).RowHeight = 22

    # Borde inferior decorativo
    borde = ws.Range("B2:C3")
    borde.Borders(9).LineStyle = 1   # xlBottom
    borde.Borders(9).Color     = C_AZUL_OSCURO
    borde.Borders(9).Weight    = 3   # xlThick

    # Como usar
    ws.Cells(5, 2).Value = "COMO USAR"
    set_style(ws.Cells(5, 2), bold=True, size=11, fg=C_AZUL_OSCURO)

    pasos = [
        ("Paso 1",
         "Descargue el TXT de comprobantes del portal SRI\n"
         "         (srienlinea.sri.gob.ec > Servicios en Linea > "
         "Comprobantes Electronicos > Recibidos > Exportar TXT)"),
        ("Paso 2",
         "Haga clic en '1. Cargar TXT' en la pestana 'XML Manager' "
         "y seleccione el archivo TXT."),
        ("Paso 3",
         "Haga clic en '2. Descargar Recibidos'. El motor descargara automaticamente "
         "cada comprobante usando el Web Service oficial del SRI (sin navegador)."),
        ("Paso 4",
         "Haga clic en '3. Reporte Recibidos' para generar el reporte Excel "
         "con todos los datos desglosados."),
        ("Paso 5",
         "El reporte Excel se guarda en la carpeta de reportes y "
         "puede abrirse automaticamente."),
    ]

    for fila, (etiqueta, texto) in enumerate(pasos, 6):
        lbl = ws.Cells(fila, 2)
        lbl.Value = etiqueta + ":"
        set_style(lbl, bold=True, size=10, fg=C_AZUL_OSCURO)

        txt = ws.Cells(fila, 3)
        txt.Value = texto
        set_style(txt, size=10, wrap=True)
        ws.Rows(fila).RowHeight = 30

    # Pie
    pie = ws.Cells(13, 2)
    pie.Value = f"Version {VERSION}  |  Motor: sri_engine.exe  |  " \
                "Asegurese de que sri_engine.exe este en la misma carpeta que este archivo"
    set_style(pie, italic=True, size=8, fg=C_GRIS_TEXTO)
    ws.Rows(13).RowHeight = 18

    print("  Hoja 'Inicio' creada.")


def crear_hoja_config(xl, wb, ws_anterior):
    ws = wb.Sheets.Add(None, ws_anterior)
    ws.Name = "Config"
    ws.Tab.Color = C_VERDE_OSCURO

    xl.ActiveWindow.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 25
    ws.Columns("C").ColumnWidth = 65

    # Titulo
    c = ws.Range("B2")
    c.Value = "CONFIGURACION"
    set_style(c, bold=True, size=14, fg=C_AZUL_OSCURO)
    ws.Rows(2).RowHeight = 28

    # Filas de config: (fila, label, valor_default, hint)
    filas_base = [
        (4,  "RUC:",              "",            "Solo informativo. El motor lee el RUC del TXT del SRI."),
        (6,  "Archivo TXT:",      "",            "Se llena automaticamente al usar '1. Cargar TXT'."),
        (8,  "Carpeta XMLs:",     "",            "Carpeta donde se guardan los XMLs descargados."),
        (10, "Carpeta Reportes:", "",            "Carpeta donde se guardan los reportes Excel."),
        (12, "Ambiente:",         "PRODUCCION",  "PRODUCCION o PRUEBAS (afecta al endpoint SOAP del SRI)."),
        (14, "Clave SRI portal:", "",            "Clave de acceso a srienlinea.sri.gob.ec (Etapa 3: Descarga Automatica)."),
    ]

    for fila, label, valor, hint in filas_base:
        lbl = ws.Cells(fila, 2)
        lbl.Value = label
        set_style(lbl, bold=True, size=10, bg=C_GRIS_CLARO)
        ws.Rows(fila).RowHeight = 22

        val = ws.Cells(fila, 3)
        val.Value = valor
        set_style(val, size=10, bg=C_BLANCO, border=True)

        h = ws.Cells(fila + 1, 3)
        h.Value = "   " + hint
        set_style(h, italic=True, size=8, fg=C_GRIS_TEXTO)
        ws.Rows(fila + 1).RowHeight = 14

    # Validacion dropdown Ambiente
    dv = ws.Cells(12, 3).Validation
    try:
        dv.Delete()
        dv.Add(3, 1, 1, "PRODUCCION,PRUEBAS")
        dv.InCellDropdown = True
    except Exception:
        pass

    # ── Seccion Etapa 3: Descarga Automatica ─────────────────
    sep = ws.Cells(17, 2)
    sep.Value = "─── DESCARGA AUTOMATICA (Etapa 3) ───"
    set_style(sep, bold=True, size=9, fg=C_AZUL_OSCURO)
    ws.Rows(17).RowHeight = 18

    filas_auto = [
        (18, "Fecha Desde:",   "",  "Fecha inicio de busqueda en formato DD/MM/AAAA (ej: 01/01/2025)."),
        (20, "Fecha Hasta:",   "",  "Fecha fin de busqueda en formato DD/MM/AAAA (ej: 31/12/2025)."),
    ]
    for fila, label, valor, hint in filas_auto:
        lbl = ws.Cells(fila, 2)
        lbl.Value = label
        set_style(lbl, bold=True, size=10, bg=C_GRIS_CLARO)
        ws.Rows(fila).RowHeight = 22

        val = ws.Cells(fila, 3)
        val.Value = valor
        # Celda formateada como fecha
        try:
            val.NumberFormat = "DD/MM/YYYY"
        except Exception:
            pass
        set_style(val, size=10, bg=C_BLANCO, border=True)

        h = ws.Cells(fila + 1, 3)
        h.Value = "   " + hint
        set_style(h, italic=True, size=8, fg=C_GRIS_TEXTO)
        ws.Rows(fila + 1).RowHeight = 14

    # ── Seccion v2: Documentos Emitidos ──────────────────────
    sep2 = ws.Cells(24, 2)
    sep2.Value = "─── DOCUMENTOS EMITIDOS (v2) ───"
    set_style(sep2, bold=True, size=9, fg=C_AZUL_OSCURO)
    ws.Rows(24).RowHeight = 18

    filas_emit = [
        (26, "Archivo TXT Emitidos:", "", "Se llena automaticamente al usar '1. Cargar TXT Emitidos'."),
        (28, "Carpeta XMLs Emitidos:", "", "Carpeta donde se guardan los XMLs/PDFs emitidos (automatica)."),
    ]
    for fila, label, valor, hint in filas_emit:
        lbl = ws.Cells(fila, 2)
        lbl.Value = label
        set_style(lbl, bold=True, size=10, bg=C_GRIS_CLARO)
        ws.Rows(fila).RowHeight = 22

        val = ws.Cells(fila, 3)
        val.Value = valor
        set_style(val, size=10, bg=C_BLANCO, border=True)

        h = ws.Cells(fila + 1, 3)
        h.Value = "   " + hint
        set_style(h, italic=True, size=8, fg=C_GRIS_TEXTO)
        ws.Rows(fila + 1).RowHeight = 14

    # Nota al pie
    nota = ws.Cells(30, 2)
    nota.Value = "Nota: Los campos de ruta se actualizan automaticamente al usar los botones del ribbon."
    set_style(nota, italic=True, size=9, fg=C_GRIS_TEXTO)

    print("  Hoja 'Config' creada.")
    return ws


def crear_hoja_fc(xl, wb, ws_anterior):
    ws = wb.Sheets.Add(None, ws_anterior)
    ws.Name = "FC_Recibidas"
    ws.Tab.Color = C_AZUL_OSCURO

    xl.ActiveWindow.DisplayGridlines = False

    cols = [
        ("FECHA EMISION",      12), ("RUC EMISOR",         14),
        ("RAZON SOCIAL EMISOR",30), ("NUMERO",              16),
        ("RUC/CI COMPRADOR",   14), ("COMPRADOR",           28),
        ("BASE 0%",            11), ("BASE 12%",            11),
        ("BASE 14%",           11), ("BASE 15%",            11),
        ("BASE NO OBJETO",     13), ("BASE EXENTA",         11),
        ("SUBTOTAL",           11), ("DESCUENTO",           11),
        ("IVA",                11), ("PROPINA",             10),
        ("TOTAL",              11), ("CLAVE ACCESO",        50),
    ]
    header_row(ws, cols, bg_color=C_AZUL_OSCURO)

    # Congelar primera fila
    ws.Cells(2, 1).Select()
    xl.ActiveWindow.FreezePanes = True

    # Mensaje de datos vacíos
    c = ws.Cells(2, 1)
    c.Value = "Los datos se generan automaticamente con el boton '3. Reporte Recibidos' del ribbon."
    set_style(c, italic=True, size=9, fg=C_GRIS_TEXTO)

    print("  Hoja 'FC_Recibidas' creada.")
    return ws


def crear_hoja_multiruc(xl, wb, ws_anterior):
    ws = wb.Sheets.Add(None, ws_anterior)
    ws.Name = "Multi-RUC"
    ws.Tab.Color = C_VERDE_OSCURO

    xl.ActiveWindow.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 16   # RUC
    ws.Columns("C").ColumnWidth = 18   # Clave
    ws.Columns("D").ColumnWidth = 13   # Desde
    ws.Columns("E").ColumnWidth = 13   # Hasta
    ws.Columns("F").ColumnWidth = 16   # Estado
    ws.Columns("G").ColumnWidth = 10   # XMLs
    ws.Columns("H").ColumnWidth = 55   # Reporte

    # Titulo
    c = ws.Range("B2")
    c.Value = "DESCARGA MASIVA — MULTI-RUC"
    set_style(c, bold=True, size=14, fg=C_AZUL_OSCURO)
    ws.Rows(2).RowHeight = 28

    c2 = ws.Range("B3")
    c2.Value = "Configure hasta 20 RUCs. El boton 'Descarga Masiva' procesara cada uno automaticamente."
    set_style(c2, italic=True, size=9, fg=C_GRIS_TEXTO)
    ws.Rows(3).RowHeight = 16

    # Encabezados de tabla
    FILA_HDR = 4
    hdrs = [("RUC", 2), ("Clave SRI portal", 3), ("Fecha Desde", 4),
            ("Fecha Hasta", 5), ("Estado", 6), ("XMLs", 7), ("Reporte generado", 8)]
    for label, col in hdrs:
        c = ws.Cells(FILA_HDR, col)
        c.Value = label
        set_style(c, bold=True, size=9, fg=C_BLANCO, bg=C_AZUL_OSCURO, h_align=XL_CENTER)
    ws.Rows(FILA_HDR).RowHeight = 24

    # Filas de datos (5 a 24 = 20 RUCs max)
    for fila in range(5, 25):
        color_bg = C_BLANCO if fila % 2 == 1 else C_GRIS_CLARO
        for col in range(2, 9):
            c = ws.Cells(fila, col)
            set_style(c, size=10, bg=color_bg, border=True)
            if col in (4, 5):   # Fechas: formato DD/MM/YYYY
                try:
                    c.NumberFormat = "DD/MM/YYYY"
                except Exception:
                    pass
        ws.Rows(fila).RowHeight = 20

    # Numeracion automatica columna A (solo decorativa)
    for fila in range(5, 25):
        c = ws.Cells(fila, 1)
        c.Value = fila - 4
        set_style(c, size=8, fg=C_GRIS_TEXTO, h_align=XL_RIGHT)

    # Congelar fila de encabezado
    ws.Cells(5, 2).Select()
    xl.ActiveWindow.FreezePanes = True

    print("  Hoja 'Multi-RUC' creada.")
    return ws


def crear_hoja_cr(xl, wb, ws_anterior):
    ws = wb.Sheets.Add(None, ws_anterior)
    ws.Name = "CR_Recibidas"
    ws.Tab.Color = C_CAFE_OSCURO

    xl.ActiveWindow.DisplayGridlines = False

    cols = [
        ("FECHA EMISION",       12), ("RUC RETENEDOR",      14),
        ("RAZON RETENEDOR",     30), ("NUMERO RETENCION",   18),
        ("RUC/CI RETENIDO",     14), ("RAZON RETENIDO",     28),
        ("PERIODO FISCAL",      13), ("COD DOC SUSTENTO",   16),
        ("NUM DOC SUSTENTO",    20), ("FECHA DOC SUSTENTO", 16),
        ("TIPO IMPUESTO",       13), ("COD RETENCION",      13),
        ("BASE IMPONIBLE",      13), ("% RETENCION",        11),
        ("VALOR RETENIDO",      13), ("CLAVE ACCESO",       50),
    ]
    header_row(ws, cols, bg_color=C_CAFE_OSCURO)

    ws.Cells(2, 1).Select()
    xl.ActiveWindow.FreezePanes = True

    c = ws.Cells(2, 1)
    c.Value = "Los datos se generan automaticamente con el boton '3. Reporte Recibidos' del ribbon."
    set_style(c, italic=True, size=9, fg=C_GRIS_TEXTO)

    print("  Hoja 'CR_Recibidas' creada.")
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# VBA — Inyeccion directa en ZIP (sin Excel COM)
# ══════════════════════════════════════════════════════════════════════════════

def inyectar_vba_bin(xlsm_path: Path):
    """
    Genera vbaProject.bin desde modMain.bas e inyecta en el ZIP del xlsm.

    No usa Excel COM (que requiere Trust Center habilitado).
    El vbaProject.bin se crea puramente en Python usando el generador propio.
    Excel recompila el p-code al abrir (normal para archivos de otra version).
    """
    sys.path.insert(0, str(SCRIPT_DIR))
    try:
        from create_vba_project import generar_vba_project_bin
    except ImportError as e:
        print(f"  Error importando create_vba_project: {e}")
        return False

    vba_path = VBA_DIR / "modMain.bas"
    if not vba_path.exists():
        print(f"  Advertencia: {vba_path} no encontrado. VBA omitido.")
        return False

    # Leer codigo VBA, quitar atributo VB_Name
    vba_code = vba_path.read_text(encoding="utf-8")
    lineas   = [l for l in vba_code.splitlines()
                if not l.strip().startswith("Attribute VB_Name")]
    vba_code = "\n".join(lineas)

    try:
        print("  Generando vbaProject.bin desde modMain.bas...")
        vba_bin = generar_vba_project_bin({"modMain": vba_code})
        print(f"  vbaProject.bin: {len(vba_bin)} bytes")
    except Exception as e:
        print(f"  Error generando vbaProject.bin: {e}")
        return False

    # Leer el ZIP actual
    with zipfile.ZipFile(xlsm_path, "r") as zin:
        archivos = {n: zin.read(n) for n in zin.namelist()}

    # Inyectar xl/vbaProject.bin
    archivos["xl/vbaProject.bin"] = vba_bin

    # Actualizar [Content_Types].xml
    ct = archivos.get("[Content_Types].xml", b"").decode("utf-8")
    if "vbaProject" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/xl/vbaProject.bin" '
            'ContentType="application/vnd.ms-office.vbaProject"/>\n</Types>'
        )
        archivos["[Content_Types].xml"] = ct.encode("utf-8")

    # Actualizar xl/_rels/workbook.xml.rels
    rels_key = "xl/_rels/workbook.xml.rels"
    rels = archivos.get(rels_key, b"").decode("utf-8")
    if "vbaProject" not in rels:
        rels = rels.replace(
            "</Relationships>",
            '<Relationship Id="rIdVBA" '
            'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
            'Target="vbaProject.bin"/>\n</Relationships>'
        )
        archivos[rels_key] = rels.encode("utf-8")

    # Reescribir ZIP
    tmp = xlsm_path.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for nombre, contenido in archivos.items():
            zout.writestr(nombre, contenido)

    xlsm_path.unlink()
    tmp.rename(xlsm_path)

    print("  VBA inyectado correctamente en el ZIP.")
    return True


# ══════════════════════════════════════════════════════════════════════════════
# CustomUI (ribbon) — inyeccion en ZIP
# ══════════════════════════════════════════════════════════════════════════════

def inyectar_ribbon(xlsm_path: Path):
    """
    Inyecta customUI.xml (2006) Y customUI14.xml (2009) en el ZIP del .xlsm.
    Ademas elimina la etiqueta de clasificacion corporativa MIP (docMetadata/LabelInfo.xml)
    que bloquea el custom ribbon en entornos corporativos.
    """
    cui_path   = CUI_DIR / "customUI.xml"
    cui14_path = CUI_DIR / "customUI14.xml"

    if not cui_path.exists():
        print(f"  Advertencia: customUI.xml no encontrado en {cui_path}")
        return

    backup = xlsm_path.with_suffix(".bak")
    shutil.copy2(xlsm_path, backup)

    try:
        with zipfile.ZipFile(xlsm_path, "r") as zin:
            archivos = {n: zin.read(n) for n in zin.namelist()}

        ct   = archivos.get("[Content_Types].xml", b"").decode("utf-8")
        rels = archivos.get("_rels/.rels", b"").decode("utf-8")

        # ── Eliminar etiqueta MIP corporativa (bloquea ribbon) ────────────
        # docMetadata/LabelInfo.xml viene del seed template corporativo y
        # hace que Excel aplique restricciones que impiden cargar custom UI.
        mip_keys = [k for k in archivos if k.startswith("docMetadata/")]
        for k in mip_keys:
            del archivos[k]
            print(f"  Eliminando etiqueta corporativa: {k}")

        # Limpiar relacion MIP de _rels/.rels
        # Usamos [^>]* para tolerar rutas con '/' dentro del atributo Target
        import re
        rels = re.sub(
            r'<Relationship\s[^>]*classificationlabels[^>]*/>\s*',
            '',
            rels
        )
        # Limpiar Override MIP de [Content_Types].xml
        ct = re.sub(
            r'<Override\s[^>]*classificationlabels[^>]*/>\s*',
            '',
            ct
        )

        # ── customUI (Office 2006 / Excel 2007–365) ──────────────────────────
        # IMPORTANTE: Solo inyectamos el customUI (2006), NO el customUI14.
        # Razon: Si ambos estan presentes, Excel 2010+ carga SOLO customUI14 e
        # ignora el 2006 completamente (regla MS-OI29500). Si customUI14 tiene
        # cualquier problema, el ribbon no aparece y no hay mensaje de error.
        # Usando solo customUI (2006), Excel 365 hace fallback a el cuando no
        # existe customUI14 — exactamente como funciona TribuExcel v5.3.
        #
        # Eliminar cualquier customUI14 que hubiera quedado de builds previos
        for k in list(archivos.keys()):
            if "customUI14" in k:
                del archivos[k]
                print(f"  Eliminando customUI14 residual: {k}")
        ct = re.sub(r'<Override\s[^>]*customUI14[^>]*/>\s*', '', ct)
        rels = re.sub(r'<Relationship\s[^>]*customUI14[^>]*/>\s*', '', rels)
        rels = re.sub(r'<Relationship\s[^>]*2007/relationships/ui/extensibility[^>]*/>\s*', '', rels)

        # Inyectar customUI (2006) — replicando exactamente la estructura de TribuExcel:
        # - Usar Default Extension="xml" con ContentType="application/xml" (no Override especifico)
        # - NO inyectar Override para customUI/customUI.xml
        # - Agregar customUI/_rels/customUI.xml.rels (archivo de relaciones, aunque vacio)
        archivos["customUI/customUI.xml"] = cui_path.read_bytes()

        # Agregar Default para xml si no existe (como TribuExcel)
        if 'Extension="xml"' not in ct:
            ct = ct.replace("</Types>",
                '<Default Extension="xml" ContentType="application/xml"/>\n</Types>')

        # Eliminar cualquier Override especifico para customUI que pudiera existir
        ct = re.sub(r'<Override\s[^>]*customUI[^>]*/>\s*', '', ct)

        # Agregar relaciones de customUI (imagenes — vacio porque usamos imageMso)
        archivos["customUI/_rels/customUI.xml.rels"] = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'</Relationships>'
        )

        # Relacion en _rels/.rels — PRIMERO en la lista (como TribuExcel)
        if 'Target="customUI/customUI.xml"' not in rels:
            # Insertar al inicio, despues de la apertura del tag Relationships
            rels = rels.replace(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="R1a2b3c4d5e6f7890" '
                'Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" '
                'Target="customUI/customUI.xml"/>'
            )

        archivos["[Content_Types].xml"] = ct.encode("utf-8")
        archivos["_rels/.rels"]         = rels.encode("utf-8")

        # Reescribir ZIP
        tmp = xlsm_path.with_suffix(".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for nombre, contenido in archivos.items():
                zout.writestr(nombre, contenido)

        xlsm_path.unlink()
        tmp.rename(xlsm_path)
        backup.unlink(missing_ok=True)

        has14 = cui14_path.exists()
        print(f"  Ribbon inyectado: customUI {'+ customUI14 ' if has14 else ''}OK.")

    except Exception as e:
        print(f"  Error inyectando ribbon: {e}")
        if backup.exists():
            shutil.copy2(backup, xlsm_path)
            backup.unlink(missing_ok=True)


# ══════════════════════════════════════════════════════════════════════════════
# VBA via COM — inyección mientras el workbook está abierto
# ══════════════════════════════════════════════════════════════════════════════

def _habilitar_vbom():
    """Activa 'Confiar en el acceso al modelo de objetos VBA' en el registro."""
    try:
        import winreg
        # Intentar con version 16.0 (Office 2016/2019/365)
        for ver in ("16.0", "15.0", "14.0"):
            try:
                key = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    rf"Software\Microsoft\Office\{ver}\Excel\Security",
                    0, winreg.KEY_ALL_ACCESS)
                winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
                winreg.CloseKey(key)
                return True
            except FileNotFoundError:
                continue
    except Exception:
        pass
    return False


def _leer_vba_code() -> str | None:
    """Lee modMain.bas y elimina la línea Attribute VB_Name."""
    vba_path = VBA_DIR / "modMain.bas"
    if not vba_path.exists():
        return None
    code = vba_path.read_text(encoding="utf-8")
    lines = [l for l in code.splitlines()
             if not l.strip().startswith("Attribute VB_Name")]
    return "\r\n".join(lines)


def inyectar_vba_via_com(xl, wb) -> bool:
    """
    Intenta inyectar modMain en el workbook abierto usando Excel COM.

    Estrategia 1: wb.VBProject (directo, requiere HasVBProject=True)
    Estrategia 2: xl.VBE.VBProjects (via IDE, puede funcionar aunque
                  HasVBProject=False en Excel 365)

    Retorna True si el VBA fue inyectado correctamente.
    """
    _habilitar_vbom()
    xl.AutomationSecurity = 1  # msoAutomationSecurityLow

    vba_code = _leer_vba_code()
    if vba_code is None:
        print("  modMain.bas no encontrado.")
        return False

    def _add_module(proj):
        """Elimina módulos estándar existentes y agrega modMain."""
        # Eliminar módulos estándar previos (tipo 1)
        to_remove = [proj.VBComponents.Item(i)
                     for i in range(1, proj.VBComponents.Count + 1)
                     if proj.VBComponents.Item(i).Type == 1]
        for comp in to_remove:
            try:
                proj.VBComponents.Remove(comp)
            except Exception:
                pass
        comp = proj.VBComponents.Add(1)   # vbext_ct_StdModule
        comp.Name = "modMain"
        comp.CodeModule.AddFromString(vba_code)

    # ── Estrategia 1: wb.VBProject ──────────────────────────
    try:
        vbp = wb.VBProject
        if vbp is not None:
            _add_module(vbp)
            print("  VBA inyectado via wb.VBProject.")
            return True
    except Exception as e:
        pass  # Puede fallar con pywintypes.com_error si AccessVBOM=0

    # ── Estrategia 2: xl.VBE.VBProjects ─────────────────────
    try:
        vbe = xl.VBE
        if vbe is None:
            return False

        # Buscar el proyecto correspondiente al workbook abierto
        proj = None
        wb_name = wb.Name  # ej: "Libro1" o nombre del seed

        for i in range(1, vbe.VBProjects.Count + 1):
            try:
                p = vbe.VBProjects.Item(i)
                if p.Name == wb_name or p.Name == Path(wb_name).stem:
                    proj = p
                    break
            except Exception:
                continue

        # Si no encontramos por nombre, tomar el primero que no sea PERSONAL
        if proj is None:
            for i in range(1, vbe.VBProjects.Count + 1):
                try:
                    p = vbe.VBProjects.Item(i)
                    fname = ""
                    try:
                        fname = p.Filename
                    except Exception:
                        pass
                    if "PERSONAL" not in fname.upper():
                        proj = p
                        break
                except Exception:
                    continue

        if proj is not None:
            _add_module(proj)
            print("  VBA inyectado via xl.VBE.VBProjects.")
            return True

    except Exception as e:
        print(f"  VBE strategy failed: {e}")

    return False


def usar_seed_template(xl, seed_path: Path, output_path: Path) -> bool:
    """
    Construye el workbook usando seed.xlsm como base.
    El seed ya tiene un VBProject válido, por lo que COM puede
    inyectar el VBA sin restricciones.

    Retorna True si el archivo fue generado correctamente con VBA via COM.
    """
    print(f"  Usando seed template: {seed_path.name}")
    wb = None
    try:
        wb = xl.Workbooks.Open(str(seed_path))

        # ── Limpiar hojas del seed ────────────────────────────
        # Estrategia: eliminar DESDE EL FINAL hasta quedarnos solo con Sheets(1),
        # luego crear_hoja_inicio lo renombra. Nunca tocamos la última hoja visible.
        while wb.Sheets.Count > 1:
            wb.Sheets(wb.Sheets.Count).Delete()

        # Sheets(1) es ahora la única hoja (nombre original del seed: ej. "Hoja1")
        # crear_hoja_inicio la renombra a "Inicio"
        crear_hoja_inicio(xl, wb)
        ws_inicio = wb.Sheets("Inicio")
        ws_inicio.Activate()

        ws_config    = crear_hoja_config(xl, wb, ws_inicio)
        ws_fc        = crear_hoja_fc(xl, wb, ws_config)
        ws_cr        = crear_hoja_cr(xl, wb, ws_fc)
        _            = crear_hoja_multiruc(xl, wb, ws_cr)

        ws_inicio.Activate()

        # ── Inyectar VBA (HasVBProject=True en el seed) ───────
        vba_ok = inyectar_vba_via_com(xl, wb)
        if vba_ok:
            print("  VBA inyectado correctamente via COM.")
        else:
            print("  Advertencia: VBA no inyectado via COM desde seed.")

        # ── Guardar como XML_Manager.xlsm ─────────────────────
        try:
            if output_path.exists():
                output_path.unlink()
        except PermissionError:
            raise PermissionError(
                f"No se puede sobreescribir '{output_path.name}' — "
                "cierre el archivo en Excel antes de ejecutar el build.")
        wb.SaveAs(str(output_path), XL_XLSM)
        wb.Close(False)
        return vba_ok

    except Exception as e:
        print(f"  Error con seed template: {e}")
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        return False


# ══════════════════════════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="Crea XML_Manager.xlsm")
    parser.add_argument("--output",   default=str(SCRIPT_DIR / "XML_Manager.xlsm"))
    parser.add_argument("--no-abrir", action="store_true")
    args = parser.parse_args()

    output_path  = Path(args.output)
    templates_dir = SCRIPT_DIR / "templates"
    # Acepta cualquier .xlsm en la carpeta templates/ como seed
    seed_candidates = sorted(templates_dir.glob("*.xlsm")) if templates_dir.exists() else []
    seed_path = seed_candidates[0] if seed_candidates else (SCRIPT_DIR / "templates" / "seed.xlsm")

    print()
    print("=" * 55)
    print(f"  SRI XML Manager v{VERSION} — Build Excel")
    print("=" * 55)

    try:
        import win32com.client as win32
    except ImportError:
        print("ERROR: pywin32 no instalado.")
        print("Ejecute: pip install pywin32")
        sys.exit(1)

    xl  = None
    wb  = None
    vba_via_com = False

    try:
        # ── [1] Iniciar Excel COM ────────────────────────────
        print("\n[1/4] Iniciando Excel COM...")
        # Habilitar AccesVBOM ANTES de lanzar Excel (se lee al iniciar)
        vbom_ok = _habilitar_vbom()
        if vbom_ok:
            print("  AccessVBOM habilitado en registro.")
        xl = win32.Dispatch("Excel.Application")
        xl.Visible        = False
        xl.DisplayAlerts  = False
        xl.ScreenUpdating = False
        xl.AutomationSecurity = 1  # msoAutomationSecurityLow
        print("  Excel COM listo.")

        # ── [2] Crear workbook (o usar seed) ─────────────────
        print("\n[2/4] Creando hojas...")

        if seed_path.exists():
            # ── Ruta seed: usa template con VBProject válido ─
            print(f"  Seed template encontrado: {seed_path.name}")
            vba_via_com = usar_seed_template(xl, seed_path, output_path)
            if vba_via_com:
                print("  Hojas y VBA creados desde seed.")
            # El seed_path path ya guardó y cerró; saltamos al paso ribbon
        else:
            # ── Ruta normal: workbook nuevo ──────────────────
            print("  (sin seed template — intentando VBE injection)")
            wb = xl.Workbooks.Add()

            crear_hoja_inicio(xl, wb)
            ws_inicio = wb.Sheets("Inicio")
            ws_inicio.Activate()

            ws_config = crear_hoja_config(xl, wb, ws_inicio)
            ws_fc     = crear_hoja_fc(xl, wb, ws_config)
            ws_cr     = crear_hoja_cr(xl, wb, ws_fc)
            _         = crear_hoja_multiruc(xl, wb, ws_cr)

            # Eliminar hojas por defecto vacías
            for sh in list(wb.Sheets):
                if sh.Name not in ("Inicio", "Config", "FC_Recibidas", "CR_Recibidas"):
                    sh.Delete()

            ws_inicio.Activate()

            # Intentar VBE injection ANTES de guardar
            print("\n[3/4] Inyectando VBA via COM...")
            vba_via_com = inyectar_vba_via_com(xl, wb)
            if not vba_via_com:
                print("  VBE injection no disponible — se usara binario.")

            # Guardar como .xlsm
            print(f"\n[3/4] Guardando: {output_path.name}...")
            if output_path.exists():
                output_path.unlink()
            wb.SaveAs(str(output_path), XL_XLSM)
            wb.Close(False)
            wb = None
            print("  Guardado correctamente.")

        xl.Quit()
        xl = None

    except Exception as e:
        print(f"\nERROR en Excel COM: {e}")
        if wb is not None:
            try: wb.Close(False)
            except: pass
        if xl is not None:
            try: xl.Quit()
            except: pass
        raise

    # ── [4] Inyectar ribbon + VBA fallback en ZIP ────────────
    print("\n[4/4] Inyectando ribbon en ZIP...")
    inyectar_ribbon(output_path)

    if not vba_via_com:
        print("  Inyectando vbaProject.bin (metodo binario)...")
        vba_bin_ok = inyectar_vba_bin(output_path)
        if not vba_bin_ok:
            print()
            print("  AVISO: VBA no pudo ser inyectado automaticamente.")
            print("  Para habilitar VBA confiablemente, cree el seed template:")
            print()
            print("    1. Abra Excel → nuevo libro en blanco")
            print("    2. Alt+F11 → Insertar → Modulo → escriba: Sub Tmp(): End Sub")
            print("    3. Cierre el editor y guarde como:")
            print(f"       {seed_path}")
            print("    4. Ejecute build_excel.py de nuevo")
            print()

    # ── Resumen ──────────────────────────────────────────────
    size_kb = output_path.stat().st_size // 1024
    metodo_vba = "COM (confiable)" if vba_via_com else "binario (CFB)"
    print()
    print("=" * 55)
    print(f"  BUILD EXITOSO")
    print("=" * 55)
    print(f"  Archivo  : {output_path}")
    print(f"  Tamano   : {size_kb} KB")
    print(f"  VBA via  : {metodo_vba}")
    print()
    print("  Para distribuir al cliente:")
    print("    XML_Manager_v1.0.zip")
    print("      |- XML_Manager.xlsm  <-- este archivo")
    print("      |- sri_engine.exe    <-- compilar con build.ps1")
    print()

    if not args.no_abrir:
        resp = input("  Abrir el archivo ahora? [S/n]: ").strip().upper()
        if resp in ("", "S", "SI", "Y", "YES"):
            import os
            os.startfile(str(output_path))


if __name__ == "__main__":
    main()
