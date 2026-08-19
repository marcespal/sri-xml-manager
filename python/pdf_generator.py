"""
pdf_generator.py — Genera el RIDE en PDF a partir de un objeto Factura.

Diseño minimalista, sin logo, basado en el RIDE oficial SRI.  Usa ReportLab
(pure-Python, se empaqueta sin problemas con PyInstaller).
"""
from __future__ import annotations

from pathlib import Path
from typing import List, Optional

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether,
)
from reportlab.pdfgen import canvas as rl_canvas

from utils import get_logger
from xml_parser import Factura, Retencion

logger = get_logger("pdf_generator")

# ── Tabla SRI: identificación del comprador ──────────────────────────────────
_TIPO_ID = {
    "04": "RUC",
    "05": "CÉDULA",
    "06": "PASAPORTE",
    "07": "VENTA A CONSUMIDOR FINAL",
    "08": "IDENTIFICACIÓN DEL EXTERIOR",
}

# ── Tabla SRI: formas de pago (simplificada) ─────────────────────────────────
_FORMA_PAGO = {
    "01": "Sin utilización del sistema financiero",
    "15": "Compensación de deudas",
    "16": "Tarjeta de débito",
    "17": "Dinero electrónico",
    "18": "Tarjeta prepago",
    "19": "Tarjeta de crédito",
    "20": "Otros con utilización del sistema financiero",
    "21": "Endoso de títulos",
}


# ── Estilos de párrafo ───────────────────────────────────────────────────────

def _estilos():
    base = getSampleStyleSheet()["Normal"]
    return {
        "titulo": ParagraphStyle("titulo", parent=base, fontName="Helvetica-Bold",
                                 fontSize=14, leading=16, alignment=1),
        "h":      ParagraphStyle("h", parent=base, fontName="Helvetica-Bold",
                                 fontSize=9, leading=11, textColor=colors.white),
        "label":  ParagraphStyle("label", parent=base, fontName="Helvetica-Bold",
                                 fontSize=8, leading=10),
        "val":    ParagraphStyle("val", parent=base, fontName="Helvetica",
                                 fontSize=8, leading=10),
        "small":  ParagraphStyle("small", parent=base, fontName="Helvetica",
                                 fontSize=7, leading=8.5),
        "smallB": ParagraphStyle("smallB", parent=base, fontName="Helvetica-Bold",
                                 fontSize=7, leading=8.5),
        "clave":  ParagraphStyle("clave", parent=base, fontName="Courier-Bold",
                                 fontSize=9, leading=11, alignment=1),
    }


def _moneda(v: float) -> str:
    """Formato '1,234.56' con coma de miles y punto decimal."""
    try:
        return f"{float(v):,.2f}"
    except Exception:
        return str(v)


# ── Bloques del PDF ──────────────────────────────────────────────────────────

def _bloque_encabezado(f: Factura, S: dict) -> Table:
    """Encabezado: emisor (izquierda) + datos del comprobante (derecha)."""
    # Bloque izquierdo: emisor
    izq = [
        [Paragraph(f.razon_social or "—", S["titulo"])],
        [Paragraph(f"<b>Nombre Comercial:</b> {f.nombre_comercial or '—'}", S["val"])],
        [Paragraph(f"<b>Dirección Matriz:</b> {f.dir_matriz or '—'}", S["val"])],
        [Paragraph(f"<b>Dirección Sucursal:</b> {f.dir_establecimiento or '—'}", S["val"])],
        [Paragraph(f"<b>Obligado a Contabilidad:</b> {f.obligado_contabilidad or 'NO'}", S["val"])],
    ]
    if f.contribuyente_especial:
        izq.append([Paragraph(
            f"<b>Contribuyente Especial Nro:</b> {f.contribuyente_especial}", S["val"])])

    t_izq = Table(izq, colWidths=[95*mm])
    t_izq.setStyle(TableStyle([
        ("LEFTPADDING",   (0,0), (-1,-1), 0),
        ("RIGHTPADDING",  (0,0), (-1,-1), 0),
        ("TOPPADDING",    (0,0), (-1,-1), 1),
        ("BOTTOMPADDING", (0,0), (-1,-1), 1),
    ]))

    # Título según tipo de comprobante
    titulo_doc = {
        "01": "FACTURA", "03": "LIQUIDACIÓN DE COMPRA",
        "04": "NOTA DE CRÉDITO", "05": "NOTA DE DÉBITO",
    }.get(f.cod_doc, "COMPROBANTE")

    # Bloque derecho: RUC, tipo doc, número, autorización, clave
    der = [
        [Paragraph(f"<b>R.U.C.:</b> {f.ruc}", S["val"])],
        [Paragraph(f"<b>{titulo_doc}</b>", S["titulo"])],
        [Paragraph(f"<b>Nº:</b> {f.numero}", S["val"])],
        [Paragraph(f"<b>NÚMERO DE AUTORIZACIÓN</b>", S["smallB"])],
        [Paragraph(f.clave_acceso or "—", S["clave"])],
        [Paragraph(f"<b>FECHA Y HORA DE AUTORIZACIÓN:</b> {f.fecha_autorizacion or '—'}", S["small"])],
        [Paragraph(f"<b>AMBIENTE:</b> {'PRODUCCIÓN' if f.ambiente == '2' else 'PRUEBAS'}", S["small"])],
        [Paragraph("<b>EMISIÓN:</b> NORMAL", S["small"])],
        [Paragraph("<b>CLAVE DE ACCESO:</b>", S["smallB"])],
        [Paragraph(f.clave_acceso or "—", S["clave"])],
    ]
    t_der = Table(der, colWidths=[90*mm])
    t_der.setStyle(TableStyle([
        ("BOX",           (0,0), (-1,-1), 0.7, colors.black),
        ("LEFTPADDING",   (0,0), (-1,-1), 4),
        ("RIGHTPADDING",  (0,0), (-1,-1), 4),
        ("TOPPADDING",    (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))

    tabla = Table([[t_izq, t_der]], colWidths=[95*mm, 90*mm])
    tabla.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING",   (0,0), (-1,-1), 0),
        ("RIGHTPADDING",  (0,0), (-1,-1), 0),
        ("TOPPADDING",    (0,0), (-1,-1), 0),
        ("BOTTOMPADDING", (0,0), (-1,-1), 0),
    ]))
    return tabla


def _bloque_comprador(f: Factura, S: dict) -> Table:
    tipo = _TIPO_ID.get(f.tipo_identificacion_comprador, f.tipo_identificacion_comprador)
    # Etiqueta según rol: 'Proveedor' para liquidación de compra, 'Comprador' resto
    rol_lbl = "Proveedor" if getattr(f, "rol_contraparte", "comprador") == "proveedor" else "Razón Social / Nombres y Apellidos"
    filas = [[
        Paragraph(f"<b>{rol_lbl}:</b> {f.razon_social_comprador or '—'}", S["val"]),
        Paragraph(f"<b>Identificación ({tipo}):</b> {f.identificacion_comprador or '—'}", S["val"]),
    ],[
        Paragraph(f"<b>Fecha Emisión:</b> {f.fecha_emision or '—'}", S["val"]),
        Paragraph(f"<b>Dirección:</b> {f.direccion_comprador or '—'}", S["val"]),
    ]]
    # NC / ND: documento que modifica
    if f.cod_doc_modificado or f.num_doc_modificado:
        filas.append([
            Paragraph(f"<b>Comprobante que modifica:</b> {f.cod_doc_modificado}-{f.num_doc_modificado}", S["val"]),
            Paragraph(f"<b>Fecha doc. sustento:</b> {f.fecha_doc_sustento or '—'}", S["val"]),
        ])
    if f.motivo:
        filas.append([
            Paragraph(f"<b>Motivo:</b> {f.motivo}", S["val"]),
            Paragraph("", S["val"]),
        ])
    t = Table(filas, colWidths=[110*mm, 75*mm])
    t.setStyle(TableStyle([
        ("BOX",           (0,0), (-1,-1), 0.7, colors.black),
        ("INNERGRID",     (0,0), (-1,-1), 0.3, colors.grey),
        ("LEFTPADDING",   (0,0), (-1,-1), 4),
        ("RIGHTPADDING",  (0,0), (-1,-1), 4),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    return t


def _bloque_detalles(f: Factura, S: dict) -> Table:
    """Tabla de líneas: cód, descripción, cantidad, P.Unit, descuento, subtotal."""
    enc = [Paragraph(t, S["h"]) for t in
           ["Cód. Principal", "Descripción", "Cant.", "P.Unit.", "Dscto.", "Subtotal"]]
    filas = [enc]
    for d in f.detalles:
        filas.append([
            Paragraph(d.codigo_principal or "", S["small"]),
            Paragraph(d.descripcion or "", S["small"]),
            Paragraph(_moneda(d.cantidad), S["small"]),
            Paragraph(_moneda(d.precio_unitario), S["small"]),
            Paragraph(_moneda(d.descuento), S["small"]),
            Paragraph(_moneda(d.total_sin_impuesto), S["small"]),
        ])
    t = Table(filas, colWidths=[20*mm, 80*mm, 18*mm, 22*mm, 18*mm, 27*mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), colors.HexColor("#2C3E50")),
        ("ALIGN",         (2,1), (-1,-1), "RIGHT"),
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
        ("BOX",           (0,0), (-1,-1), 0.7, colors.black),
        ("INNERGRID",     (0,0), (-1,-1), 0.3, colors.grey),
        ("LEFTPADDING",   (0,0), (-1,-1), 3),
        ("RIGHTPADDING",  (0,0), (-1,-1), 3),
        ("TOPPADDING",    (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    return t


def _bloque_totales(f: Factura, S: dict) -> Table:
    """Bloque derecho de totales (subtotales por tarifa, IVA, total)."""
    filas = []
    def linea(lbl, val):
        filas.append([Paragraph(f"<b>{lbl}</b>", S["small"]),
                      Paragraph(_moneda(val), S["small"])])
    if f.base_iva_15:  linea("Subtotal 15%",        f.base_iva_15)
    if f.base_iva_12:  linea("Subtotal 12%",        f.base_iva_12)
    if f.base_iva_14:  linea("Subtotal 14%",        f.base_iva_14)
    if f.base_iva_0:   linea("Subtotal 0%",         f.base_iva_0)
    if f.base_exenta:  linea("Subtotal Exento IVA", f.base_exenta)
    if f.base_no_objeto: linea("Subtotal No Objeto IVA", f.base_no_objeto)
    linea("Subtotal sin impuestos", f.total_sin_impuestos)
    linea("Descuento",              f.total_descuento)
    if f.valor_iva:  linea("IVA",   f.valor_iva)
    if f.propina:    linea("Propina", f.propina)
    filas.append([
        Paragraph("<b>VALOR TOTAL</b>", S["smallB"]),
        Paragraph(f"<b>{_moneda(f.importe_total)}</b>", S["smallB"]),
    ])

    t = Table(filas, colWidths=[35*mm, 25*mm], hAlign="RIGHT")
    t.setStyle(TableStyle([
        ("BOX",           (0,0), (-1,-1), 0.7, colors.black),
        ("INNERGRID",     (0,0), (-1,-1), 0.3, colors.grey),
        ("BACKGROUND",    (0,-1), (-1,-1), colors.HexColor("#ECF0F1")),
        ("ALIGN",         (1,0), (1,-1), "RIGHT"),
        ("LEFTPADDING",   (0,0), (-1,-1), 3),
        ("RIGHTPADDING",  (0,0), (-1,-1), 3),
        ("TOPPADDING",    (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    return t


def _bloque_info_adicional(f: Factura, S: dict) -> Optional[Table]:
    if not f.info_adicional:
        return None
    filas = [[Paragraph("<b>Información Adicional</b>", S["h"]), Paragraph("", S["h"])]]
    for k, v in f.info_adicional.items():
        filas.append([
            Paragraph(k, S["small"]),
            Paragraph(v, S["small"]),
        ])
    t = Table(filas, colWidths=[60*mm, 125*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), colors.HexColor("#2C3E50")),
        ("BOX",           (0,0), (-1,-1), 0.7, colors.black),
        ("INNERGRID",     (0,0), (-1,-1), 0.3, colors.grey),
        ("LEFTPADDING",   (0,0), (-1,-1), 3),
        ("RIGHTPADDING",  (0,0), (-1,-1), 3),
        ("TOPPADDING",    (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    return t


def _bloque_pagos(f: Factura, S: dict) -> Optional[Table]:
    if not f.pagos:
        return None
    filas = [[Paragraph(t, S["h"]) for t in ["Forma de pago", "Plazo", "Valor"]]]
    for p in f.pagos:
        forma = _FORMA_PAGO.get(p.get("forma_pago", ""), p.get("forma_pago", ""))
        plazo_txt = f"{p.get('plazo','')} {p.get('unidad_tiempo','')}".strip()
        filas.append([
            Paragraph(forma, S["small"]),
            Paragraph(plazo_txt, S["small"]),
            Paragraph(_moneda(p.get("total", 0.0)), S["small"]),
        ])
    t = Table(filas, colWidths=[120*mm, 35*mm, 30*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), colors.HexColor("#2C3E50")),
        ("ALIGN",         (2,1), (2,-1), "RIGHT"),
        ("BOX",           (0,0), (-1,-1), 0.7, colors.black),
        ("INNERGRID",     (0,0), (-1,-1), 0.3, colors.grey),
        ("LEFTPADDING",   (0,0), (-1,-1), 3),
        ("RIGHTPADDING",  (0,0), (-1,-1), 3),
        ("TOPPADDING",    (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    return t


# ── API pública ──────────────────────────────────────────────────────────────

def generar_pdf_factura(factura: Factura, pdf_path: Path) -> Path:
    """Genera un PDF (RIDE) para una factura. Retorna la ruta al archivo."""
    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    S = _estilos()
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=12*mm, rightMargin=12*mm,
        topMargin=10*mm,  bottomMargin=12*mm,
        title=f"Factura {factura.numero}",
        author=factura.razon_social,
    )

    story = []
    story.append(_bloque_encabezado(factura, S))
    story.append(Spacer(1, 4*mm))
    story.append(_bloque_comprador(factura, S))
    story.append(Spacer(1, 3*mm))
    story.append(_bloque_detalles(factura, S))
    story.append(Spacer(1, 3*mm))
    story.append(_bloque_totales(factura, S))
    story.append(Spacer(1, 3*mm))
    pagos = _bloque_pagos(factura, S)
    if pagos is not None:
        story.append(pagos)
        story.append(Spacer(1, 3*mm))
    info = _bloque_info_adicional(factura, S)
    if info is not None:
        story.append(info)

    doc.build(story)
    return pdf_path


import unicodedata as _ud
from datetime import datetime as _dt

# Mapa codDoc SRI → abreviatura de 3 letras para el nombre del archivo
_TIPO_ABREV = {
    "01": "FAC",   # Factura
    "03": "LIQ",   # Liquidación de compras
    "04": "NCR",   # Nota de crédito
    "05": "NDB",   # Nota de débito
    "06": "GRE",   # Guía de remisión
    "07": "RET",   # Comprobante de retención
}


def _abrev_tipo(cod_doc: str) -> str:
    """codDoc SRI ('01','04',...) → abreviatura ('FAC','NCR',...). Default FAC."""
    return _TIPO_ABREV.get((cod_doc or "").strip().zfill(2), "FAC")


def _slug_razon(razon: str, max_len: int = 10) -> str:
    """Toma las primeras N LETRAS de la razón social, sin tildes/espacios/símbolos."""
    s = _ud.normalize("NFKD", razon or "").encode("ascii", "ignore").decode()
    # Solo letras (descarta números, puntos, guiones, comas, &, etc.)
    s = "".join(c for c in s if c.isalpha())
    return s[:max_len].upper()


def _fecha_iso(fecha_str: str) -> str:
    """'01/05/2026' o '2026-05-01' o '2026-05-01T...' → '2026-05-01'."""
    s = (fecha_str or "").strip().split(" ")[0].split("T")[0]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return _dt.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def nombre_base(f: Factura, tipo_abrev: str = None, emitido: bool = False) -> str:
    """
    Nombre base (sin extensión) compartido por el XML y el PDF:
        <YYYY-MM-DD>_<RAZON10>_<TIPO>_<numero>
    Ej.: 2026-05-01_TIENDASTUT_FAC_379-002-000157631

    En RECIBIDOS la razón es el emisor (la contraparte que te facturó).
    En EMITIDOS la razón es la contraparte: cliente (FC/NC/ND) o proveedor (LIQ),
    porque el emisor sería siempre tu propia empresa.
    """
    fecha  = _fecha_iso(f.fecha_emision) or "sinfecha"
    rs     = (f.razon_social_comprador if emitido else f.razon_social) or f.razon_social
    razon  = _slug_razon(rs) or "sincontraparte"
    numero = f.numero or (f.clave_acceso[:15] if f.clave_acceso else "snum")
    tipo   = tipo_abrev or _abrev_tipo(getattr(f, "cod_doc", "") or "01")
    return f"{fecha}_{razon}_{tipo}_{numero}"


def _nombre_pdf(f: Factura, tipo_abrev: str = None, emitido: bool = False) -> str:
    return nombre_base(f, tipo_abrev, emitido) + ".pdf"


# ── Retención (RIDE propio) ──────────────────────────────────────────────────

def nombre_base_retencion(r: Retencion, emitido: bool = False) -> str:
    """
    Nombre base (sin extensión) para retenciones, compartido por XML y PDF.
    En EMITIDAS usa el sujeto retenido (la contraparte); en RECIBIDAS, el agente
    de retención (emisor).
    """
    fecha  = _fecha_iso(r.fecha_emision) or "sinfecha"
    rs     = (r.razon_social_sujeto if emitido else r.razon_social) or r.razon_social
    razon  = _slug_razon(rs) or "sinsujeto"
    numero = r.numero or (r.clave_acceso[:15] if r.clave_acceso else "snum")
    return f"{fecha}_{razon}_RET_{numero}"


def _nombre_pdf_retencion(r: Retencion, emitido: bool = False) -> str:
    return nombre_base_retencion(r, emitido) + ".pdf"


def generar_pdf_retencion(ret: Retencion, pdf_path: Path) -> Path:
    """Genera el RIDE en PDF para un comprobante de retención."""
    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    S = _estilos()
    doc = SimpleDocTemplate(
        str(pdf_path), pagesize=A4,
        leftMargin=12*mm, rightMargin=12*mm, topMargin=10*mm, bottomMargin=12*mm,
        title=f"Retencion {ret.numero}", author=ret.razon_social,
    )
    story = []

    # Encabezado: agente de retención + datos del comprobante
    izq = Table([
        [Paragraph(ret.razon_social or "—", S["titulo"])],
        [Paragraph(f"<b>Nombre Comercial:</b> {ret.nombre_comercial or '—'}", S["val"])],
    ], colWidths=[95*mm])
    izq.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),1),
                             ("BOTTOMPADDING",(0,0),(-1,-1),1)]))
    der = Table([
        [Paragraph(f"<b>R.U.C.:</b> {ret.ruc}", S["val"])],
        [Paragraph("<b>COMPROBANTE DE RETENCIÓN</b>", S["titulo"])],
        [Paragraph(f"<b>Nº:</b> {ret.numero}", S["val"])],
        [Paragraph("<b>NÚMERO DE AUTORIZACIÓN / CLAVE DE ACCESO</b>", S["smallB"])],
        [Paragraph(ret.clave_acceso or "—", S["clave"])],
        [Paragraph(f"<b>AMBIENTE:</b> {'PRODUCCIÓN' if ret.ambiente == '2' else 'PRUEBAS'}", S["small"])],
        [Paragraph(f"<b>PERÍODO FISCAL:</b> {ret.periodo_fiscal or '—'}", S["small"])],
    ], colWidths=[90*mm])
    der.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.7,colors.black),
                             ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
                             ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2)]))
    enc = Table([[izq, der]], colWidths=[95*mm, 90*mm])
    enc.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(enc)
    story.append(Spacer(1, 4*mm))

    # Sujeto retenido
    suj = Table([[
        Paragraph(f"<b>Sujeto Retenido:</b> {ret.razon_social_sujeto or '—'}", S["val"]),
        Paragraph(f"<b>Identificación:</b> {ret.identificacion_sujeto or '—'}", S["val"]),
    ],[
        Paragraph(f"<b>Fecha Emisión:</b> {ret.fecha_emision or '—'}", S["val"]),
        Paragraph(f"<b>Período Fiscal:</b> {ret.periodo_fiscal or '—'}", S["val"]),
    ]], colWidths=[110*mm, 75*mm])
    suj.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.7,colors.black),
                             ("INNERGRID",(0,0),(-1,-1),0.3,colors.grey),
                             ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
                             ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3)]))
    story.append(suj)
    story.append(Spacer(1, 3*mm))

    # Tabla de retenciones
    TIPO_IMP = {"1": "RENTA", "2": "IVA", "6": "ISD"}
    enc_cols = ["Comprobante", "Nº Doc. Sustento", "Impuesto", "Cód.", "Base Imp.", "%", "Valor Ret."]
    filas = [[Paragraph(t, S["h"]) for t in enc_cols]]
    for imp in ret.impuestos:
        filas.append([
            Paragraph(imp.cod_doc_sustento or "", S["small"]),
            Paragraph(imp.num_doc_sustento or "", S["small"]),
            Paragraph(TIPO_IMP.get(imp.codigo, imp.codigo), S["small"]),
            Paragraph(imp.codigo_retencion or "", S["small"]),
            Paragraph(_moneda(imp.base_imponible), S["small"]),
            Paragraph(_moneda(imp.porcentaje_retener), S["small"]),
            Paragraph(_moneda(imp.valor_retenido), S["small"]),
        ])
    t = Table(filas, colWidths=[22*mm, 38*mm, 20*mm, 14*mm, 28*mm, 14*mm, 28*mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#833C00")),
        ("ALIGN",(4,1),(-1,-1),"RIGHT"),("VALIGN",(0,0),(-1,-1),"TOP"),
        ("BOX",(0,0),(-1,-1),0.7,colors.black),("INNERGRID",(0,0),(-1,-1),0.3,colors.grey),
        ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
    ]))
    story.append(t)
    story.append(Spacer(1, 3*mm))

    # Total retenido
    tot = Table([[Paragraph("<b>TOTAL RETENIDO</b>", S["smallB"]),
                  Paragraph(f"<b>{_moneda(ret.total_retenido)}</b>", S["smallB"])]],
                colWidths=[35*mm, 25*mm], hAlign="RIGHT")
    tot.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.7,colors.black),
                             ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#ECF0F1")),
                             ("ALIGN",(1,0),(1,-1),"RIGHT"),
                             ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
                             ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2)]))
    story.append(tot)

    doc.build(story)
    return pdf_path


def generar_pdfs_lote(facturas: List[Factura], pdf_dir: Path,
                      retenciones: Optional[List[Retencion]] = None,
                      emitido: bool = False) -> dict:
    """
    Genera un PDF por cada comprobante (facturas/NC/ND/LIQ y retenciones).
    Si emitido=True, el nombre del archivo usa la contraparte (cliente/sujeto).
    Retorna stats: {ok, errores, archivos, pdf_dir}.
    """
    pdf_dir = Path(pdf_dir)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    ok = 0
    errores = 0
    archivos = []

    for f in facturas or []:
        if not f or getattr(f, "error_parseo", ""):
            errores += 1
            continue
        try:
            ruta = pdf_dir / _nombre_pdf(f, emitido=emitido)
            generar_pdf_factura(f, ruta)
            archivos.append(str(ruta))
            ok += 1
        except Exception as e:
            errores += 1
            logger.warning(f"Error PDF {getattr(f,'numero','?')}: {e}")

    for r in retenciones or []:
        if not r or getattr(r, "error_parseo", ""):
            errores += 1
            continue
        try:
            ruta = pdf_dir / _nombre_pdf_retencion(r, emitido=emitido)
            generar_pdf_retencion(r, ruta)
            archivos.append(str(ruta))
            ok += 1
        except Exception as e:
            errores += 1
            logger.warning(f"Error PDF RET {getattr(r,'numero','?')}: {e}")

    logger.info(f"PDFs generados: {ok} OK, {errores} errores en {pdf_dir}")
    return {"ok": ok, "errores": errores, "archivos": archivos, "pdf_dir": str(pdf_dir)}
