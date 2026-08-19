"""
txt_parser.py — Parsea el archivo TXT del portal SRI ("Comprobantes recibidos").

Formato real exportado por el portal SRI (separado por TABs, con encabezado):

    RUC_EMISOR  RAZON_SOCIAL_EMISOR  TIPO_COMPROBANTE  SERIE_COMPROBANTE
    CLAVE_ACCESO  FECHA_AUTORIZACION  FECHA_EMISION  IDENTIFICACION_RECEPTOR
    VALOR_SIN_IMPUESTOS  IVA  IMPORTE_TOTAL  NUMERO_DOCUMENTO_MODIFICADO

Cada línea de datos representa un comprobante con su clave de acceso de 49 dígitos.
El parser detecta el encabezado y mapea por NOMBRE de columna (robusto ante
reordenamientos). Si no hay encabezado reconocible, usa un parseo posicional de
respaldo.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from utils import get_logger, normalizar_fecha

logger = get_logger("txt_parser")


@dataclass
class RegistroTXT:
    tipo_identificacion: str = ""
    ruc_emisor: str = ""
    razon_emisor: str = ""
    tipo_comprobante: str = ""          # codDoc SRI: 01 factura, 07 retención, ...
    establecimiento: str = ""
    punto_emision: str = ""
    secuencial: str = ""
    fecha_emision: str = ""
    fecha_autorizacion: str = ""
    clave_acceso: str = ""
    importe_total: str = ""
    estado: str = ""
    identificacion_receptor: str = ""
    raw: dict = field(default_factory=dict, repr=False)

    @property
    def numero_comprobante(self) -> str:
        return f"{self.establecimiento}-{self.punto_emision}-{self.secuencial}"

    @property
    def tipo_nombre(self) -> str:
        from config import TIPOS_COMPROBANTE
        return TIPOS_COMPROBANTE.get(self.tipo_comprobante, f"TIPO_{self.tipo_comprobante}")

    @property
    def es_factura(self) -> bool:
        return self.tipo_comprobante == "01"

    @property
    def es_retencion(self) -> bool:
        return self.tipo_comprobante == "07"

    @property
    def es_valida(self) -> bool:
        return len(self.clave_acceso) == 49 and self.clave_acceso.isdigit()


# ── Helpers ──────────────────────────────────────────────────────────────────

def _norm(texto: str) -> str:
    """Normaliza: sin acentos, minúsculas, espacios colapsados."""
    t = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode()
    return " ".join(t.lower().split())


# Texto del TIPO_COMPROBANTE (portal) → codDoc SRI
_TIPO_TEXTO_A_CODIGO = {
    "factura": "01",
    "liquidacion de compra de bienes y prestacion de servicios": "03",
    "liquidacion de compra": "03",
    "nota de credito": "04",
    "notas de credito": "04",
    "nota de debito": "05",
    "notas de debito": "05",
    "guia de remision": "06",
    "comprobante de retencion": "07",
    "retencion": "07",
}


def _tipo_texto_a_codigo(texto: str) -> str:
    """Convierte el texto del tipo de comprobante a su codDoc SRI (01,03,...)."""
    t = _norm(texto)
    if not t:
        return ""
    # Coincidencia exacta
    if t in _TIPO_TEXTO_A_CODIGO:
        return _TIPO_TEXTO_A_CODIGO[t]
    # Coincidencia por subcadena (tolerante a variaciones)
    if "retencion" in t:
        return "07"
    if "credito" in t:
        return "04"
    if "debito" in t:
        return "05"
    if "liquidacion" in t:
        return "03"
    if "guia" in t:
        return "06"
    if "factura" in t:
        return "01"
    # Si ya viene como código de 2 dígitos
    if t.isdigit():
        return t.zfill(2)
    return ""


def _split_serie(serie: str):
    """'379-002-000157631' → ('379', '002', '000157631')."""
    partes = (serie or "").split("-")
    if len(partes) == 3:
        return partes[0].strip().zfill(3), partes[1].strip().zfill(3), partes[2].strip().zfill(9)
    return "", "", (serie or "").strip()


# ── Parser principal ─────────────────────────────────────────────────────────

def parsear_txt(ruta_txt) -> list:
    ruta = Path(ruta_txt)
    if not ruta.exists():
        raise FileNotFoundError(f"Archivo TXT no encontrado: {ruta}")

    for encoding in ("utf-8-sig", "latin-1", "cp1252", "utf-8"):
        try:
            contenido = ruta.read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(f"No se pudo leer {ruta}")

    lineas = [l for l in contenido.splitlines() if l.strip()]
    if not lineas:
        logger.warning(f"TXT vacio: {ruta}")
        return []

    sep = "\t" if lineas[0].count("\t") >= lineas[0].count("|") else "|"
    cab = lineas[0].upper()
    tiene_encabezado = ("CLAVE_ACCESO" in cab) or \
                       ("RUC_EMISOR" in cab and "TIPO_COMPROBANTE" in cab)

    logger.info(f"TXT: {len(lineas)} lineas, sep='{sep}', "
                f"encabezado={tiene_encabezado}, archivo={ruta.name}")

    if tiene_encabezado:
        registros = _parsear_con_encabezado(lineas, sep)
    else:
        registros = []
        for i, linea in enumerate(lineas, 1):
            try:
                r = _parsear_linea(linea, sep, i)
                if r is not None:
                    registros.append(r)
            except Exception as e:
                logger.debug(f"Linea {i} ignorada: {e}")

    logger.info(
        f"TXT parseado: {len(registros)} registros, "
        f"{sum(1 for r in registros if r.es_valida)} con clave válida "
        f"({sum(1 for r in registros if r.es_factura)} FC, "
        f"{sum(1 for r in registros if r.es_retencion)} CR)"
    )
    return registros


def _parsear_con_encabezado(lineas: list, sep: str) -> list:
    """Parsea usando los nombres del encabezado (formato 'Recibidos' del SRI)."""
    header = [h.strip().upper() for h in lineas[0].split(sep)]
    idx = {name: i for i, name in enumerate(header)}

    def col(partes, name) -> str:
        i = idx.get(name)
        return partes[i].strip() if (i is not None and i < len(partes)) else ""

    registros = []
    for linea in lineas[1:]:
        partes = linea.split(sep)

        clave = col(partes, "CLAVE_ACCESO")
        if not (len(clave) == 49 and clave.isdigit()):
            # Respaldo: buscar la clave de 49 dígitos en cualquier parte de la línea
            m = re.search(r"\d{49}", linea)
            clave = m.group(0) if m else ""

        estab, punto, sec = _split_serie(col(partes, "SERIE_COMPROBANTE"))

        registros.append(RegistroTXT(
            ruc_emisor              = col(partes, "RUC_EMISOR"),
            razon_emisor            = col(partes, "RAZON_SOCIAL_EMISOR"),
            tipo_comprobante        = _tipo_texto_a_codigo(col(partes, "TIPO_COMPROBANTE")),
            establecimiento         = estab,
            punto_emision           = punto,
            secuencial              = sec,
            fecha_emision           = normalizar_fecha(col(partes, "FECHA_EMISION")),
            fecha_autorizacion      = normalizar_fecha(col(partes, "FECHA_AUTORIZACION")),
            clave_acceso            = clave,
            importe_total           = col(partes, "IMPORTE_TOTAL"),
            estado                  = "AUTORIZADO",
            identificacion_receptor = col(partes, "IDENTIFICACION_RECEPTOR"),
            raw                     = {header[j] if j < len(header) else str(j): v
                                       for j, v in enumerate(partes)},
        ))
    return registros


def _parsear_linea(linea: str, sep: str, num_linea: int):
    """Parseo posicional de respaldo (formato sin encabezado / heredado)."""
    partes = [p.strip() for p in linea.split(sep)]
    if not partes or partes[0] in ("", "TIPO", "#", "TIPO_IDENTIFICACION", "RUC_EMISOR"):
        return None
    if len(partes) < 9:
        return None
    while len(partes) < 12:
        partes.append("")

    clave = ""
    importe = ""
    estado = ""
    for i, p in enumerate(partes):
        if len(p) == 49 and p.isdigit():
            clave = p
            importe = partes[i + 1] if i + 1 < len(partes) else ""
            estado  = partes[i + 2] if i + 2 < len(partes) else ""
            break

    return RegistroTXT(
        tipo_identificacion = partes[0],
        ruc_emisor          = partes[1],
        razon_emisor        = partes[2],
        tipo_comprobante    = partes[3].zfill(2),
        establecimiento     = partes[4].zfill(3),
        punto_emision       = partes[5].zfill(3),
        secuencial          = partes[6].zfill(9),
        fecha_emision       = normalizar_fecha(partes[7]),
        fecha_autorizacion  = normalizar_fecha(partes[8]),
        clave_acceso        = clave,
        importe_total       = importe,
        estado              = estado,
        raw                 = {str(j): v for j, v in enumerate(partes)},
    )


def filtrar_por_tipo(registros: list, tipos: list) -> list:
    return [r for r in registros if r.tipo_comprobante in tipos]


def claves_unicas(registros: list) -> list:
    vistas = set()
    resultado = []
    for r in registros:
        if r.es_valida and r.clave_acceso not in vistas:
            vistas.add(r.clave_acceso)
            resultado.append(r.clave_acceso)
    return resultado
