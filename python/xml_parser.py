"""
xml_parser.py — Parsea XML de comprobantes electronicos SRI.
Soporta: Facturas (01) y Retenciones (07).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from lxml import etree
from utils import get_logger

logger = get_logger("xml_parser")


@dataclass
class ImpuestoFactura:
    codigo: str = ""
    codigo_porcentaje: str = ""
    tarifa: str = ""
    base_imponible: float = 0.0
    valor: float = 0.0


@dataclass
class DetalleFactura:
    codigo_principal: str = ""
    codigo_auxiliar: str = ""
    descripcion: str = ""
    cantidad: float = 0.0
    precio_unitario: float = 0.0
    descuento: float = 0.0
    total_sin_impuesto: float = 0.0
    tarifa_iva: str = ""
    base_iva: float = 0.0
    valor_iva: float = 0.0


@dataclass
class Factura:
    ambiente: str = ""
    cod_doc: str = "01"
    razon_social: str = ""
    nombre_comercial: str = ""
    ruc: str = ""
    dir_matriz: str = ""
    clave_acceso: str = ""
    establecimiento: str = ""
    punto_emision: str = ""
    secuencial: str = ""
    fecha_emision: str = ""
    fecha_autorizacion: str = ""
    numero_autorizacion: str = ""
    dir_establecimiento: str = ""
    contribuyente_especial: str = ""
    obligado_contabilidad: str = ""
    tipo_identificacion_comprador: str = ""
    razon_social_comprador: str = ""
    identificacion_comprador: str = ""
    direccion_comprador: str = ""
    total_sin_impuestos: float = 0.0
    total_descuento: float = 0.0
    importe_total: float = 0.0
    propina: float = 0.0
    moneda: str = "DOLAR"
    base_iva_0: float = 0.0
    base_iva_12: float = 0.0
    base_iva_14: float = 0.0
    base_iva_15: float = 0.0
    base_no_objeto: float = 0.0
    base_exenta: float = 0.0
    valor_iva: float = 0.0
    impuestos_totales: list = field(default_factory=list)
    detalles: list = field(default_factory=list)
    info_adicional: dict = field(default_factory=dict)
    pagos: list = field(default_factory=list)
    # Campos específicos de Nota de Crédito / Nota de Débito
    cod_doc_modificado: str = ""
    num_doc_modificado: str = ""
    fecha_doc_sustento: str = ""
    motivo: str = ""
    # 'comprador' (FC/NC/ND) o 'proveedor' (Liquidación de compra)
    rol_contraparte: str = "comprador"
    ruta_xml: str = ""
    error_parseo: str = ""

    @property
    def numero(self) -> str:
        return f"{self.establecimiento}-{self.punto_emision}-{self.secuencial}"

    @property
    def tipo_nombre(self) -> str:
        return {
            "01": "FACTURA", "03": "LIQUIDACION DE COMPRA",
            "04": "NOTA DE CREDITO", "05": "NOTA DE DEBITO",
        }.get(self.cod_doc, "COMPROBANTE")


@dataclass
class ImpuestoRetencion:
    codigo: str = ""
    codigo_retencion: str = ""
    base_imponible: float = 0.0
    porcentaje_retener: float = 0.0
    valor_retenido: float = 0.0
    cod_doc_sustento: str = ""
    num_doc_sustento: str = ""
    fecha_emision_doc_sustento: str = ""


@dataclass
class Retencion:
    ambiente: str = ""
    razon_social: str = ""
    nombre_comercial: str = ""
    ruc: str = ""
    clave_acceso: str = ""
    establecimiento: str = ""
    punto_emision: str = ""
    secuencial: str = ""
    fecha_emision: str = ""
    fecha_autorizacion: str = ""
    numero_autorizacion: str = ""
    tipo_identificacion_sujeto: str = ""
    razon_social_sujeto: str = ""
    identificacion_sujeto: str = ""
    periodo_fiscal: str = ""
    impuestos: list = field(default_factory=list)
    total_retenido_renta: float = 0.0
    total_retenido_iva: float = 0.0
    total_retenido: float = 0.0
    ruta_xml: str = ""
    error_parseo: str = ""

    @property
    def numero(self) -> str:
        return f"{self.establecimiento}-{self.punto_emision}-{self.secuencial}"

    @property
    def cod_doc_sustento(self) -> str:
        """Tipo de doc sustento (del primer impuesto). Para reporte de retenciones."""
        return self.impuestos[0].cod_doc_sustento if self.impuestos else ""

    @property
    def num_doc_sustento(self) -> str:
        return self.impuestos[0].num_doc_sustento if self.impuestos else ""

    @property
    def renta(self):
        """Primer impuesto de Renta (codigo='1') o None."""
        return next((i for i in self.impuestos if i.codigo == "1"), None)

    @property
    def iva(self):
        """Primer impuesto de IVA (codigo='2') o None."""
        return next((i for i in self.impuestos if i.codigo == "2"), None)


def _t(el, xpath: str, default: str = "") -> str:
    if el is None:
        return default
    nodos = el.xpath(xpath)
    if not nodos:
        return default
    val = nodos[0]
    return (val.text or "").strip() if hasattr(val, "text") else str(val).strip()


def _f(el, xpath: str, default: float = 0.0) -> float:
    txt = _t(el, xpath, "")
    if not txt:
        return default
    try:
        return float(txt.replace(",", "."))
    except ValueError:
        return default


def parsear_xml(ruta_o_str, es_string: bool = False):
    try:
        if es_string:
            contenido = ruta_o_str.encode("utf-8") if isinstance(ruta_o_str, str) else ruta_o_str
            root = etree.fromstring(contenido)
            ruta_xml = ""
        else:
            ruta = Path(ruta_o_str)
            if not ruta.exists():
                return None
            with open(ruta, "rb") as f:
                root = etree.parse(f).getroot()
            ruta_xml = str(ruta)
    except etree.XMLSyntaxError as e:
        logger.error(f"XML invalido: {e}")
        return None

    def _tag(el):
        return el.tag.lower().split("}")[-1] if "}" in el.tag else el.tag.lower()

    tag = _tag(root)

    # Datos del wrapper de autorización (si existe): fecha y número de autorización.
    # No están en el comprobante interno — solo el SRI los asigna al autorizar.
    aut_fecha = ""
    aut_numero = ""

    # Si viene envuelto en <autorizacion> (SOAP/portal), extraer el comprobante
    # interno del CDATA y conservar los metadatos de autorización.
    if tag == "autorizacion" or tag == "respuestaautorizacioncomprobante":
        aut_fecha  = _t(root, ".//fechaAutorizacion")
        aut_numero = _t(root, ".//numeroAutorizacion")
        comp = root.find(".//comprobante")
        if comp is not None and comp.text:
            try:
                inner = etree.fromstring(comp.text.encode("utf-8"))
                root = inner
                tag = _tag(root)
            except Exception as e:
                logger.error(f"No se pudo extraer comprobante de autorizacion: {e}")
                return None

    # Comprobantes "tipo venta" → un solo parser generalizado
    INFO_BLOCKS = {
        "factura":          ("infoFactura",          "comprador"),
        "notacredito":      ("infoNotaCredito",      "comprador"),
        "notadebito":       ("infoNotaDebito",       "comprador"),
        "liquidacioncompra":("infoLiquidacionCompra","proveedor"),
    }
    if tag in INFO_BLOCKS:
        bloque, rol = INFO_BLOCKS[tag]
        obj = _parsear_comprobante(root, ruta_xml, bloque, rol)
    elif tag == "comprobanteretencion":
        obj = _parsear_retencion(root, ruta_xml)
    else:
        logger.warning(f"Tipo no soportado: {root.tag}")
        return None

    # Propagar metadatos de autorización (solo si los trae el wrapper)
    if obj is not None:
        if aut_fecha:
            obj.fecha_autorizacion = aut_fecha
        if aut_numero and not getattr(obj, "numero_autorizacion", ""):
            obj.numero_autorizacion = aut_numero
    return obj


import re as _re
_RE_CLAVE49 = _re.compile(r'(\d{49})')


def clave_de_xml(ruta) -> Optional[str]:
    """Extrae la claveAcceso (49 dígitos) del contenido de un XML."""
    try:
        txt = Path(ruta).read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return None
    m = _re.search(r'claveAcceso>\s*(\d{49})', txt) or _RE_CLAVE49.search(txt)
    return m.group(1) if m else None


def claves_en_directorio(xml_dir) -> set:
    """
    Conjunto de claves de acceso ya presentes en el directorio, leídas del
    CONTENIDO de cada XML (robusto aunque los archivos estén renombrados).
    """
    xml_dir = Path(xml_dir)
    if not xml_dir.exists():
        return set()
    claves = set()
    for f in xml_dir.glob("*.xml"):
        c = clave_de_xml(f)
        if c:
            claves.add(c)
    return claves


def parsear_directorio(xml_dir: Path):
    xml_dir = Path(xml_dir)
    archivos = sorted(xml_dir.glob("*.xml"))
    facturas = []
    retenciones = []
    errores = 0
    logger.info(f"Parseando {len(archivos)} XMLs en {xml_dir}")
    for ruta in archivos:
        resultado = parsear_xml(ruta)
        if resultado is None:
            errores += 1
        elif isinstance(resultado, Factura):
            facturas.append(resultado)
        elif isinstance(resultado, Retencion):
            retenciones.append(resultado)
    logger.info(f"Resultado: {len(facturas)} FC, {len(retenciones)} CR, {errores} errores")
    return facturas, retenciones


def _parsear_comprobante(root, ruta_xml: str, bloque: str, rol: str) -> Factura:
    """
    Parser generalizado para comprobantes 'tipo venta':
    Factura, Nota de Crédito, Nota de Débito y Liquidación de compra.

    Args:
        bloque: nombre del bloque de info ('infoFactura', 'infoNotaCredito', ...)
        rol:    'comprador' (FC/NC/ND) o 'proveedor' (Liquidación)
    """
    f = Factura(ruta_xml=ruta_xml, rol_contraparte=rol)
    try:
        it = root.find("infoTributaria")
        f.ambiente        = _t(it, "ambiente")
        f.cod_doc         = _t(it, "codDoc") or "01"
        f.razon_social    = _t(it, "razonSocial")
        f.nombre_comercial= _t(it, "nombreComercial") or f.razon_social
        f.ruc             = _t(it, "ruc")
        f.dir_matriz      = _t(it, "dirMatriz")
        f.clave_acceso    = _t(it, "claveAcceso")
        f.establecimiento = _t(it, "estab")
        f.punto_emision   = _t(it, "ptoEmi")
        f.secuencial      = _t(it, "secuencial")

        inf = root.find(bloque)
        f.fecha_emision          = _t(inf, "fechaEmision")
        f.dir_establecimiento    = _t(inf, "dirEstablecimiento") or f.dir_matriz
        f.contribuyente_especial = _t(inf, "contribuyenteEspecial")
        f.obligado_contabilidad  = _t(inf, "obligadoContabilidad")
        f.moneda                 = _t(inf, "moneda") or "DOLAR"

        # Contraparte: comprador (FC/NC/ND) o proveedor (Liquidación).
        # Se guarda siempre en los campos *_comprador para reutilizar reporte/PDF.
        if rol == "proveedor":
            f.tipo_identificacion_comprador = _t(inf, "tipoIdentificacionProveedor")
            f.razon_social_comprador        = _t(inf, "razonSocialProveedor")
            f.identificacion_comprador      = _t(inf, "identificacionProveedor")
            f.direccion_comprador           = _t(inf, "direccionProveedor")
        else:
            f.tipo_identificacion_comprador = _t(inf, "tipoIdentificacionComprador")
            f.razon_social_comprador        = _t(inf, "razonSocialComprador")
            f.identificacion_comprador      = _t(inf, "identificacionComprador")
            f.direccion_comprador           = _t(inf, "direccionComprador")

        f.total_sin_impuestos = _f(inf, "totalSinImpuestos")
        f.total_descuento     = _f(inf, "totalDescuento")
        # Total: factura/NC/LIQ usan importeTotal; NC usa valorModificacion;
        # ND usa valorTotal. Tomamos el primero que exista.
        f.importe_total = (_f(inf, "importeTotal")
                           or _f(inf, "valorTotal")
                           or _f(inf, "valorModificacion"))
        f.propina       = _f(inf, "propina")

        # Campos específicos de NC / ND
        f.cod_doc_modificado = _t(inf, "codDocModificado")
        f.num_doc_modificado = _t(inf, "numDocModificado")
        f.fecha_doc_sustento = _t(inf, "fechaEmisionDocSustento")
        f.motivo             = _t(inf, "motivo")
        # Nota de débito: motivos/motivo/razon
        if not f.motivo:
            razones = [r for r in root.xpath(".//motivos/motivo/razon/text()")]
            if razones:
                f.motivo = "; ".join(x.strip() for x in razones if x.strip())

        # Impuestos consolidados. FC/NC/LIQ → <bloque>/totalConImpuestos/totalImpuesto
        # ND → <infoNotaDebito>/impuestos/impuesto
        nodos_imp = root.xpath(f".//{bloque}/totalConImpuestos/totalImpuesto")
        if not nodos_imp:
            nodos_imp = root.xpath(f".//{bloque}/impuestos/impuesto")
        for imp in nodos_imp:
            iva = ImpuestoFactura(
                codigo=_t(imp, "codigo"), codigo_porcentaje=_t(imp, "codigoPorcentaje"),
                tarifa=_t(imp, "tarifa"), base_imponible=_f(imp, "baseImponible"),
                valor=_f(imp, "valor"),
            )
            f.impuestos_totales.append(iva)
            cp = iva.codigo_porcentaje
            if cp == "0":   f.base_iva_0   += iva.base_imponible
            elif cp == "2": f.base_iva_12  += iva.base_imponible; f.valor_iva += iva.valor
            elif cp == "3": f.base_iva_14  += iva.base_imponible; f.valor_iva += iva.valor
            elif cp == "4": f.base_iva_15  += iva.base_imponible; f.valor_iva += iva.valor
            elif cp == "6": f.base_no_objeto += iva.base_imponible
            elif cp == "7": f.base_exenta  += iva.base_imponible

        # Detalles (líneas). codigoPrincipal (FC) o codigoInterno (NC/ND/LIQ).
        for det in root.xpath(".//detalles/detalle"):
            d = DetalleFactura(
                codigo_principal   = _t(det, "codigoPrincipal") or _t(det, "codigoInterno"),
                codigo_auxiliar    = _t(det, "codigoAuxiliar") or _t(det, "codigoAdicional"),
                descripcion        = _t(det, "descripcion"),
                cantidad           = _f(det, "cantidad"),
                precio_unitario    = _f(det, "precioUnitario"),
                descuento          = _f(det, "descuento"),
                total_sin_impuesto = _f(det, "precioTotalSinImpuesto"),
            )
            imp_lin = det.xpath("./impuestos/impuesto")
            if imp_lin:
                d.tarifa_iva  = _t(imp_lin[0], "tarifa")
                d.base_iva    = _f(imp_lin[0], "baseImponible")
                d.valor_iva   = _f(imp_lin[0], "valor")
            f.detalles.append(d)

        # Info adicional
        for ca in root.xpath(".//infoAdicional/campoAdicional"):
            nombre = ca.get("nombre", "").strip()
            valor  = (ca.text or "").strip()
            if nombre:
                f.info_adicional[nombre] = valor

        # Pagos
        for pg in root.xpath(".//pagos/pago"):
            f.pagos.append({
                "forma_pago":     _t(pg, "formaPago"),
                "total":          _f(pg, "total"),
                "plazo":          _t(pg, "plazo"),
                "unidad_tiempo":  _t(pg, "unidadTiempo"),
            })
    except Exception as e:
        f.error_parseo = str(e)
        logger.error(f"Error comprobante {ruta_xml}: {e}")
    return f


def _parsear_retencion(root, ruta_xml: str) -> Retencion:
    r = Retencion(ruta_xml=ruta_xml)
    try:
        it = root.find("infoTributaria")
        r.razon_social    = _t(it, "razonSocial")
        r.nombre_comercial= _t(it, "nombreComercial")
        r.ruc             = _t(it, "ruc")
        r.clave_acceso    = _t(it, "claveAcceso")
        r.establecimiento = _t(it, "estab")
        r.punto_emision   = _t(it, "ptoEmi")
        r.secuencial      = _t(it, "secuencial")

        ic = root.find("infoCompRetencion")
        r.fecha_emision              = _t(ic, "fechaEmision")
        r.tipo_identificacion_sujeto = _t(ic, "tipoIdentificacionSujetoRetenido")
        r.razon_social_sujeto        = _t(ic, "razonSocialSujetoRetenido")
        r.identificacion_sujeto      = _t(ic, "identificacionSujetoRetenido")
        r.periodo_fiscal             = _t(ic, "periodoFiscal")

        # Formato 1.0.0: <impuestos>/<impuesto> con codDocSustento dentro de cada uno.
        nodos = root.xpath(".//impuestos/impuesto")
        if nodos:
            for imp in nodos:
                i = ImpuestoRetencion(
                    codigo=_t(imp, "codigo"), codigo_retencion=_t(imp, "codigoRetencion"),
                    base_imponible=_f(imp, "baseImponible"),
                    porcentaje_retener=_f(imp, "porcentajeRetener"),
                    valor_retenido=_f(imp, "valorRetenido"),
                    cod_doc_sustento=_t(imp, "codDocSustento"),
                    num_doc_sustento=_t(imp, "numDocSustento"),
                    fecha_emision_doc_sustento=_t(imp, "fechaEmisionDocSustento"),
                )
                _acumular_retencion(r, i)
        else:
            # Formato 2.0.0: <docsSustento>/<docSustento> con <retenciones>/<retencion>.
            # El doc sustento (codDocSustento, numDocSustento) está a nivel del docSustento.
            for ds in root.xpath(".//docsSustento/docSustento"):
                cod_sus = _t(ds, "codDocSustento")
                num_sus = _t(ds, "numDocSustento")
                fecha_sus = _t(ds, "fechaEmisionDocSustento")
                for ret in ds.xpath(".//retenciones/retencion"):
                    i = ImpuestoRetencion(
                        codigo=_t(ret, "codigo"),
                        codigo_retencion=_t(ret, "codigoRetencion"),
                        base_imponible=_f(ret, "baseImponible"),
                        porcentaje_retener=_f(ret, "porcentajeRetener"),
                        valor_retenido=_f(ret, "valorRetenido"),
                        cod_doc_sustento=cod_sus,
                        num_doc_sustento=num_sus,
                        fecha_emision_doc_sustento=fecha_sus,
                    )
                    _acumular_retencion(r, i)
    except Exception as e:
        r.error_parseo = str(e)
        logger.error(f"Error CR {ruta_xml}: {e}")
    return r


def _acumular_retencion(r: Retencion, i: ImpuestoRetencion):
    """Agrega un impuesto de retención y acumula totales (Renta/IVA)."""
    r.impuestos.append(i)
    if i.codigo == "1":
        r.total_retenido_renta += i.valor_retenido
    elif i.codigo == "2":
        r.total_retenido_iva += i.valor_retenido
    r.total_retenido += i.valor_retenido
