"""
main.py — CLI del SRI XML Manager.
Comandos: descargar | reportar | todo | auto | version
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import config
from utils import get_logger, timestamp
from txt_parser import parsear_txt, filtrar_por_tipo, claves_unicas
from sri_soap import SRISOAPClient
from xml_parser import parsear_directorio, claves_en_directorio, Retencion
from excel_writer import generar_reporte
from pdf_generator import generar_pdfs_lote, nombre_base, nombre_base_retencion

logger = get_logger("main")


def _cargar_config(ruta: str) -> dict:
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)


def _escribir_output(datos: dict, ruta: str):
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)


def _tipos_a_codigos(tipos) -> list:
    mapa = {"fc": "01", "cr": "07", "nc": "04", "nd": "05", "lc": "08"}
    if not tipos:
        return ["01", "07"]
    return [mapa.get(t.lower(), t) for t in (tipos if isinstance(tipos, list) else [tipos])]


def _recolectar_txts(params: dict, args) -> tuple:
    """
    Reúne los TXT a procesar. Devuelve (lista_de_Path, error).

    Prioridad:
      1. params["carpeta_txt"]  → TODOS los *.txt de esa carpeta (mes completo).
      2. params["txts"]         → lista explícita de rutas.
      3. params["txt"] / --txt  → un solo archivo (compatibilidad).

    Permite descargar un mes entero cargando varios TXT (el SRI a veces solo
    deja bajar un TXT por día). Las claves repetidas entre archivos se deduplican
    más adelante (claves_unicas), así que solapamientos son inofensivos.
    """
    carpeta = str(params.get("carpeta_txt", "") or "").strip()
    if carpeta:
        d = Path(carpeta)
        if not d.is_dir():
            return [], f"Carpeta de TXT no encontrada: {d}"
        archivos = sorted(d.glob("*.txt"))
        if not archivos:
            return [], f"La carpeta no contiene archivos .txt: {d}"
        logger.info(f"Carpeta de TXT: {len(archivos)} archivos en {d}")
        return archivos, ""

    lista = params.get("txts")
    if lista:
        archivos = [Path(str(p)) for p in lista if str(p).strip()]
        faltan = [p for p in archivos if not p.exists()]
        if faltan:
            return [], f"TXT no encontrado(s): {', '.join(str(p) for p in faltan)}"
        return archivos, ""

    txt_uno = str(params.get("txt", getattr(args, "txt", None) or "") or "").strip()
    if txt_uno:
        p = Path(txt_uno)
        if not p.exists():
            return [], f"TXT no encontrado: {p}"
        return [p], ""

    return [], "No se especificó ningún TXT (use 'txt' o 'carpeta_txt')"


def cmd_descargar(args) -> dict:
    params = {}
    if hasattr(args, "config") and args.config:
        params = _cargar_config(args.config)

    dest_dir   = Path(params.get("dest", getattr(args, "dest", None) or config.XML_DIR))
    tipos_raw  = params.get("tipos", getattr(args, "tipo", None) or [])
    produccion = params.get("produccion", getattr(args, "produccion", True))

    txts, err = _recolectar_txts(params, args)
    if err:
        return {"ok": False, "error": err}

    # Unir los registros de todos los TXT (un mes puede venir en ~30 archivos)
    registros = []
    for t in txts:
        try:
            registros.extend(parsear_txt(t))
        except Exception as e:                               # noqa: BLE001
            logger.warning(f"No se pudo leer TXT {t}: {e}")
    if len(txts) > 1:
        logger.info(f"TXT combinados: {len(txts)} archivos, {len(registros)} registros")
    if not registros:
        return {"ok": True, "descargados": 0, "errores": 0, "total": 0}

    # Descargar TODOS los tipos de comprobante presentes en el TXT
    # (FC, NC, ND, Liquidaciones, Retenciones). No se aplica filtro.
    claves = claves_unicas(registros)
    # Detectar ya descargadas por el CONTENIDO (los XML se renombran luego)
    ya = claves_en_directorio(dest_dir)

    logger.info(f"Descarga: {len(claves)} claves, {len(ya)} ya descargadas")

    cliente    = SRISOAPClient(produccion=produccion)
    # "emitidos" en config.json (enviado por VBA) es la fuente de verdad.
    # Fallback: buscar "emitido" en toda la ruta (no solo el nombre de la
    # carpeta final), por si el usuario nombro la carpeta manualmente.
    emitido    = bool(params.get("emitidos")) or ("emitido" in str(dest_dir).lower())
    max_workers = int(params.get("paralelo", 6))      # configurable (6-8 recomendado)
    pendientes = [c for c in claves if c not in ya]

    # Trabajo completo: descarga en paralelo + PDFs + faltantes.
    # Se ejecuta dentro de una ventana de progreso (canal).
    estado = {}

    def trabajo(canal):
        resultados = cliente.descargar_lote(
            claves=claves, xml_dir=dest_dir, claves_ya_descargadas=ya,
            max_workers=max_workers, progreso=canal.progreso)
        canal.mensaje("Generando PDFs y renombrando XMLs...")
        pdf_stats = _generar_pdfs(dest_dir, emitido=emitido)
        canal.mensaje("Registrando faltantes...")
        n_falt = _escribir_faltantes(dest_dir, registros, resultados)
        estado["resultados"] = resultados
        estado["pdf_stats"]  = pdf_stats
        estado["n_falt"]     = n_falt
        return True

    titulo = "SRI - Descargando comprobantes" + (" emitidos" if emitido else "")
    if pendientes:
        from progreso import ejecutar_con_ventana
        ejecutar_con_ventana(titulo, trabajo)
    else:
        from progreso import CanalNulo
        trabajo(CanalNulo())

    resultados = estado.get("resultados", {})
    pdf_stats  = estado.get("pdf_stats", {})
    n_falt     = estado.get("n_falt", 0)
    ok      = sum(1 for r in resultados.values() if r.exitoso)
    errores = sum(1 for r in resultados.values() if not r.exitoso)

    return {"ok": True, "total": len(claves), "descargados": ok,
            "errores": errores, "faltantes": n_falt, "xml_dir": str(dest_dir),
            "txts_procesados": len(txts),
            "pdfs_generados": pdf_stats.get("ok", 0),
            "pdfs_errores":   pdf_stats.get("errores", 0),
            "pdf_dir":        pdf_stats.get("pdf_dir", "")}


def _escribir_faltantes(xml_dir: Path, registros: list, resultados: dict) -> int:
    """
    Compara las claves del TXT contra los XML presentes y guarda en
    xml_dir/_faltantes.json los comprobantes que NO se descargaron, con su
    motivo. Si no falta ninguno, elimina el archivo. Retorna la cantidad.
    """
    import json
    xml_dir = Path(xml_dir)
    presentes = claves_en_directorio(xml_dir)
    faltantes = []
    vistas = set()
    for r in registros:
        c = (r.clave_acceso or "").strip()
        if not (len(c) == 49 and c.isdigit()) or c in presentes or c in vistas:
            continue
        vistas.add(c)
        res = (resultados or {}).get(c)
        motivo = (res.error if (res and res.error) else "No descargado")
        try:
            total = float(str(r.importe_total).replace(",", "."))
        except (ValueError, AttributeError):
            total = r.importe_total
        faltantes.append({
            "clave":      c,
            "tipo":       r.tipo_nombre if r.tipo_comprobante else "",
            "ruc_emisor": r.ruc_emisor,
            "razon":      r.razon_emisor,
            "numero":     r.numero_comprobante,
            "fecha":      r.fecha_emision,
            "total":      total,
            "motivo":     motivo,
        })

    ruta = xml_dir / "_faltantes.json"
    if faltantes:
        ruta.write_text(json.dumps(faltantes, ensure_ascii=False, indent=2),
                        encoding="utf-8")
        logger.info(f"Faltantes registrados: {len(faltantes)} en {ruta.name}")
    elif ruta.exists():
        ruta.unlink()  # ya no falta ninguno
    return len(faltantes)


def _leer_faltantes(xml_dir) -> list:
    """Lee xml_dir/_faltantes.json (lista) si existe."""
    import json
    ruta = Path(xml_dir) / "_faltantes.json"
    if not ruta.exists():
        return []
    try:
        return json.loads(ruta.read_text(encoding="utf-8"))
    except Exception:
        return []


def _renombrar_xmls(facturas: list, retenciones: list, emitido: bool = False):
    """
    Renombra cada XML con la MISMA estructura del PDF:
        <YYYY-MM-DD>_<RAZON10>_<TIPO>_<numero>.xml
    Si emitido=True, la razón es la contraparte (cliente/proveedor/sujeto).
    Actualiza obj.ruta_xml al nuevo nombre.
    """
    import os
    renombrados = 0
    for obj in list(facturas) + list(retenciones):
        ruta = getattr(obj, "ruta_xml", "")
        if not ruta:
            continue
        p = Path(ruta)
        if not p.exists():
            continue
        base = (nombre_base_retencion(obj, emitido=emitido) if isinstance(obj, Retencion)
                else nombre_base(obj, emitido=emitido))
        destino = p.with_name(base + ".xml")
        if destino == p:
            continue
        if destino.exists():
            continue  # evitar sobrescribir otro XML
        try:
            os.rename(p, destino)
            obj.ruta_xml = str(destino)
            renombrados += 1
        except OSError as e:
            logger.warning(f"No se pudo renombrar {p.name}: {e}")
    if renombrados:
        logger.info(f"XMLs renombrados con estructura descriptiva: {renombrados}")


def _generar_pdfs(xml_dir: Path, emitido: bool = False) -> dict:
    """
    Parsea los XMLs en xml_dir, los renombra con la estructura descriptiva
    (igual que los PDFs) y genera un PDF (RIDE) por cada comprobante.
    Los PDFs se guardan en una carpeta "PDFs" HERMANA de xml_dir (por ejemplo,
    si xml_dir es ".../Documentos Recibidos/XMLs", los PDFs van a
    ".../Documentos Recibidos/PDFs"), para que XMLs y PDFs queden ordenados
    bajo la misma carpeta del tipo de documento.
    """
    try:
        facturas, retenciones = parsear_directorio(xml_dir)
        if not facturas and not retenciones:
            return {"ok": 0, "errores": 0, "pdf_dir": ""}
        _renombrar_xmls(facturas, retenciones, emitido=emitido)
        pdf_dir = Path(xml_dir).parent / "PDFs"
        return generar_pdfs_lote(facturas, pdf_dir, retenciones=retenciones, emitido=emitido)
    except Exception as e:
        logger.warning(f"Error generando PDFs: {e}")
        return {"ok": 0, "errores": 0, "pdf_dir": "", "error": str(e)}


def cmd_reportar(args) -> dict:
    params = {}
    if hasattr(args, "config") and args.config:
        params = _cargar_config(args.config)

    xml_dir = Path(params.get("xml_dir", getattr(args, "xmldir", None) or config.XML_DIR))
    output  = Path(params.get("output",  getattr(args, "output", None)
                              or config.OUTPUT_DIR / f"SRI_Reporte_{timestamp()}.xlsx"))

    if not xml_dir.exists():
        return {"ok": False, "error": f"Directorio no encontrado: {xml_dir}"}

    facturas, retenciones = parsear_directorio(xml_dir)
    faltantes = _leer_faltantes(xml_dir)
    if not facturas and not retenciones and not faltantes:
        return {"ok": True, "facturas": 0, "retenciones": 0, "output": str(output)}

    ruta = generar_reporte(facturas=facturas, retenciones=retenciones,
                           ruta_salida=output,
                           abrir_al_finalizar=getattr(args, "abrir", False),
                           faltantes=faltantes)
    return {"ok": True, "facturas": len(facturas), "retenciones": len(retenciones),
            "faltantes": len(faltantes), "output": str(ruta)}


def cmd_reportar_emitidos(args) -> dict:
    """Genera el reporte de comprobantes EMITIDOS (una hoja por tipo)."""
    from excel_writer import generar_reporte_emitidos

    params = {}
    if hasattr(args, "config") and args.config:
        params = _cargar_config(args.config)

    xml_dir = Path(params.get("xml_dir", getattr(args, "xmldir", None) or config.XML_DIR))
    output  = Path(params.get("output",  getattr(args, "output", None)
                              or config.OUTPUT_DIR / f"SRI_Emitidos_{timestamp()}.xlsx"))

    if not xml_dir.exists():
        return {"ok": False, "error": f"Directorio no encontrado: {xml_dir}"}

    facturas, retenciones = parsear_directorio(xml_dir)
    faltantes = _leer_faltantes(xml_dir)
    if not facturas and not retenciones and not faltantes:
        return {"ok": True, "facturas": 0, "retenciones": 0, "output": str(output)}

    ruta = generar_reporte_emitidos(facturas=facturas, retenciones=retenciones,
                                    ruta_salida=output,
                                    abrir_al_finalizar=getattr(args, "abrir", False),
                                    faltantes=faltantes)
    fc  = sum(1 for f in facturas if getattr(f, "cod_doc", "01") == "01")
    nc  = sum(1 for f in facturas if getattr(f, "cod_doc", "") in ("04", "05"))
    liq = sum(1 for f in facturas if getattr(f, "cod_doc", "") == "03")
    return {"ok": True, "facturas": fc, "notas": nc, "liquidaciones": liq,
            "retenciones": len(retenciones), "faltantes": len(faltantes),
            "output": str(ruta)}


def cmd_todo(args) -> dict:
    res = cmd_descargar(args)
    if not res.get("ok"):
        return res
    if "xml_dir" in res:
        args.xmldir = res["xml_dir"]
    return {**res, **cmd_reportar(args)}


def cmd_auto(args) -> dict:
    """
    Descarga automática completa:
      1. Login SRI → descargar TXT vía Selenium
      2. Parsear TXT → descargar XMLs vía SOAP
      3. Generar reporte Excel
    """
    try:
        from sri_selenium import SRISelenium
    except ImportError as e:
        return {"ok": False,
                "error": f"selenium no instalado: {e}. "
                         "Ejecute: pip install selenium webdriver-manager"}

    params = {}
    if hasattr(args, "config") and args.config:
        params = _cargar_config(args.config)

    ruc      = params.get("ruc")      or getattr(args, "ruc",      None) or ""
    clave    = params.get("clave")    or getattr(args, "clave",    None) or ""
    desde    = params.get("desde")    or getattr(args, "desde",    None) or ""
    hasta    = params.get("hasta")    or getattr(args, "hasta",    None) or ""
    headless = params.get("headless", not getattr(args, "visible",  False))
    dest_str = params.get("dest")     or getattr(args, "dest",     None) or str(config.XML_DIR)
    output_j = params.get("output_json")
    output_p = params.get("output")   or getattr(args, "output",   None)
    produccion = params.get("produccion", True)

    if not ruc:
        return {"ok": False, "error": "RUC requerido (Config C4 o --ruc)"}
    if not clave:
        return {"ok": False, "error": "Clave SRI requerida (Config C14 o --clave)"}
    if not desde or not hasta:
        return {"ok": False,
                "error": "Fechas requeridas: --desde YYYY-MM-DD --hasta YYYY-MM-DD"}

    dest_dir = Path(dest_str)
    xml_dir  = dest_dir   # El VBA ya construye la ruta completa: xmls\RUC\
    logger.info(f"Descarga automática: RUC={ruc}, {desde}→{hasta}")

    # ── Paso 1: Login → navegar → descargar XMLs directamente ─
    # El portal SRI (2025+) usa Angular+Keycloak: se descargan los XMLs
    # directo desde la tabla del portal, sin necesidad de SOAP.
    selenium  = SRISelenium(ruc=ruc, clave=clave, headless=headless)
    resultado = selenium.descargar_xmls(desde=desde, hasta=hasta, xml_dir=xml_dir)

    if not resultado.exitoso:
        return {"ok": False, "error": f"Error portal SRI: {resultado.mensaje}"}

    ok_count = resultado.total_xmls
    logger.info(f"XMLs descargados: {ok_count} en {xml_dir}")

    # ── Paso 2: Generar reporte + PDFs ────────────────────────
    facturas, retenciones = parsear_directorio(xml_dir)

    ruta_out = None
    if facturas or retenciones:
        out_path = Path(output_p) if output_p else \
                   config.OUTPUT_DIR / f"SRI_Reporte_{timestamp()}.xlsx"
        ruta_out = generar_reporte(facturas=facturas, retenciones=retenciones,
                                   ruta_salida=out_path)

    # PDFs automáticos (un RIDE por factura) + renombrado de XMLs
    pdf_stats = {"ok": 0, "errores": 0, "pdf_dir": ""}
    if facturas or retenciones:
        _renombrar_xmls(facturas, retenciones)
        pdf_dir = Path(xml_dir) / "pdfs"
        try:
            pdf_stats = generar_pdfs_lote(facturas, pdf_dir, retenciones=retenciones)
        except Exception as e:
            logger.warning(f"Error generando PDFs: {e}")

    return {
        "ok":           True,
        "total":        ok_count,
        "descargados":  ok_count,
        "errores":      0,
        "xml_dir":      str(xml_dir),
        "facturas":     len(facturas),
        "retenciones":  len(retenciones),
        "output":       str(ruta_out) if ruta_out else None,
        "pdfs_generados": pdf_stats.get("ok", 0),
        "pdfs_errores":   pdf_stats.get("errores", 0),
        "pdf_dir":        pdf_stats.get("pdf_dir", ""),
    }


def main():
    parser = argparse.ArgumentParser(prog="sri_engine",
        description=f"{config.NOMBRE} v{config.VERSION}")
    sub = parser.add_subparsers(dest="comando", required=True)

    for name, help_txt in [("descargar", "Descarga XMLs via SOAP"),
                            ("reportar",  "Genera reporte Excel (recibidos)"),
                            ("reportar-emitidos", "Genera reporte Excel (emitidos)"),
                            ("todo",      "Descarga + reporte")]:
        p = sub.add_parser(name, help=help_txt)
        if name in ("descargar", "todo"):
            p.add_argument("--txt");  p.add_argument("--dest")
            p.add_argument("--tipo", action="append", default=[])
            p.add_argument("--pruebas", dest="produccion", action="store_false")
        if name in ("reportar", "reportar-emitidos", "todo"):
            p.add_argument("--xmldir"); p.add_argument("--output")
            p.add_argument("--abrir", action="store_true")
        p.add_argument("--config", help="JSON config (desde Excel)")

    # Etapa 3: Descarga automatica via Selenium
    p_auto = sub.add_parser("auto",
        help="Descarga automatica: login portal SRI + XMLs + reporte")
    p_auto.add_argument("--ruc",   help="RUC del contribuyente")
    p_auto.add_argument("--clave", help="Clave del portal SRI")
    p_auto.add_argument("--desde", help="Fecha inicio YYYY-MM-DD")
    p_auto.add_argument("--hasta", help="Fecha fin   YYYY-MM-DD")
    p_auto.add_argument("--dest",  help="Directorio destino para XMLs")
    p_auto.add_argument("--output",help="Ruta del reporte Excel de salida")
    p_auto.add_argument("--visible", action="store_true",
                        help="Mostrar navegador Chrome (para depuracion)")
    p_auto.add_argument("--pruebas", dest="produccion", action="store_false")
    p_auto.add_argument("--config", help="JSON config (desde Excel)")

    sub.add_parser("version")
    args = parser.parse_args()

    if args.comando == "version":
        print(f"{config.NOMBRE} v{config.VERSION}"); sys.exit(0)

    output_json = None
    if hasattr(args, "config") and args.config:
        try:
            output_json = _cargar_config(args.config).get("output_json")
        except Exception:
            pass

    try:
        funcs = {
            "descargar":         cmd_descargar,
            "reportar":          cmd_reportar,
            "reportar-emitidos": cmd_reportar_emitidos,
            "todo":              cmd_todo,
            "auto":              cmd_auto,
        }
        resultado = funcs[args.comando](args)
        if output_json:
            _escribir_output(resultado, output_json)
        if not resultado.get("ok", True):
            logger.error(resultado.get("error", "error desconocido"))
            sys.exit(1)
    except KeyboardInterrupt:
        sys.exit(0)
    except Exception as e:
        logger.exception(f"Error inesperado: {e}")
        if output_json:
            _escribir_output({"ok": False, "error": str(e)}, output_json)
        sys.exit(1)


if __name__ == "__main__":
    main()
