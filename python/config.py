"""
config.py — Configuración centralizada del SRI XML Manager
"""
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ─── Directorios ───────────────────────────────────────────────────────────────
import sys as _sys

# Cuando corre desde sri_engine.exe (PyInstaller), usar la carpeta del .exe.
# Cuando corre en desarrollo, usar la raíz del proyecto.
if getattr(_sys, 'frozen', False):
    _EXE_DIR = Path(_sys.executable).parent   # carpeta donde está sri_engine.exe
else:
    _EXE_DIR = Path(__file__).parent.parent   # raíz del proyecto en desarrollo

BASE_DIR   = _EXE_DIR
PYTHON_DIR = BASE_DIR / "python"
XSLT_DIR   = BASE_DIR / "xslt"
DIST_DIR   = BASE_DIR / "dist"

WORK_DIR   = Path(os.getenv("SRI_WORK_DIR", _EXE_DIR / "trabajo")).resolve()
WORK_DIR.mkdir(parents=True, exist_ok=True)

XML_DIR    = WORK_DIR / "xmls"
OUTPUT_DIR = WORK_DIR / "reportes"
LOG_DIR    = WORK_DIR / "logs"

for d in [XML_DIR, OUTPUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ─── SRI SOAP Endpoints ────────────────────────────────────────────────────────
# El WAF del SRI bloquea descargas GET del WSDL desde clientes que no son
# navegadores (curl/Python/zeep dan "Connection reset"). Solución: empaquetar
# los WSDL como recursos LOCALES y configurar zeep para leerlos del disco.
# Mismo patrón que la librería Java "veronica-platform/veronica-soap".
#
# En el .exe (PyInstaller) los WSDL viven en _MEIPASS/wsdl/
# En desarrollo viven en python/wsdl/
if getattr(_sys, 'frozen', False):
    WSDL_DIR = Path(_sys._MEIPASS) / "wsdl"
else:
    WSDL_DIR = Path(__file__).parent / "wsdl"

SRI_WSDL_PRODUCCION = str(WSDL_DIR / "AutorizacionComprobantesOffline_prod.xml")
SRI_WSDL_PRUEBAS    = str(WSDL_DIR / "AutorizacionComprobantesOffline_pruebas.xml")
# Alias por compatibilidad con código existente
SRI_WSDL_OFFLINE    = SRI_WSDL_PRODUCCION

# URLs públicas (sólo informativas — NO se usan al runtime)
SRI_WSDL_URL_PRODUCCION = (
    "https://cel.sri.gob.ec/comprobantes-electronicos-ws/"
    "AutorizacionComprobantesOffline?wsdl"
)
SRI_WSDL_URL_PRUEBAS = (
    "https://celcer.sri.gob.ec/comprobantes-electronicos-ws/"
    "AutorizacionComprobantesOffline?wsdl"
)

# ─── SRI Portal URLs ───────────────────────────────────────────────────────────
SRI_URL_LOGIN     = (
    "https://srienlinea.sri.gob.ec/auth/realms/Internet/protocol/openid-connect/auth"
    "?client_id=app-sri-claves-angular"
    "&redirect_uri=https%3A%2F%2Fsrienlinea.sri.gob.ec%2Fsri-en-linea%2F%2Fcontribuyente%2Fperfil"
    "&response_mode=fragment&response_type=code&scope=openid"
)
SRI_URL_RECIBIDOS = (
    "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/"
    "pages/consultas/recibidos/comprobantesRecibidos.jsf"
)

# ─── Timeouts ──────────────────────────────────────────────────────────────────
SOAP_TIMEOUT      = 30
SELENIUM_TIMEOUT  = 20
PAGE_LOAD_TIMEOUT = 40

# ─── Tipos de comprobante SRI ──────────────────────────────────────────────────
TIPOS_COMPROBANTE = {
    "01": "FACTURA",
    "04": "NOTA_CREDITO",
    "05": "NOTA_DEBITO",
    "06": "GUIA_REMISION",
    "07": "RETENCION",
    "08": "LIQUIDACION_COMPRA",
}

TXT_COLUMNAS = [
    "tipo_identificacion", "ruc_receptor", "razon_receptor",
    "tipo_comprobante", "serie_establecimiento", "serie_punto_emision",
    "secuencial", "fecha_emision", "fecha_autorizacion",
    "numero_autorizacion", "importe_total", "estado",
]

VERSION = "1.0.0"
NOMBRE  = "SRI XML Manager"
