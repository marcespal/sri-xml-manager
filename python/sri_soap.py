"""
sri_soap.py — Cliente SOAP para el Web Service oficial del SRI Ecuador.

Descarga el XML de cualquier comprobante autorizado usando su clave de
acceso (49 dígitos). NO requiere autenticación de usuario.

Importante: el WSDL se carga desde un archivo LOCAL empaquetado en el .exe
(ver config.SRI_WSDL_PRODUCCION / SRI_WSDL_PRUEBAS).  El WAF del SRI bloquea
descargas GET del WSDL desde clientes no-navegador, pero las llamadas SOAP
reales (POST) sí pasan.  Mismo patrón que la librería Java Verónica.
"""
from __future__ import annotations

import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import requests
import zeep
from zeep.exceptions import Fault, TransportError
from zeep.transports import Transport
from lxml import etree

import config
from utils import get_logger, esperar

logger = get_logger("sri_soap")


class _RespuestaVaciaError(Exception):
    """El SRI respondió sin autorizaciones — transitorio, se debe reintentar."""


class _SRITransport(Transport):
    """
    Transport de zeep que NO sigue redirects.

    El balanceador del SRI a veces responde 302 hacia su IP interna
    (https://181.113.x.x/) cuyo certificado no es válido para esa IP,
    causando SSLError. Tratamos cualquier 3xx como error transitorio
    para que el wrapper de reintentos vuelva a llamar al dominio correcto.
    """
    def post(self, address, message, headers):
        response = self.session.post(
            address, data=message, headers=headers,
            timeout=self.operation_timeout,
            allow_redirects=False,
        )
        if 300 <= response.status_code < 400:
            destino = response.headers.get("Location", "?")
            raise TransportError(
                f"SRI redirigió ({response.status_code}) hacia {destino} — "
                f"transitorio, se reintentará",
                status_code=response.status_code,
            )
        return response


@dataclass
class ResultadoSOAP:
    clave_acceso: str
    exitoso: bool
    estado: str = ""
    fecha_autorizacion: str = ""
    ambiente: str = ""
    xml_comprobante: str = ""
    xml_elemento: object = None
    error: str = ""


class SRISOAPClient:
    def __init__(self, produccion: bool = True, timeout: int = None,
                 max_reintentos: int = 5, pausa_reintento: float = 2.0):
        self.produccion = produccion
        self.timeout = timeout or config.SOAP_TIMEOUT
        self.max_reintentos = max_reintentos
        self.pausa_reintento = pausa_reintento
        self._client = None

    @property
    def wsdl_url(self) -> str:
        """Ruta al archivo WSDL LOCAL (no descarga de la red)."""
        return config.SRI_WSDL_PRODUCCION if self.produccion else config.SRI_WSDL_PRUEBAS

    def _build_session(self) -> requests.Session:
        """
        Sesión HTTP reutilizable con headers tipo navegador y respeto al
        proxy del sistema. Mitiga filtros WAF que bloquean clientes Python.
        """
        s = requests.Session()
        s.trust_env = True  # respeta HTTP_PROXY/HTTPS_PROXY del SO
        s.headers.update({
            "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                           "AppleWebKit/537.36 (KHTML, like Gecko) "
                           "Chrome/120.0.0.0 Safari/537.36"),
            "Accept":          "text/xml,application/xml,application/soap+xml,*/*",
            "Accept-Language": "es-EC,es;q=0.9,en;q=0.8",
            "Cache-Control":   "no-cache",
        })
        return s

    def _get_client(self):
        """Crea el cliente UNA SOLA VEZ y lo reutiliza para todas las claves."""
        if self._client is None:
            modo = "PRODUCCIÓN" if self.produccion else "PRUEBAS"
            logger.info(f"Cargando WSDL local ({modo}): {self.wsdl_url}")
            if not Path(self.wsdl_url).exists():
                raise FileNotFoundError(f"WSDL local no encontrado: {self.wsdl_url}")
            session   = self._build_session()
            transport = _SRITransport(session=session, timeout=self.timeout,
                                      operation_timeout=self.timeout)
            settings  = zeep.Settings(strict=False, xml_huge_tree=True,
                                      force_https=False)
            self._client = zeep.Client(self.wsdl_url,
                                       transport=transport, settings=settings)
            logger.info("Cliente SOAP listo (WSDL local, sesión con UA Chrome).")
        return self._client

    def descargar(self, clave_acceso: str) -> ResultadoSOAP:
        clave_acceso = clave_acceso.strip()
        if len(clave_acceso) != 49 or not clave_acceso.isdigit():
            return ResultadoSOAP(clave_acceso=clave_acceso, exitoso=False,
                                 error=f"Clave invalida (len={len(clave_acceso)})")

        ultimo_error = ""
        for intento in range(1, self.max_reintentos + 1):
            try:
                return self._llamar_soap(clave_acceso)
            except (TransportError, TimeoutError, _RespuestaVaciaError) as e:
                ultimo_error = str(e)
                logger.warning(f"Intento {intento}/{self.max_reintentos}: {e}")
                # Fail-fast: si la respuesta es vacía y el comprobante tiene
                # más de 35 días, está fuera de la ventana del WS — no insistir.
                if isinstance(e, _RespuestaVaciaError) and intento >= 2:
                    dias = self._dias_desde_emision(clave_acceso)
                    if dias is not None and dias > 35:
                        break
                if intento < self.max_reintentos:
                    esperar(self.pausa_reintento * intento)
                    # NO recreamos el cliente: el WSDL es local, sólo reintentamos
            except Fault as e:
                return ResultadoSOAP(clave_acceso=clave_acceso, exitoso=False,
                                     error=f"SOAP Fault: {e.message}")
            except Exception as e:
                ultimo_error = str(e)
                logger.warning(f"Intento {intento}/{self.max_reintentos}: {e}")
                if intento < self.max_reintentos:
                    esperar(self.pausa_reintento * intento)

        # Mensaje claro si el comprobante probablemente salió de la ventana
        # de consulta del WS del SRI (~30 días desde la emisión).
        if "Respuesta vacia" in ultimo_error or "redirigi" in ultimo_error:
            dias = self._dias_desde_emision(clave_acceso)
            if dias is not None and dias > 30:
                return ResultadoSOAP(
                    clave_acceso=clave_acceso, exitoso=False,
                    error=(f"Comprobante de hace {dias} días: el WS del SRI solo "
                           f"sirve ~30 días. Use 'Descarga Automática' (portal) "
                           f"para obtenerlo."))

        return ResultadoSOAP(clave_acceso=clave_acceso, exitoso=False,
                             error=f"Fallo tras {self.max_reintentos} intentos: {ultimo_error}")

    @staticmethod
    def _dias_desde_emision(clave_acceso: str):
        """Días transcurridos desde la fecha de emisión embebida en la clave
        (primeros 8 dígitos, formato ddmmaaaa). None si no se puede parsear."""
        try:
            from datetime import datetime
            fecha = datetime.strptime(clave_acceso[:8], "%d%m%Y")
            return (datetime.now() - fecha).days
        except Exception:
            return None

    def _llamar_soap(self, clave_acceso: str) -> ResultadoSOAP:
        client = self._get_client()
        respuesta = client.service.autorizacionComprobante(
            claveAccesoComprobante=clave_acceso)

        if not respuesta or not respuesta.autorizaciones:
            # Transitorio (inestabilidad del SRI): lanzar para que se reintente
            raise _RespuestaVaciaError("Respuesta vacia del SRI")

        auth = respuesta.autorizaciones.autorizacion[0]
        estado = str(auth.estado or "").upper()

        if estado != "AUTORIZADO":
            return ResultadoSOAP(clave_acceso=clave_acceso, exitoso=False,
                                 estado=estado, error=f"Estado SRI: {estado}")

        xml_str = str(auth.comprobante or "")
        if not xml_str:
            return ResultadoSOAP(clave_acceso=clave_acceso, exitoso=False,
                                 estado=estado, error="XML vacio")

        try:
            xml_elemento = etree.fromstring(xml_str.encode("utf-8"))
        except Exception:
            xml_elemento = None

        return ResultadoSOAP(
            clave_acceso=clave_acceso, exitoso=True, estado=estado,
            fecha_autorizacion=str(getattr(auth, "fechaAutorizacion", "") or ""),
            ambiente=str(getattr(auth, "ambiente", "") or ""),
            xml_comprobante=xml_str, xml_elemento=xml_elemento,
        )

    def descargar_lote(self, claves: list, xml_dir: Path,
                       claves_ya_descargadas=None,
                       max_workers: int = 6, progreso=None) -> dict:
        """
        Descarga el lote EN PARALELO (max_workers a la vez) para acelerar
        volúmenes grandes. Cada hilo usa su propio cliente SOAP.

        progreso: callable opcional progreso(done, total, ok, err) para una
                  barra de avance externa.
        """
        import threading
        from concurrent.futures import ThreadPoolExecutor, as_completed

        xml_dir = Path(xml_dir)
        xml_dir.mkdir(parents=True, exist_ok=True)
        ya = claves_ya_descargadas or set()
        pendientes = [c for c in claves if c not in ya]
        total = len(pendientes)

        logger.info(f"Lote: {len(claves)} claves, {total} pendientes, "
                    f"{max_workers} en paralelo")
        resultados = {}
        ok = error = completados = 0

        if total == 0:
            if progreso:
                progreso(0, 0, 0, 0)
            return resultados

        lock = threading.Lock()
        tl   = threading.local()

        def _worker(clave):
            cli = getattr(tl, "cli", None)
            if cli is None:
                cli = SRISOAPClient(produccion=self.produccion, timeout=self.timeout,
                                    max_reintentos=self.max_reintentos,
                                    pausa_reintento=self.pausa_reintento)
                tl.cli = cli
            return clave, cli.descargar(clave)

        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futuros = [ex.submit(_worker, c) for c in pendientes]
            for fut in as_completed(futuros):
                try:
                    clave, res = fut.result()
                except Exception as e:
                    with lock:
                        error += 1; completados += 1
                        if progreso: progreso(completados, total, ok, error)
                    logger.warning(f"Error en worker: {e}")
                    continue
                with lock:
                    resultados[clave] = res
                    if res.exitoso:
                        (xml_dir / f"{clave}.xml").write_text(
                            _envolver_autorizacion(res), encoding="utf-8")
                        # liberar memoria (faltantes sólo usa res.error)
                        res.xml_comprobante = ""
                        res.xml_elemento = None
                        ok += 1
                    else:
                        error += 1
                        logger.warning(f"x {clave[:8]}: {res.error}")
                    completados += 1
                    if progreso:
                        progreso(completados, total, ok, error)

        logger.info(f"Lote completado: {ok} OK, {error} errores")
        return resultados


def _envolver_autorizacion(res: ResultadoSOAP) -> str:
    """
    Envuelve el comprobante interno en el sobre <autorizacion> del SRI, para
    conservar estado, número y fecha de autorización (que NO están dentro del
    comprobante). El parser sabe leer este formato y extraer el comprobante.
    """
    inner = res.xml_comprobante or ""
    # Quitar el prólogo <?xml ...?> del interno (irá dentro de CDATA igual)
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<autorizacion>'
        f'<estado>{res.estado or "AUTORIZADO"}</estado>'
        f'<numeroAutorizacion>{res.clave_acceso}</numeroAutorizacion>'
        f'<fechaAutorizacion>{res.fecha_autorizacion}</fechaAutorizacion>'
        f'<ambiente>{res.ambiente}</ambiente>'
        f'<comprobante><![CDATA[{inner}]]></comprobante>'
        '</autorizacion>'
    )
