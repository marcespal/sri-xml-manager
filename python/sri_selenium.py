"""
sri_selenium.py — Descarga automática de XMLs desde el portal SRI.

Portal actual (2025+): Angular + Keycloak.
Flujo:
  1. Login Keycloak (usuario/password/kc-login)
  2. Menú hamburguesa → Buscar "comprobantes electronicos recibidos" → Click
  3. Dropdowns Año/Mes/Día → Consultar → tabla de resultados
  4. Por cada fila: click enlace XML → esperar descarga
  5. Repetir por cada mes del rango solicitado
"""
import os
import shutil
import sys
import time
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Optional, List

from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select as SeleniumSelect
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, ElementNotInteractableException,
    StaleElementReferenceException
)

import config
from utils import get_logger

logger = get_logger("sri_selenium")

# Nombres de mes en español (como aparecen en los dropdowns del portal SRI)
MESES_ES = {
    1: "Enero",    2: "Febrero",  3: "Marzo",    4: "Abril",
    5: "Mayo",     6: "Junio",    7: "Julio",    8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

# Tipos de comprobante del portal JSF (frmPrincipal:cmbTipoComprobante)
# Valor interno → texto visible
TIPOS_COMPROBANTE_JSF = [
    ("1", "Factura"),
    ("2", "Liquidación de compra"),
    ("3", "Notas de Crédito"),
    ("4", "Notas de Débito"),
    ("6", "Comprobante de Retención"),
]


def _chromedriver_path() -> Optional[str]:
    """Ruta a chromedriver.exe bundleado (PyInstaller) o None para Selenium Manager."""
    if getattr(sys, 'frozen', False):
        cd = Path(sys._MEIPASS) / 'chromedriver.exe'
        if cd.exists():
            return str(cd)
    return None


# ══════════════════════════════════════════════════════════════════════════════
# Resultado
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class ResultadoSelenium:
    """Resultado de la descarga automática vía Selenium."""
    exitoso:         bool
    xml_dir:         Optional[Path] = None   # directorio con XMLs descargados
    total_xmls:      int            = 0      # cantidad de XMLs descargados
    mensaje:         str            = ""
    capturas:        list           = field(default_factory=list)


# ══════════════════════════════════════════════════════════════════════════════
# Cliente Selenium SRI
# ══════════════════════════════════════════════════════════════════════════════

class SRISelenium:
    """
    Descarga XMLs de comprobantes recibidos desde el portal SRI (portal Angular 2025+).

    Flujo completo:
      1. Login Keycloak
      2. Menú → Buscar "comprobantes electronicos recibidos"
      3. Por cada mes del rango: seleccionar año/mes/Todos → Consultar → descargar XMLs
    """

    T_PAGINA   = 40
    T_ELEMENTO = 20
    T_DESCARGA = 45
    POLL       = 0.4

    def __init__(self, ruc: str, clave: str,
                 headless: bool = True,
                 debug_dir: Optional[Path] = None,
                 tipos: Optional[List[str]] = None):
        self.ruc       = ruc.strip()
        self.clave     = clave.strip()
        self.headless  = headless
        self.debug_dir = Path(debug_dir or (config.LOG_DIR / "selenium_debug"))
        self.driver: Optional[webdriver.Chrome] = None
        self._xml_dir: Optional[Path] = None
        # Tipos de comprobante a descargar (valores JSF). Por defecto: solo Facturas.
        self.tipos     = tipos or ["1"]
        # Banderas para guardar HTML de diagnóstico solo la primera vez
        self._tabla_dumpeada = False
        self._modal_dumpeado = False

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _capturar(self, nombre: str) -> Optional[Path]:
        if not self.driver:
            return None
        try:
            self.debug_dir.mkdir(parents=True, exist_ok=True)
            ts   = datetime.now().strftime("%H%M%S")
            ruta = self.debug_dir / f"{ts}_{nombre}.png"
            self.driver.save_screenshot(str(ruta))
            logger.debug(f"Screenshot: {ruta.name}")
            return ruta
        except Exception:
            return None

    def _dump_html(self, nombre: str):
        """Guarda el HTML de la página actual en xml_dir para diagnóstico."""
        try:
            ruta = (self._xml_dir.parent if self._xml_dir else Path(".")) / f"debug_{nombre}.html"
            ruta.write_text(self.driver.page_source, encoding="utf-8")
            logger.info(f"HTML guardado: {ruta}")
        except Exception as e:
            logger.warning(f"No pudo guardar HTML '{nombre}': {e}")

    def _wait(self, timeout: int = None) -> WebDriverWait:
        return WebDriverWait(self.driver, timeout or self.T_ELEMENTO,
                             poll_frequency=self.POLL)

    def _click_js(self, element):
        """Click via JavaScript como fallback."""
        self.driver.execute_script("arguments[0].click();", element)

    # ── Inspección de formulario vía JavaScript ──────────────────────────────

    def _inspeccionar_formulario_js(self) -> dict:
        """
        Usa JavaScript para enumerar TODOS los controles del formulario
        (selects, inputs, buttons) con sus IDs, nombres y valores.
        Retorna un dict con listas 'selects', 'inputs', 'buttons'.
        """
        try:
            import json as _json
            raw = self.driver.execute_script(r"""
                var info = {selects: [], inputs: [], buttons: []};
                document.querySelectorAll('select').forEach(function(s, i) {
                    var opts = Array.from(s.options).map(function(o) {
                        return {text: o.text.trim(), value: o.value};
                    });
                    info.selects.push({
                        index: i,
                        id:    s.id,
                        name:  s.name,
                        value: s.value,
                        opts:  opts
                    });
                });
                document.querySelectorAll('input, textarea').forEach(function(inp, i) {
                    info.inputs.push({
                        index:       i,
                        id:          inp.id,
                        name:        inp.name,
                        type:        inp.type,
                        value:       inp.value,
                        placeholder: inp.placeholder,
                        visible:     inp.offsetParent !== null
                    });
                });
                document.querySelectorAll('button, input[type="submit"], input[type="button"]')
                .forEach(function(b, i) {
                    info.buttons.push({
                        index: i,
                        id:    b.id,
                        type:  b.getAttribute('type') || b.type,
                        text:  (b.textContent || b.value || '').trim().substring(0, 80),
                        value: b.value || ''
                    });
                });
                return JSON.stringify(info);
            """)
            data = _json.loads(raw)
            logger.info(f"JS form: {len(data['selects'])} selects, "
                        f"{len(data['inputs'])} inputs, {len(data['buttons'])} buttons")
            for s in data['selects']:
                textos = [o['text'] for o in s['opts'][:8]]
                logger.info(f"  SELECT id={s['id']!r} name={s['name']!r} "
                            f"opts={textos}")
            for inp in data['inputs']:
                if inp.get('visible'):
                    logger.info(f"  INPUT  id={inp['id']!r} type={inp['type']!r} "
                                f"placeholder={inp['placeholder']!r}")
            for b in data['buttons']:
                logger.info(f"  BUTTON id={b['id']!r} type={b['type']!r} text={b['text']!r}")
            return data
        except Exception as e:
            logger.warning(f"_inspeccionar_formulario_js falló: {e}")
            return {"selects": [], "inputs": [], "buttons": []}

    def _set_select_js(self, select_id: str, option_text: str) -> bool:
        """Establece el valor de un <select> por texto de opción, via JS + dispatchEvent."""
        try:
            ok = self.driver.execute_script("""
                var sel = arguments[0];
                var text = arguments[1];
                var opts = Array.from(sel.options);
                var opt  = opts.find(function(o) {
                    return o.text.trim() === text || o.value === text;
                });
                if (!opt) {
                    // Búsqueda parcial insensible a mayúsculas
                    opt = opts.find(function(o) {
                        return o.text.trim().toLowerCase().indexOf(text.toLowerCase()) >= 0;
                    });
                }
                if (!opt) return false;
                sel.value = opt.value;
                sel.dispatchEvent(new Event('change', {bubbles: true}));
                sel.dispatchEvent(new Event('input',  {bubbles: true}));
                return true;
            """, self.driver.find_element(By.ID, select_id), option_text)
            if ok:
                logger.info(f"  select#{select_id} ← '{option_text}' (JS OK)")
            else:
                logger.warning(f"  select#{select_id}: opción '{option_text}' no encontrada")
            return bool(ok)
        except Exception as e:
            logger.warning(f"  _set_select_js({select_id!r}, {option_text!r}): {e}")
            return False

    def _set_select_by_elem(self, elem, option_text: str) -> bool:
        """Establece el valor de un elemento <select> por texto de opción, via JS."""
        try:
            ok = self.driver.execute_script("""
                var sel = arguments[0];
                var text = arguments[1];
                var opts = Array.from(sel.options);
                var opt  = opts.find(function(o) {
                    return o.text.trim() === text || o.value === text;
                });
                if (!opt) {
                    opt = opts.find(function(o) {
                        return o.text.trim().toLowerCase().indexOf(text.toLowerCase()) >= 0;
                    });
                }
                if (!opt) return false;
                sel.value = opt.value;
                sel.dispatchEvent(new Event('change', {bubbles: true}));
                sel.dispatchEvent(new Event('input',  {bubbles: true}));
                return true;
            """, elem, option_text)
            return bool(ok)
        except Exception as e:
            logger.warning(f"  _set_select_by_elem: {e}")
            return False

    def _init_driver(self, xml_dir: Path):
        import tempfile
        self._xml_dir = xml_dir
        xml_dir.mkdir(parents=True, exist_ok=True)

        opts = ChromeOptions()

        # Visible pero fuera de pantalla en modo headless para evitar
        # el crash del renderer en entornos corporativos con AV/EDR estricto
        opts.add_argument("--window-position=100,100")

        # Perfil temporal unico — evita conflictos con Chrome ya abierto
        tmp_profile = Path(tempfile.gettempdir()) / f"sri_chrome_{os.getpid()}"
        opts.add_argument(f"--user-data-dir={tmp_profile}")
        opts.add_argument("--no-first-run")
        opts.add_argument("--no-default-browser-check")

        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1400,900")
        opts.add_argument("--lang=es-EC,es;q=0.9")
        opts.add_argument("--disable-notifications")
        opts.add_argument("--ignore-certificate-errors")
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_argument("--disable-extensions")
        opts.add_argument("--disable-background-networking")
        opts.add_argument("--disable-default-apps")
        opts.add_argument("--disable-sync")
        opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        opts.add_experimental_option("useAutomationExtension", False)

        # Descarga automatica al directorio xml_dir
        prefs = {
            "download.default_directory":   str(xml_dir.resolve()),
            "download.prompt_for_download": False,
            "download.directory_upgrade":   True,
            "safebrowsing.enabled":         False,
            "profile.default_content_setting_values.automatic_downloads": 1,
        }
        opts.add_experimental_option("prefs", prefs)

        cd_path = _chromedriver_path()
        if cd_path:
            service = ChromeService(cd_path)
            self.driver = webdriver.Chrome(service=service, options=opts)
        else:
            self.driver = webdriver.Chrome(options=opts)

        self.driver.set_page_load_timeout(self.T_PAGINA)
        self.driver.implicitly_wait(2)
        logger.info(f"ChromeDriver iniciado.")

    # ── Login ────────────────────────────────────────────────────────────────

    def _login(self):
        """Login Keycloak: navega a la URL de autorizacion y rellena usuario/clave."""
        logger.info(f"Navegando a login SRI...")
        self.driver.get(config.SRI_URL_LOGIN)
        time.sleep(3)
        self._capturar("01_login_page")
        logger.info(f"URL: {self.driver.current_url} | Titulo: {self.driver.title}")

        # Campo usuario.
        # SRI Keycloak tiene id="usuario" (visible) e id="username" (hidden).
        # Intentamos el visible primero para no perder 20 s esperando que el
        # campo hidden sea clickeable (nunca lo será).
        sel_usuario = [
            (By.ID,   "usuario"),
            (By.NAME, "usuario"),
            (By.ID,   "username"),
            (By.NAME, "username"),
            (By.CSS_SELECTOR, "input[type='text']:not([type='hidden'])"),
        ]
        try:
            for by, sel in sel_usuario:
                try:
                    el = self._wait(20).until(EC.element_to_be_clickable((by, sel)))
                    el.clear(); el.send_keys(self.ruc)
                    logger.info(f"Usuario ingresado (selector: {sel})")
                    break
                except TimeoutException:
                    continue
            else:
                raise TimeoutException("Ningún selector de usuario funcionó")
        except TimeoutException:
            self._capturar("login_form_not_found")
            logger.error(f"URL al fallar: {self.driver.current_url}")
            raise Exception(
                "No se encontró el formulario de login en el portal SRI. "
                "Verifique la URL o si el portal está disponible."
            )

        # Campo contraseña
        try:
            pwd = self.driver.find_element(By.CSS_SELECTOR, "input[type='password']")
            pwd.clear(); pwd.send_keys(self.clave)
        except Exception:
            raise Exception("No se encontró el campo de contraseña.")

        self._capturar("02_credenciales")

        # Botón submit
        sel_btn = [
            (By.ID,           "kc-login"),
            (By.CSS_SELECTOR, "input[type='submit']"),
            (By.CSS_SELECTOR, "button[type='submit']"),
        ]
        for by, sel in sel_btn:
            try:
                btn = self.driver.find_element(by, sel)
                try:
                    btn.click()
                except ElementNotInteractableException:
                    self._click_js(btn)
                break
            except NoSuchElementException:
                continue

        # Esperar redireccion post-login (tuportal-internet o sri-en-linea)
        try:
            self._wait(30).until(lambda d:
                "tuportal-internet" in d.current_url or
                "sri-en-linea"      in d.current_url or
                "contribuyente"     in d.current_url
            )
            time.sleep(2)
        except TimeoutException:
            self._capturar("post_login_timeout")
            logger.error(f"URL post-login: {self.driver.current_url}")
            src = self.driver.page_source.lower()
            for txt in ["credenciales no válidas", "invalid", "incorrecta", "incorrecto"]:
                if txt in src:
                    raise Exception("Credenciales incorrectas (RUC o clave).")
            raise Exception(
                f"Timeout esperando redireccion post-login. "
                f"URL actual: {self.driver.current_url}"
            )

        self._capturar("03_post_login")
        logger.info(f"Login exitoso. URL: {self.driver.current_url}")
        # Guardar HTML del portal Angular para diagnóstico del menú
        self._dump_html("angular_portal")

    # ── Navegacion al modulo de comprobantes ─────────────────────────────────

    # ── Navegacion al modulo de comprobantes ─────────────────────────────────

    def _url_acceder_aplicacion(self) -> str:
        """
        Obtiene la URL del servlet de handoff que abre la app JSF de
        comprobantes recibidos CON sesión válida.

        El portal Angular renderiza el menú lateral con enlaces del tipo:
            /tuportal-internet/accederAplicacion.jspa?redireccion=57&idGrupo=55
        Ese servlet usa la sesión SSO activa para autenticarse en la app JSF
        (comprobantes-electronicos-internet) y luego redirige a
        comprobantesRecibidos.jsf.  La navegación directa a la .jsf NO funciona
        porque la app JSF tiene su propia sesión separada.

        Busca el enlace por su texto en el menú (más robusto ante cambios de ID)
        y, si no lo encuentra, usa la URL conocida como fallback.
        """
        FALLBACK = ("https://srienlinea.sri.gob.ec/tuportal-internet/"
                    "accederAplicacion.jspa?redireccion=57&idGrupo=55")
        try:
            # Los enlaces del menú están en el DOM aunque el sidebar esté colapsado.
            enlaces = self.driver.find_elements(
                By.XPATH, "//a[contains(@href,'accederAplicacion.jspa')]"
            )
            logger.info(f"Enlaces accederAplicacion en menú: {len(enlaces)}")
            for a in enlaces:
                # textContent funciona aunque el elemento esté oculto
                txt  = (a.get_attribute("textContent") or "").strip().lower()
                href = a.get_attribute("href") or ""
                if "recibid" in txt and "comprobante" in txt:
                    logger.info(f"Link 'Comprobantes recibidos' encontrado: {href}")
                    return href
        except Exception as e:
            logger.warning(f"Error buscando link en menú: {e}")

        logger.info(f"Usando URL fallback accederAplicacion: {FALLBACK}")
        return FALLBACK

    def _navegar_comprobantes_recibidos(self):
        """
        Navega a comprobantes recibidos a través del servlet de handoff
        accederAplicacion.jspa, que establece la sesión de la app JSF.
        """
        # Dar tiempo a que el portal Angular cargue el menú lateral
        time.sleep(4)

        url = self._url_acceder_aplicacion()
        logger.info(f"Navegando vía handoff: {url}")
        self.driver.get(url)

        # accederAplicacion.jspa hace varios redirects (tuportal → app JSF →
        # comprobantesRecibidos.jsf).  Esperamos hasta 60 s a aterrizar.
        try:
            self._wait(60).until(lambda d:
                "comprobantesRecibidos" in d.current_url
                or "comprobantes-electronicos-internet" in d.current_url
            )
            logger.info(f"JSF cargado: {self.driver.current_url}")
            time.sleep(3)
        except TimeoutException:
            cur_url = self.driver.current_url
            self._capturar("04_handoff_timeout")
            self._dump_html("comprobantes_timeout")

            # Si nos mandó a Keycloak login, reintentar con 2da autenticación
            if "auth/realms" in cur_url:
                logger.info("Handoff redirigió a Keycloak — autenticando de nuevo...")
                self._login_keycloak_simple()
                self.driver.get(url)
                try:
                    self._wait(60).until(lambda d:
                        "comprobantesRecibidos" in d.current_url
                        or "comprobantes-electronicos-internet" in d.current_url
                    )
                    logger.info(f"JSF cargado (2do intento): {self.driver.current_url}")
                    time.sleep(3)
                except TimeoutException:
                    cur_url = self.driver.current_url
                    self._dump_html("comprobantes_timeout2")
                    raise Exception(
                        f"No se pudo acceder a comprobantes recibidos tras handoff + "
                        f"2do login. URL final: {cur_url}. "
                        f"Revise debug_comprobantes_timeout2.html."
                    )
            else:
                raise Exception(
                    f"No se pudo acceder a comprobantes recibidos vía handoff. "
                    f"URL final: {cur_url}. Revise debug_comprobantes_timeout.html "
                    f"en la carpeta xmls del RUC."
                )

        self._capturar("05_jsf_cargado")
        self._dump_html("comprobantes_recibidos")
        self._inspeccionar_formulario_js()
        logger.info("Página de comprobantes recibidos cargada.")

    def _login_keycloak_simple(self):
        """Rellena y envía el formulario Keycloak (reutilizado en reintentos)."""
        try:
            self._wait(15).until(EC.presence_of_element_located((By.ID, "usuario")))
        except TimeoutException:
            logger.warning("No apareció formulario Keycloak en reintento")
            return
        try:
            el = self.driver.find_element(By.ID, "usuario")
            el.clear(); el.send_keys(self.ruc)
            pw = self.driver.find_element(By.ID, "password")
            pw.clear(); pw.send_keys(self.clave)
            btn = self.driver.find_element(By.ID, "kc-login")
            self._click_js(btn)
            logger.info("Reintento de login Keycloak enviado")
            time.sleep(3)
        except Exception as e:
            logger.warning(f"Error en reintento de login: {e}")

    # ── Filtros, consulta y descarga ─────────────────────────────────────────

    def _select_por_value(self, candidatos: List[str], value: str, desc: str) -> bool:
        """
        Selecciona un <select> JSF por value, probando varios ids candidatos.
        Usa XPath (más confiable que By.ID con ':' de JSF).
        """
        for sid in candidatos:
            try:
                el = self.driver.find_element(By.XPATH, f"//select[@id='{sid}']")
                SeleniumSelect(el).select_by_value(value)
                logger.debug(f"    {desc}={value} OK (#{sid})")
                return True
            except Exception:
                continue
        logger.warning(f"    No pudo setear {desc}={value} (ids: {candidatos})")
        return False

    def _seleccionar_filtros(self, year: int, month: int, tipo_val: str):
        """
        Rellena los filtros del formulario JSF (frmPrincipal) con los IDs reales
        capturados del DOM en vivo:
          frmPrincipal:ano               → value str(year)
          frmPrincipal:mes               → value str(month)  (1..12)
          frmPrincipal:dia               → value "0" (Todos)
          frmPrincipal:tipoComprobante   → value tipo_val   (1=Factura)
        """
        # Esperar a que el formulario de filtros esté presente
        try:
            self._wait(15).until(EC.presence_of_element_located(
                (By.XPATH, "//select[@id='frmPrincipal:ano']")))
        except TimeoutException:
            logger.warning("    No apareció el select de año (formulario no listo)")

        self._select_por_value(["frmPrincipal:ano"], str(year), "año")
        time.sleep(0.4)
        self._select_por_value(["frmPrincipal:mes"], str(month), "mes")
        time.sleep(0.4)
        self._select_por_value(["frmPrincipal:dia"], "0", "día")
        time.sleep(0.3)
        # El tipo puede llamarse tipoComprobante (DOM real) o cmbTipoComprobante (legacy)
        self._select_por_value(
            ["frmPrincipal:tipoComprobante", "frmPrincipal:cmbTipoComprobante"],
            tipo_val, "tipo")
        time.sleep(0.3)

    def _click_consultar(self) -> bool:
        """Clic en el botón Consultar (btnConsultar real, btnBuscar legacy, o por texto)."""
        for bid in ["frmPrincipal:btnConsultar", "frmPrincipal:btnBuscar"]:
            try:
                btn = self._wait(8).until(EC.element_to_be_clickable(
                    (By.XPATH, f"//button[@id='{bid}'] | //input[@id='{bid}']")))
                try:
                    btn.click()
                except ElementNotInteractableException:
                    self._click_js(btn)
                logger.info(f"    Consultar enviado (#{bid})")
                return True
            except TimeoutException:
                continue
        # Fallback por texto
        try:
            btn = self.driver.find_element(
                By.XPATH, "//button[contains(normalize-space(.),'Consultar')] "
                          "| //input[@type='submit' and @value='Consultar']")
            self._click_js(btn)
            logger.info("    Consultar enviado (por texto)")
            return True
        except NoSuchElementException:
            logger.error("    No se encontró el botón Consultar")
            return False

    def _consultar_mes(self, year: int, month: int) -> int:
        """
        Para el año/mes dado consulta los tipos de comprobante en self.tipos
        (por defecto solo Facturas), hace clic en Consultar y descarga los XMLs.
        Retorna el total de XMLs descargados en el mes.
        """
        mes_nombre = MESES_ES[month]
        logger.info(f"=== {mes_nombre} {year} ===")
        total_mes = 0

        # Mapa valor → nombre legible para los tipos solicitados
        nombres = dict(TIPOS_COMPROBANTE_JSF)

        for tipo_val in self.tipos:
            tipo_nombre = nombres.get(tipo_val, f"Tipo {tipo_val}")
            logger.info(f"  [{tipo_val}] {tipo_nombre}")
            try:
                # Rellenar filtros
                self._seleccionar_filtros(year, month, tipo_val)

                # Clic en Consultar
                self._click_consultar()

                time.sleep(4)
                self._capturar(f"res_{year}_{month:02d}_t{tipo_val}")

                # Diagnóstico: guardar HTML de la tabla de resultados la 1ra vez
                # para conocer la estructura exacta de los enlaces de descarga.
                if not self._tabla_dumpeada:
                    self._dump_html(f"tabla_resultados_{year}_{month:02d}_t{tipo_val}")
                    self._tabla_dumpeada = True

                n = self._descargar_xmls_tabla()
                logger.info(f"    → {n} XMLs")
                total_mes += n

            except Exception as e:
                logger.warning(f"  Error tipo {tipo_val}: {e}")
                self._capturar(f"err_{year}_{month:02d}_t{tipo_val}")

        return total_mes

    # ── Descarga de XMLs desde la tabla ─────────────────────────────────────

    # Filas de datos de la tabla PrimeFaces (id real: frmPrincipal:tablaCompRecibidos).
    # Las filas de datos tienen el atributo data-ri ("row index").
    _ROWS_XP = "//*[contains(@id,'tablaCompRecibidos')]//tr[@data-ri]"
    _NODATOS_XP = ("//*[contains(translate(normalize-space(.),"
                   "'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚ',"
                   "'abcdefghijklmnopqrstuvwxyzaeiou'),"
                   "'no se encontr') or contains(normalize-space(.),'No existen') "
                   "or contains(normalize-space(.),'sin resultado')]")

    def _esperar_descarga(self, archivos_antes: set, idx: int) -> bool:
        """Espera a que aparezca un nuevo .xml en el directorio de descargas."""
        fin = time.time() + self.T_DESCARGA
        while time.time() < fin:
            nuevos = set(self._xml_dir.glob("*.xml")) - archivos_antes
            nuevos = {
                f for f in nuevos
                if not (self._xml_dir / (f.name + ".crdownload")).exists()
                and f.stat().st_size > 0
            }
            if nuevos:
                f = sorted(nuevos, key=lambda p: p.stat().st_mtime)[-1]
                logger.info(f"      XML [{idx}]: {f.name}")
                return True
            time.sleep(self.POLL)
        logger.warning(f"      Timeout esperando XML #{idx}")
        return False

    def _descargar_fila(self, fila, idx: int) -> bool:
        """
        Descarga el XML de una fila.

        El ícono XML es:
          <a id="...:lnkXml" onclick="mojarra.jsfcljs(frmPrincipal,{...},'')">
        Ese onclick envía el form frmPrincipal y el servidor responde con el
        archivo .xml como descarga directa (NO abre modal).  Por eso basta con
        hacer clic y esperar que aparezca el archivo.
        """
        try:
            lnk = fila.find_element(By.XPATH, ".//a[contains(@id,':lnkXml')]")
        except NoSuchElementException:
            # Fallback: ícono cuya imagen es xml.gif
            try:
                lnk = fila.find_element(
                    By.XPATH, ".//a[.//img[contains(@src,'xml')]]")
            except NoSuchElementException:
                logger.debug(f"      Fila {idx}: sin ícono XML")
                return False

        archivos_antes = set(self._xml_dir.glob("*.xml"))

        # Clic en el ícono → dispara mojarra.jsfcljs → descarga directa
        try:
            lnk.click()
        except (ElementNotInteractableException, StaleElementReferenceException):
            self._click_js(lnk)

        return self._esperar_descarga(archivos_antes, idx)

    def _descargar_xmls_tabla(self) -> int:
        """
        Descarga todos los XMLs de la tabla de resultados (PrimeFaces datatable
        id=frmPrincipal:tablaCompRecibidos). Cada fila: clic en ícono :lnkXml →
        descarga directa. Pagina con span.ui-paginator-next.
        Retorna la cantidad de XMLs descargados en esta consulta.
        """
        total = 0

        # Esperar la tabla (con filas) o un mensaje de "sin datos"
        try:
            self._wait(15).until(lambda d:
                d.find_elements(By.XPATH, self._ROWS_XP)
                or d.find_elements(By.XPATH, self._NODATOS_XP)
            )
        except TimeoutException:
            logger.info("    Tabla de resultados no apareció en 15 s (sin datos)")
            return 0

        if not self.driver.find_elements(By.XPATH, self._ROWS_XP):
            logger.info("    Sin filas de datos para este filtro")
            return 0

        pagina = 1
        while True:
            filas = self.driver.find_elements(By.XPATH, self._ROWS_XP)
            if not filas:
                break
            n_filas = len(filas)
            logger.info(f"    Página {pagina}: {n_filas} filas")

            for i in range(n_filas):
                try:
                    # Re-fetch para evitar StaleElement
                    filas = self.driver.find_elements(By.XPATH, self._ROWS_XP)
                    if i >= len(filas):
                        break
                    if self._descargar_fila(filas[i], total + 1):
                        total += 1
                    time.sleep(0.4)
                except StaleElementReferenceException:
                    logger.debug(f"    StaleElement fila {i+1}, continúa")
                    continue
                except Exception as e:
                    logger.warning(f"    Error fila {i+1}: {e}")

            # ── Paginación PrimeFaces: span.ui-paginator-next no deshabilitado ─
            siguientes = self.driver.find_elements(
                By.XPATH,
                "//span[contains(@class,'ui-paginator-next') "
                "and not(contains(@class,'ui-state-disabled'))]")
            if not siguientes:
                break
            self._click_js(siguientes[0])
            pagina += 1
            time.sleep(2.5)
            logger.info(f"    → Página {pagina}")

        return total

    # ── API pública ──────────────────────────────────────────────────────────

    def descargar_xmls(self, desde: str, hasta: str,
                       xml_dir: Path) -> ResultadoSelenium:
        """
        Descarga XMLs del portal SRI directamente (sin SOAP).

        Args:
            desde:   Fecha inicio YYYY-MM-DD
            hasta:   Fecha fin   YYYY-MM-DD
            xml_dir: Directorio donde guardar los XMLs

        Returns:
            ResultadoSelenium con xml_dir y total_xmls si exitoso.
        """
        xml_dir = Path(xml_dir)

        try:
            self._init_driver(xml_dir)
            logger.info(f"=== Descarga SRI | {desde} → {hasta} ===")

            self._login()
            self._navegar_comprobantes_recibidos()

            # Iterar meses en el rango
            d_desde = datetime.strptime(desde, "%Y-%m-%d")
            d_hasta = datetime.strptime(hasta, "%Y-%m-%d")
            current = date(d_desde.year, d_desde.month, 1)
            end     = date(d_hasta.year, d_hasta.month, 1)

            total_xmls = 0
            while current <= end:
                n = self._consultar_mes(current.year, current.month)
                total_xmls += n
                logger.info(f"  {MESES_ES[current.month]} {current.year}: {n} XMLs total")

                # Avanzar al siguiente mes
                if current.month == 12:
                    current = date(current.year + 1, 1, 1)
                else:
                    current = date(current.year, current.month + 1, 1)

            logger.info(f"Total descargado: {total_xmls} XMLs en {xml_dir}")
            return ResultadoSelenium(
                exitoso=True,
                xml_dir=xml_dir,
                total_xmls=total_xmls,
                mensaje=f"{total_xmls} XMLs descargados",
            )

        except Exception as e:
            self._capturar("error_final")
            logger.error(f"Error: {e}")
            return ResultadoSelenium(
                exitoso=False,
                mensaje=str(e),
            )

        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass
            import tempfile
            tmp_profile = Path(tempfile.gettempdir()) / f"sri_chrome_{os.getpid()}"
            if tmp_profile.exists():
                shutil.rmtree(tmp_profile, ignore_errors=True)
