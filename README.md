# SRI XML Manager

Herramienta para descargar y procesar comprobantes electrónicos del **SRI Ecuador**.
Combina una interfaz en Excel (`.xlsm` con ribbon propio) y un motor en Python
compilado (`sri_engine.exe`) que consume el Web Service oficial del SRI.

## Qué hace

A partir del TXT que exporta el portal del SRI:

- Descarga el **XML** de cada comprobante vía el Web Service oficial (sin navegador).
- Genera el **PDF (RIDE)** de cada comprobante.
- Nombra XML y PDF de forma descriptiva: `AAAA-MM-DD_RAZONSOCIAL_TIPO_numero`.
- Produce **reportes en Excel** con los valores desglosados.
- Marca en el reporte los comprobantes que **no se pudieron descargar** y por qué.

Soporta facturas, notas de crédito y débito, liquidaciones de compra y comprobantes
de retención, tanto **recibidos** como **emitidos**.

## Estructura

```
python/              Motor: CLI, SOAP, parsers, generador de PDF y Excel
  main.py            Punto de entrada (descargar | reportar | reportar-emitidos | auto)
  sri_soap.py        Cliente SOAP del SRI (WSDL local + descarga en paralelo)
  txt_parser.py      Lee el TXT del portal (uno o varios archivos)
  xml_parser.py      Parsea comprobantes (facturas, NC/ND, liquidaciones, retenciones)
  pdf_generator.py   Genera el RIDE en PDF
  excel_writer.py    Construye los reportes .xlsx
  progreso.py        Ventana de progreso (Tkinter)
  wsdl/              WSDL del SRI en local (ver "WSDL local" abajo)
excel/
  vba/modMain.bas    Macros del ribbon
  customUI/          Definición del ribbon
  build_excel.py     Construye XML_Manager.xlsm (requiere Excel instalado)
sri_engine.spec      Configuración de PyInstaller (build onedir)
```

## Requisitos

- Windows de **64 bits** (el motor se compila para x64; Windows de 32 bits no es compatible).
- Python 3.12
- Microsoft Excel (solo para *construir* el `.xlsm`, no para usarlo).

```bash
pip install -r requirements.txt
```

Para la descarga automática vía portal (Selenium) hace falta `chromedriver.exe`
compatible con el Chrome instalado, en la raíz del proyecto. No se versiona:
descárgalo de https://googlechromelabs.github.io/chrome-for-testing/

## Compilar

El motor **debe** compilarse en modo `onedir` (carpeta), no `onefile`: los antivirus
corporativos suelen borrar el ejecutable auto-extraíble por falso positivo. El
`.spec` ya está configurado así, con UPX desactivado por la misma razón.

Compila fuera de OneDrive para evitar bloqueos de sincronización:

```bash
pyinstaller sri_engine.spec --distpath C:\sri_build\dist --workpath C:\sri_build\work --noconfirm
```

Y el Excel (con Excel cerrado, porque usa automatización COM):

```bash
python excel/build_excel.py --output C:\sri_build\stage\XML_Manager.xlsm --no-abrir
```

Para distribuir: un ZIP con `XML_Manager.xlsm` y la carpeta `sri_engine` juntos.

## Instalación en el equipo del usuario

1. Descomprimir el ZIP en una carpeta **local** (por ejemplo `C:\SRI\`).
   No usar OneDrive: con el autoguardado activo, `ThisWorkbook.Path` devuelve una
   URL y las macros no pueden abrir archivos locales (error 52).
2. Antes de extraer: clic derecho al ZIP → Propiedades → **Desbloquear**.
3. Abrir `XML_Manager.xlsm` y habilitar macros.

Si el antivirus corporativo elimina `sri_engine.exe`, la solución correcta es una
**exclusión de TI** para esa carpeta o **firmar el ejecutable**.

## Uso

En la pestaña **XML Manager** del ribbon:

**Documentos Recibidos / Documentos Emitidos**
1. `Cargar TXT` (un archivo) o `Cargar Carpeta` (varios TXT a la vez, p. ej. uno
   por día del mes; las claves repetidas se omiten solas).
2. `Descargar` — corre en segundo plano con su propia ventana de progreso, así que
   Excel queda libre. Descarga en paralelo y genera los PDF.
3. `Reporte` — genera el Excel con los datos desglosados.

Los archivos se organizan así:

```
Documentos Recibidos/    Documentos Emitidos/
  XMLs/                    XMLs/
  PDFs/                    PDFs/
```

En **emitidos**, el nombre del archivo usa la **contraparte** (cliente o sujeto
retenido), no el emisor.

## Notas técnicas

**WSDL local.** El WAF del SRI bloquea la descarga del WSDL por `GET` desde
clientes que no son navegadores, aunque las llamadas SOAP (`POST`) sí pasan. Por eso
el WSDL viaja empaquetado en `python/wsdl/` y se carga desde disco — mismo enfoque
que la librería Java Verónica.

**Redirecciones intermitentes.** El balanceador del SRI a veces responde `302` hacia
una IP interna cuyo certificado no es válido para esa IP. El transporte propio
(`_SRITransport`) no sigue redirecciones y las trata como error transitorio para
reintentar contra el dominio correcto.

**Ventana de ~30 días.** El Web Service solo sirve comprobantes de aproximadamente
los últimos 30 días. Para los más antiguos hay que usar el portal.

## Privacidad

Los comprobantes del SRI contienen datos tributarios de empresas reales. El
`.gitignore` excluye XMLs, TXT, reportes y logs por ese motivo — no subir datos de
clientes al repositorio.
