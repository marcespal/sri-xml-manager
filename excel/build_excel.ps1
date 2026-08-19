# build_excel.ps1
# Crea XML_Manager.xlsm completo:
#   1. Abre Excel via COM
#   2. Crea hojas con formato
#   3. Inyecta codigo VBA
#   4. Guarda como .xlsm
#   5. Inyecta customUI.xml en el ZIP
#
# Requisito: Microsoft Excel instalado

param(
    [string]$OutputPath = "",
    [switch]$NoAbrir
)

$ErrorActionPreference = "Stop"
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("PATH","User")

$SCRIPT_DIR  = Split-Path -Parent $MyInvocation.MyCommand.Path
$PROJECT_DIR = Split-Path -Parent $SCRIPT_DIR
$VBA_DIR     = Join-Path $SCRIPT_DIR "vba"
$UI_DIR      = Join-Path $SCRIPT_DIR "customUI"

if ($OutputPath -eq "") {
    $OutputPath = Join-Path $SCRIPT_DIR "XML_Manager.xlsm"
}

Write-Host ""
Write-Host "╔══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   SRI XML Manager - Build Excel v1.0    ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# ─────────────────────────────────────────────────────────────
# PASO 1: Habilitar acceso al modelo de objetos VBA (registro)
# ─────────────────────────────────────────────────────────────
Write-Host "[1/5] Configurando acceso VBA..." -ForegroundColor Yellow

$accessKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$backupVal  = $null

try {
    $backupVal = (Get-ItemProperty -Path $accessKey -Name "AccessVBOM" -ErrorAction SilentlyContinue).AccessVBOM
    Set-ItemProperty -Path $accessKey -Name "AccessVBOM" -Value 1 -Type DWord -Force
    Write-Host "   Acceso VBA habilitado." -ForegroundColor Green
} catch {
    Write-Host "   Advertencia: no se pudo configurar registro VBA. Se intentara de todas formas." -ForegroundColor DarkYellow
}

# ─────────────────────────────────────────────────────────────
# PASO 2: Crear workbook con Excel COM
# ─────────────────────────────────────────────────────────────
Write-Host "[2/5] Creando workbook con Excel..." -ForegroundColor Yellow

$xl = $null
$wb = $null

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible        = $false
    $xl.DisplayAlerts  = $false
    $xl.ScreenUpdating = $false

    $wb = $xl.Workbooks.Add()

    # ── Paleta de colores (RGB) ──
    $AZUL_OSCURO  = 0x1F4E79  # RGB(31, 78, 121)  - headers FC
    $CAFE_OSCURO  = 0x833C00  # RGB(131, 60, 0)   - headers CR
    $AZUL_CLARO   = 0xD6E4F0  # fila par
    $GRIS_FONDO   = 0xF2F2F2  # Config fondo
    $BLANCO       = 0xFFFFFF

    function Set-CellStyle {
        param($cell, $bold=$false, $size=10, $fgColor=$null, $bgColor=$null,
              $hAlign="xlLeft", $wrap=$false, $italic=$false)
        $cell.Font.Bold   = $bold
        $cell.Font.Size   = $size
        $cell.Font.Italic = $italic
        if ($fgColor -ne $null) { $cell.Font.Color = $fgColor }
        if ($bgColor -ne $null) {
            $cell.Interior.Color   = $bgColor
            $cell.Interior.Pattern = 1  # xlSolid
        }
        $cell.HorizontalAlignment = switch ($hAlign) {
            "xlCenter" { -4108 }
            "xlRight"  { -4152 }
            default    { -4131 }  # xlLeft
        }
        $cell.WrapText = $wrap
    }

    # ══════════════════════════════════════════════════════════
    # HOJA 1 — Inicio
    # ══════════════════════════════════════════════════════════
    $wsInicio = $wb.Sheets(1)
    $wsInicio.Name = "Inicio"
    $wsInicio.Tab.Color = $AZUL_OSCURO

    # Ocultar gridlines
    $wsInicio.Application.ActiveWindow.DisplayGridlines = $false

    # Columna A ancho
    $wsInicio.Columns("A").ColumnWidth = 3
    $wsInicio.Columns("B").ColumnWidth = 60

    # Titulo
    $r = $wsInicio.Range("B2")
    $r.Value = "SRI XML Manager"
    Set-CellStyle $r -bold $true -size 22 -fgColor $AZUL_OSCURO
    $wsInicio.Rows("2").RowHeight = 35

    $r = $wsInicio.Range("B3")
    $r.Value = "Descarga automatica de comprobantes electronicos del SRI Ecuador"
    Set-CellStyle $r -size 12 -fgColor 0x595959 -italic $true
    $wsInicio.Rows("3").RowHeight = 20

    # Linea separadora
    $wsInicio.Range("B4").Value = ""
    $wsInicio.Range("B4").Borders([System.Runtime.InteropServices.Marshal]::GetActiveObject -as [type])
    # Borde inferior en B4
    $wsInicio.Range("B2:B4").Borders(9).LineStyle = 1  # xlContinuous bottom
    $wsInicio.Range("B2:B4").Borders(9).Color     = $AZUL_OSCURO
    $wsInicio.Range("B2:B4").Borders(9).Weight    = 3  # xlThick

    # Como usar
    $r = $wsInicio.Range("B6")
    $r.Value = "COMO USAR"
    Set-CellStyle $r -bold $true -size 11 -fgColor $AZUL_OSCURO

    $pasos = @(
        @("Paso 1", "Descargue el archivo TXT de comprobantes del portal SRI`n          (srienlinea.sri.gob.ec > Comprobantes > Recibidos > Descargar TXT)"),
        @("Paso 2", "Haga clic en '1. Cargar TXT' en la pestana 'XML Manager' y seleccione el TXT."),
        @("Paso 3", "Haga clic en '2. Descargar XMLs'. El motor descargara automaticamente`n          cada comprobante usando el Web Service oficial del SRI."),
        @("Paso 4", "Haga clic en 'Facturas' o 'Retenciones' para generar el reporte Excel."),
        @("Paso 5", "El reporte se abrira automaticamente con toda la informacion organizada.")
    )

    $fila = 7
    foreach ($paso in $pasos) {
        $rLabel = $wsInicio.Cells($fila, 2)
        $rLabel.Value = $paso[0] + ":"
        Set-CellStyle $rLabel -bold $true -size 10 -fgColor $AZUL_OSCURO
        $wsInicio.Columns("B").ColumnWidth = 5

        $rText = $wsInicio.Cells($fila, 3)
        $rText.Value = $paso[1]
        Set-CellStyle $rText -size 10
        $rText.WrapText = $true

        $wsInicio.Rows($fila).RowHeight = 28
        $fila++
    }

    $wsInicio.Columns("B").ColumnWidth = 8
    $wsInicio.Columns("C").ColumnWidth = 75

    # Version al fondo
    $wsInicio.Cells(14, 2).Value = "Version: 1.0.0  |  Motor: sri_engine.exe  |  Requiere: sri_engine.exe en la misma carpeta"
    Set-CellStyle $wsInicio.Cells(14, 2) -size 8 -fgColor 0x999999 -italic $true

    # ══════════════════════════════════════════════════════════
    # HOJA 2 — Config
    # ══════════════════════════════════════════════════════════
    $wsConfig = $wb.Sheets.Add([System.Type]::Missing, $wsInicio)
    $wsConfig.Name = "Config"
    $wsConfig.Tab.Color = 0x375623  # verde oscuro

    $wsConfig.Columns("A").ColumnWidth = 3
    $wsConfig.Columns("B").ColumnWidth = 28
    $wsConfig.Columns("C").ColumnWidth = 65

    # Titulo
    $wsConfig.Range("B2").Value = "CONFIGURACION"
    Set-CellStyle $wsConfig.Range("B2") -bold $true -size 14 -fgColor $AZUL_OSCURO

    # Filas de config
    $configRows = @(
        @(4,  "RUC:",              "",             "Informativo. El motor lee el TXT del SRI."),
        @(6,  "Archivo TXT:",      "",             "Ruta al TXT descargado del portal SRI (se llena al hacer clic en '1. Cargar TXT')"),
        @(8,  "Carpeta XMLs:",     "",             "Carpeta donde se guardaran los archivos XML descargados"),
        @(10, "Carpeta Reportes:", "",             "Carpeta donde se guardaran los reportes Excel generados"),
        @(12, "Ambiente:",         "PRODUCCION",   "PRODUCCION o PRUEBAS")
    )

    foreach ($row in $configRows) {
        $fRow = $row[0]
        # Label
        $lbl = $wsConfig.Cells($fRow, 2)
        $lbl.Value = $row[1]
        Set-CellStyle $lbl -bold $true -size 10 -bgColor $GRIS_FONDO

        # Value
        $val = $wsConfig.Cells($fRow, 3)
        $val.Value = $row[2]
        Set-CellStyle $val -size 10
        $val.Interior.Color   = $BLANCO
        $val.Interior.Pattern = 1
        # Borde
        $val.BorderAround(1, 2, 1)  # xlContinuous, xlThin, black

        # Hint
        $hint = $wsConfig.Cells($fRow + 1, 3)
        $hint.Value = "  " + $row[3]
        Set-CellStyle $hint -size 8 -fgColor 0x888888 -italic $true

        $wsConfig.Rows($fRow).RowHeight = 22
    }

    # Dropdown Ambiente
    $dv = $wsConfig.Cells(12, 3).Validation
    $dv.Delete()
    $dv.Add(3, 1, 1, "PRODUCCION,PRUEBAS")  # xlValidateList
    $dv.InCellDropdown = $true

    # Nota al pie
    $wsConfig.Cells(16, 2).Value = "Nota: Los campos se actualizan automaticamente al usar los botones del ribbon."
    Set-CellStyle $wsConfig.Cells(16, 2) -size 9 -fgColor 0x666666 -italic $true

    # ══════════════════════════════════════════════════════════
    # HOJA 3 — FC_Recibidas
    # ══════════════════════════════════════════════════════════
    $wsFC = $wb.Sheets.Add([System.Type]::Missing, $wsConfig)
    $wsFC.Name = "FC_Recibidas"
    $wsFC.Tab.Color = $AZUL_OSCURO

    $colsFC = @(
        @("FECHA EMISION",          12),
        @("RUC EMISOR",             14),
        @("RAZON SOCIAL EMISOR",    30),
        @("NUMERO",                 16),
        @("RUC/CI COMPRADOR",       14),
        @("COMPRADOR",              28),
        @("BASE 0%",                11),
        @("BASE 12%",               11),
        @("BASE 14%",               11),
        @("BASE 15%",               11),
        @("BASE NO OBJETO",         13),
        @("BASE EXENTA",            11),
        @("SUBTOTAL",               11),
        @("DESCUENTO",              11),
        @("IVA",                    11),
        @("PROPINA",                10),
        @("TOTAL",                  11),
        @("CLAVE ACCESO",           50)
    )

    $wsFC.Application.ActiveWindow.FreezePanes = $false
    for ($i = 0; $i -lt $colsFC.Count; $i++) {
        $c = $wsFC.Cells(1, $i + 1)
        $c.Value = $colsFC[$i][0]
        Set-CellStyle $c -bold $true -size 9 -fgColor 0xFFFFFF -bgColor $AZUL_OSCURO -hAlign "xlCenter" -wrap $true
        $wsFC.Columns($i + 1).ColumnWidth = $colsFC[$i][1]
    }
    $wsFC.Rows(1).RowHeight = 32
    $wsFC.Range("A2").Select() | Out-Null
    $xl.ActiveWindow.FreezePanes = $true

    # Nota de datos
    $wsFC.Cells(2, 1).Value = "Los datos se generan automaticamente. Use el boton 'Facturas' en el ribbon."
    $wsFC.Cells(2, 1).Font.Italic = $true
    $wsFC.Cells(2, 1).Font.Color  = 0x888888
    $wsFC.Cells(2, 1).Font.Size   = 9

    # ══════════════════════════════════════════════════════════
    # HOJA 4 — CR_Recibidas
    # ══════════════════════════════════════════════════════════
    $wsCR = $wb.Sheets.Add([System.Type]::Missing, $wsFC)
    $wsCR.Name = "CR_Recibidas"
    $wsCR.Tab.Color = $CAFE_OSCURO

    $colsCR = @(
        @("FECHA EMISION",          12),
        @("RUC RETENEDOR",          14),
        @("RAZON RETENEDOR",        30),
        @("NUMERO RETENCION",       18),
        @("RUC/CI RETENIDO",        14),
        @("RAZON RETENIDO",         28),
        @("PERIODO FISCAL",         13),
        @("COD DOC SUSTENTO",       16),
        @("NUM DOC SUSTENTO",       20),
        @("FECHA DOC SUSTENTO",     16),
        @("TIPO IMPUESTO",          13),
        @("COD RETENCION",          13),
        @("BASE IMPONIBLE",         13),
        @("% RETENCION",            11),
        @("VALOR RETENIDO",         13),
        @("CLAVE ACCESO",           50)
    )

    for ($i = 0; $i -lt $colsCR.Count; $i++) {
        $c = $wsCR.Cells(1, $i + 1)
        $c.Value = $colsCR[$i][0]
        Set-CellStyle $c -bold $true -size 9 -fgColor 0xFFFFFF -bgColor $CAFE_OSCURO -hAlign "xlCenter" -wrap $true
        $wsCR.Columns($i + 1).ColumnWidth = $colsCR[$i][1]
    }
    $wsCR.Rows(1).RowHeight = 32
    $wsCR.Range("A2").Select() | Out-Null
    $xl.ActiveWindow.FreezePanes = $true

    $wsCR.Cells(2, 1).Value = "Los datos se generan automaticamente. Use el boton 'Retenciones' en el ribbon."
    $wsCR.Cells(2, 1).Font.Italic = $true
    $wsCR.Cells(2, 1).Font.Color  = 0x888888
    $wsCR.Cells(2, 1).Font.Size   = 9

    # Ir a hoja Inicio
    $wsInicio.Activate()

    # ── Eliminar hoja vacía por defecto ──
    $sheetsToDelete = @()
    foreach ($sh in $wb.Sheets) {
        if ($sh.Name -notin @("Inicio","Config","FC_Recibidas","CR_Recibidas")) {
            $sheetsToDelete += $sh
        }
    }
    foreach ($sh in $sheetsToDelete) { $sh.Delete() }

    # ════════════════════════════════════════════════════════
    # PASO 3: Inyectar codigo VBA
    # ════════════════════════════════════════════════════════
    Write-Host "[3/5] Inyectando codigo VBA..." -ForegroundColor Yellow

    $vbaFile = Join-Path $VBA_DIR "modMain.bas"
    if (Test-Path $vbaFile) {
        try {
            $vbCode = Get-Content $vbaFile -Raw -Encoding UTF8
            $comp = $wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
            $comp.Name = "modMain"
            $comp.CodeModule.AddFromString($vbCode)
            Write-Host "   modMain.bas importado OK." -ForegroundColor Green
        } catch {
            Write-Host "   Advertencia VBA: $($_.Exception.Message)" -ForegroundColor DarkYellow
            Write-Host "   Habilite 'Acceso al modelo de objetos del proyecto VBA' en:" -ForegroundColor DarkYellow
            Write-Host "   Excel > Opciones > Centro de confianza > Configuracion del centro de confianza > Configuracion de macros" -ForegroundColor DarkYellow
        }
    } else {
        Write-Host "   Advertencia: modMain.bas no encontrado en $vbaFile" -ForegroundColor DarkYellow
    }

    # ════════════════════════════════════════════════════════
    # PASO 4: Guardar como .xlsm
    # ════════════════════════════════════════════════════════
    Write-Host "[4/5] Guardando como .xlsm..." -ForegroundColor Yellow

    # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.SaveAs($OutputPath, 52)
    $wb.Close($false)
    $xl.Quit()

    Write-Host "   Guardado: $OutputPath" -ForegroundColor Green

} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    if ($xl -ne $null) { try { $xl.Quit() } catch {} }
    throw
} finally {
    # Restaurar registro VBA
    if ($backupVal -ne $null) {
        Set-ItemProperty -Path $accessKey -Name "AccessVBOM" -Value $backupVal -Type DWord -Force
    } elseif ($backupVal -eq $null) {
        # Si no existia, eliminar la clave
        try {
            Remove-ItemProperty -Path $accessKey -Name "AccessVBOM" -ErrorAction SilentlyContinue
        } catch {}
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [System.GC]::Collect()
}

# ════════════════════════════════════════════════════════
# PASO 5: Inyectar customUI.xml en el ZIP del .xlsm
# ════════════════════════════════════════════════════════
Write-Host "[5/5] Inyectando ribbon personalizado (customUI)..." -ForegroundColor Yellow

$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("PATH","User")

$injectScript = Join-Path $SCRIPT_DIR "inject_ribbon.py"
if (Test-Path $injectScript) {
    python $injectScript --xlsm $OutputPath --cui (Join-Path $UI_DIR "customUI.xml")
    if ($LASTEXITCODE -eq 0) {
        Write-Host "   Ribbon inyectado OK." -ForegroundColor Green
    } else {
        Write-Host "   Advertencia: Error al inyectar ribbon (codigo $LASTEXITCODE)." -ForegroundColor DarkYellow
    }
} else {
    Write-Host "   inject_ribbon.py no encontrado, omitiendo ribbon." -ForegroundColor DarkYellow
}

# ════════════════════════════════════════════════════════
# RESULTADO FINAL
# ════════════════════════════════════════════════════════
$size = [math]::Round((Get-Item $OutputPath).Length / 1KB, 0)
Write-Host ""
Write-Host "╔══════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║   BUILD EXITOSO                         ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "Archivo: $OutputPath ($size KB)"
Write-Host ""
Write-Host "Para distribuir al cliente:"
Write-Host "  XML_Manager_v1.0.zip"
Write-Host "    XML_Manager.xlsm   <- este archivo"
Write-Host "    sri_engine.exe     <- compilar con build.ps1 en raiz"
Write-Host ""

if (-not $NoAbrir) {
    $resp = Read-Host "Abrir el archivo ahora? [S/n]"
    if ($resp -eq "" -or $resp.ToUpper() -eq "S") {
        Start-Process $OutputPath
    }
}
