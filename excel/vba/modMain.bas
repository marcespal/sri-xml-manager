Attribute VB_Name = "modMain"
Option Explicit

' ============================================================
'  SRI XML Manager v1.0 — Modulo principal
'  Llama al motor sri_engine.exe y comunica resultados
'  Etapa 3: Incluye descarga automatica via Selenium
' ============================================================

' El motor se distribuye en modo carpeta (onedir): sri_engine\sri_engine.exe
' Esto evita el falso positivo del antivirus corporativo con el exe onefile.
Private Const ENGINE_NAME  As String = "sri_engine.exe"
Private Const ENGINE_REL   As String = "sri_engine\sri_engine.exe"
Private Const CONFIG_FILE  As String = "config.json"
Private Const OUTPUT_FILE  As String = "output.json"
Private Const VERSION_STR  As String = "1.0.0"

' ─── Referencia al Ribbon (para actualizar estados) ───────
Private m_ribbon As Object

Public Sub SetRibbon(ribbon As Object)
    Set m_ribbon = ribbon
End Sub

' ============================================================
'  UTILIDADES DE ARCHIVO
' ============================================================

Private Function RutaBase() As String
    RutaBase = AsegurarRutaLocal(ThisWorkbook.Path) & "\"
End Function

' Convierte una ruta de OneDrive/SharePoint (https://...) a su ruta LOCAL.
' Si ya es local (C:\...) la devuelve sin cambios.
' Necesario porque cuando el Excel esta en OneDrive con Autoguardado activo,
' ThisWorkbook.Path devuelve una URL y VBA no puede abrir archivos en URLs
' (provoca el error 52 "Nombre o numero de archivo incorrecto").
Private Function AsegurarRutaLocal(ByVal p As String) As String
    If LCase(Left(p, 4)) <> "http" Then
        AsegurarRutaLocal = p
        Exit Function
    End If

    Dim resto As String
    resto = p
    ' Quitar protocolo (https://)
    If InStr(resto, "://") > 0 Then resto = Mid(resto, InStr(resto, "://") + 3)
    ' Quitar host (todo hasta el primer "/")
    If InStr(resto, "/") > 0 Then resto = Mid(resto, InStr(resto, "/"))

    ' Decodificar espacios y caracteres comunes de URL
    resto = Replace(resto, "%20", " ")
    resto = Replace(resto, "%C3%91", "Ñ")
    resto = Replace(resto, "%C3%A1", "á")
    resto = Replace(resto, "%C3%A9", "é")
    resto = Replace(resto, "%C3%AD", "í")
    resto = Replace(resto, "%C3%B3", "ó")
    resto = Replace(resto, "%C3%BA", "ú")

    ' Tomar la parte despues de /Documents/ o /Documentos/ (biblioteca raiz)
    Dim cola As String, marcadores As Variant, m As Variant, posM As Long
    cola = ""
    marcadores = Array("/Documents/", "/Documentos/")
    For Each m In marcadores
        posM = InStr(1, resto, CStr(m), vbTextCompare)
        If posM > 0 Then
            cola = Mid(resto, posM + Len(CStr(m)))
            Exit For
        End If
    Next m
    If cola = "" Then
        cola = resto
        If Left(cola, 1) = "/" Then cola = Mid(cola, 2)
    End If
    cola = Replace(cola, "/", "\")

    ' Probar cada raiz de OneDrive: buscar donde exista sri_engine.exe
    Dim raices(1 To 3) As String, i As Integer, cand As String
    raices(1) = Environ$("OneDriveCommercial")
    raices(2) = Environ$("OneDrive")
    raices(3) = Environ$("OneDriveConsumer")
    For i = 1 To 3
        If raices(i) <> "" Then
            ' a) raiz + cola
            cand = raices(i) & "\" & cola
            If FileExists(cand & "\" & ENGINE_REL) Or FileExists(cand & "\" & ENGINE_NAME) Then
                AsegurarRutaLocal = cand
                Exit Function
            End If
            ' b) raiz + Documentos + cola (algunas cuentas anidan "Documentos")
            cand = raices(i) & "\Documentos\" & cola
            If FileExists(cand & "\" & ENGINE_REL) Or FileExists(cand & "\" & ENGINE_NAME) Then
                AsegurarRutaLocal = cand
                Exit Function
            End If
        End If
    Next i

    ' No se pudo resolver: devolver la URL (VerificarEngine mostrara mensaje claro)
    AsegurarRutaLocal = p
End Function

Private Function RutaEngine() As String
    ' Preferir el modo carpeta (onedir); si no, el exe suelto (compatibilidad)
    If FileExists(RutaBase() & ENGINE_REL) Then
        RutaEngine = RutaBase() & ENGINE_REL
    Else
        RutaEngine = RutaBase() & ENGINE_NAME
    End If
End Function

Private Function RutaConfig() As String
    RutaConfig = RutaBase() & CONFIG_FILE
End Function

Private Function RutaOutput() As String
    RutaOutput = RutaBase() & OUTPUT_FILE
End Function

Private Function FileExists(ruta As String) As Boolean
    ' Usar On Error para tolerar rutas UNC/SharePoint/URL invalidas
    On Error Resume Next
    FileExists = (Len(Dir(ruta)) > 0)
    If Err.Number <> 0 Then FileExists = False
    Err.Clear
    On Error GoTo 0
End Function

' Verifica si una CARPETA existe (Dir() solo detecta archivos, no carpetas)
Private Function FolderExists(ruta As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(ruta) And vbDirectory) = vbDirectory)
    If Err.Number <> 0 Then FolderExists = False
    Err.Clear
    On Error GoTo 0
End Function

' Abre un selector de CARPETA y devuelve la ruta elegida ("" si se cancela)
Private Function SeleccionarCarpeta(titulo As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = titulo
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        SeleccionarCarpeta = fd.SelectedItems(1)
    Else
        SeleccionarCarpeta = ""
    End If
End Function

' Cuenta cuantos archivos *.txt hay directamente dentro de una carpeta
Private Function ContarTXT(carpeta As String) As Integer
    Dim n As Integer, f As String
    n = 0
    On Error Resume Next
    f = Dir(carpeta & "\*.txt")
    Do While f <> ""
        n = n + 1
        f = Dir
    Loop
    On Error GoTo 0
    ContarTXT = n
End Function

Private Function VerificarEngine() As Boolean
    ' Si la ruta sigue siendo una URL (OneDrive no resuelto), avisar claramente
    If LCase(Left(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Este archivo se esta ejecutando desde OneDrive/SharePoint en linea," & vbCrLf & _
               "lo que impide que la macro acceda a los archivos locales." & vbCrLf & vbCrLf & _
               "SOLUCION:" & vbCrLf & _
               "  1. Copie XML_Manager.xlsm y sri_engine.exe a una carpeta LOCAL" & vbCrLf & _
               "     (por ejemplo  C:\SRI\ )." & vbCrLf & _
               "  2. Abra el Excel desde esa carpeta local." & vbCrLf & vbCrLf & _
               "Tambien puede desactivar el 'Autoguardado' (arriba a la izquierda)" & vbCrLf & _
               "y volver a abrir el archivo.", _
               vbCritical, "SRI XML Manager — Ejecutar desde carpeta local"
        VerificarEngine = False
        Exit Function
    End If

    If Not FileExists(RutaEngine()) Then
        MsgBox "No se encontro el motor sri_engine." & vbCrLf & vbCrLf & _
               "Asegurese de que la carpeta 'sri_engine' (con sri_engine.exe" & vbCrLf & _
               "adentro) este junto a este archivo Excel:" & vbCrLf & RutaBase(), _
               vbCritical, "SRI XML Manager"
        VerificarEngine = False
    Else
        VerificarEngine = True
    End If
End Function

Private Function EscaparJSON(texto As String) As String
    Dim s As String
    s = Replace(texto, "\", "\\")
    s = Replace(s, """", "\""")
    EscaparJSON = s
End Function

Private Sub EscribirArchivo(ruta As String, contenido As String)
    Dim n As Integer
    n = FreeFile
    Open ruta For Output As #n
    Print #n, contenido
    Close #n
End Sub

Private Function LeerArchivo(ruta As String) As String
    If Not FileExists(ruta) Then
        LeerArchivo = "{}"
        Exit Function
    End If
    Dim n As Integer
    n = FreeFile
    Dim linea As String, contenido As String
    Open ruta For Input As #n
    Do While Not EOF(n)
        Line Input #n, linea
        contenido = contenido & linea
    Loop
    Close #n
    LeerArchivo = contenido
End Function

' ============================================================
'  EJECUTAR MOTOR PYTHON
' ============================================================

Private Function EjecutarPython(comando As String, _
                                  Optional mostrarVentana As Boolean = False) As Boolean
    If Not VerificarEngine() Then
        EjecutarPython = False
        Exit Function
    End If

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    Dim modo As Integer
    modo = IIf(mostrarVentana, 1, 0)  ' 0=oculto, 1=visible

    Dim cmdLine As String
    cmdLine = """" & RutaEngine() & """ " & comando

    Dim exitCode As Long
    exitCode = oShell.Run(cmdLine, modo, True)  ' True = esperar a que termine

    EjecutarPython = (exitCode = 0)
End Function

' Lanza el motor en SEGUNDO PLANO (no espera). Excel queda libre para usarse.
' El motor muestra su propia ventana de progreso. Retorna True si se lanzó.
Private Function LanzarPythonAsync(comando As String) As Boolean
    If Not VerificarEngine() Then
        LanzarPythonAsync = False
        Exit Function
    End If

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    Dim cmdLine As String
    cmdLine = """" & RutaEngine() & """ " & comando

    oShell.Run cmdLine, 0, False   ' 0 = sin ventana extra, False = NO esperar
    LanzarPythonAsync = True
End Function

' ============================================================
'  LEER CONFIGURACION DE LA HOJA CONFIG
' ============================================================

Private Function ObtenerConfig() As Object
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Hoja 'Config' no encontrada.", vbCritical
        Set ObtenerConfig = Nothing
        Exit Function
    End If

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    ' ─── Configuracion principal ──────────────────────────────
    d("ruc")        = Trim(ws.Range("C4").Value)
    d("txt")        = Trim(ws.Range("C6").Value)
    d("xml_dir")    = Trim(ws.Range("C8").Value)
    d("output_dir") = Trim(ws.Range("C10").Value)
    d("ambiente")   = Trim(ws.Range("C12").Value)
    d("produccion") = (UCase(Trim(ws.Range("C12").Value)) = "PRODUCCION")

    ' ─── Etapa 3: Descarga automatica ────────────────────────
    d("clave")   = Trim(ws.Range("C14").Value)   ' Clave SRI portal
    d("desde")   = FechaParaPython(ws.Range("C20").Value)  ' Fecha Desde
    d("hasta")   = FechaParaPython(ws.Range("C22").Value)  ' Fecha Hasta

    ' ─── v2: Documentos Emitidos ──────────────────────────────
    d("txt_emit")     = Trim(ws.Range("C26").Value)  ' TXT de emitidos
    d("xml_dir_emit") = Trim(ws.Range("C28").Value)  ' Carpeta XMLs emitidos

    Set ObtenerConfig = d
End Function

Private Sub ActualizarConfig(clave As String, valor As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    Select Case clave
        Case "txt":          ws.Range("C6").Value  = valor
        Case "xml_dir":      ws.Range("C8").Value  = valor
        Case "output_dir":   ws.Range("C10").Value = valor
        Case "txt_emit":     ws.Range("C26").Value = valor
        Case "xml_dir_emit": ws.Range("C28").Value = valor
    End Select
End Sub

' Convierte un valor de celda de fecha a formato YYYY-MM-DD para Python
Private Function FechaParaPython(valor As Variant) As String
    If IsEmpty(valor) Or valor = "" Then
        FechaParaPython = ""
        Exit Function
    End If
    On Error Resume Next
    Dim d As Date
    d = CDate(valor)
    If Err.Number = 0 Then
        FechaParaPython = Format(d, "YYYY-MM-DD")
    Else
        FechaParaPython = CStr(valor)
    End If
    On Error GoTo 0
End Function

' ============================================================
'  ACCIONES DEL RIBBON
' ============================================================

'── 1. Cargar TXT ────────────────────────────────────────────
Public Sub CargarTXT(control As IRibbonControl)
    Dim filtro As String
    filtro = "Archivo TXT (*.txt),*.txt,Todos (*.*),*.*"

    Dim ruta As Variant
    ruta = Application.GetOpenFilename(filtro, 1, "Seleccione el TXT descargado del SRI")

    If ruta = False Then Exit Sub

    ActualizarConfig "txt", CStr(ruta)

    MsgBox "Archivo TXT cargado." & vbCrLf & vbCrLf & _
           "Ahora haga clic en '2. Descargar XMLs'.", _
           vbInformation, "SRI XML Manager"
End Sub

'── 1b. Cargar CARPETA de TXT (mes completo) ─────────────────
' Permite descargar un mes entero cuando el SRI solo deja bajar un TXT por dia:
' se deja caer cada TXT diario en una carpeta y se procesan todos juntos.
Public Sub CargarCarpetaTXT(control As IRibbonControl)
    Dim carpeta As String
    carpeta = SeleccionarCarpeta("Seleccione la CARPETA con los TXT del SRI (uno por dia)")
    If carpeta = "" Then Exit Sub

    Dim n As Integer
    n = ContarTXT(carpeta)
    If n = 0 Then
        MsgBox "La carpeta seleccionada no contiene archivos .txt:" & vbCrLf & carpeta, _
               vbExclamation, "SRI XML Manager"
        Exit Sub
    End If

    ActualizarConfig "txt", carpeta
    MsgBox "Carpeta cargada: " & n & " archivo(s) TXT encontrado(s)." & vbCrLf & _
           carpeta & vbCrLf & vbCrLf & _
           "Se procesaran TODOS juntos (las claves repetidas se omiten)." & vbCrLf & _
           "Ahora haga clic en '2. Descargar XMLs'.", _
           vbInformation, "SRI XML Manager"
End Sub

'── 2. Descargar XMLs ────────────────────────────────────────
Public Sub DescargarXMLs(control As IRibbonControl)
    If Not VerificarEngine() Then Exit Sub

    Dim cfg As Object
    Set cfg = ObtenerConfig()
    If cfg Is Nothing Then Exit Sub

    Dim rutaTXT As String
    rutaTXT = cfg("txt")

    If rutaTXT = "" Then
        MsgBox "Primero cargue un TXT o una CARPETA de TXT del SRI (Paso 1).", vbExclamation
        Exit Sub
    End If

    ' La celda puede contener un ARCHIVO .txt o una CARPETA con varios TXT
    Dim esCarpeta As Boolean
    esCarpeta = FolderExists(rutaTXT)
    If Not esCarpeta And Not FileExists(rutaTXT) Then
        MsgBox "El TXT o carpeta no existe:" & vbCrLf & rutaTXT, vbCritical
        Exit Sub
    End If

    Dim destDir As String
    destDir = cfg("xml_dir")
    If destDir = "" Then
        destDir = RutaBase() & "Documentos Recibidos\XMLs"
        ActualizarConfig "xml_dir", destDir
    End If

    Dim produccion As Boolean
    produccion = cfg("produccion")

    Dim origenJson As String
    If esCarpeta Then
        origenJson = "  ""carpeta_txt"": """ & EscaparJSON(rutaTXT) & ""","
    Else
        origenJson = "  ""txt"": """ & EscaparJSON(rutaTXT) & ""","
    End If

    Dim jsonCfg As String
    jsonCfg = "{" & vbCrLf & _
        origenJson                                                & vbCrLf & _
        "  ""dest"": """   & EscaparJSON(destDir)          & """," & vbCrLf & _
        "  ""tipos"": [""fc"", ""cr""],"                           & vbCrLf & _
        "  ""produccion"": " & IIf(produccion, "true", "false") & "," & vbCrLf & _
        "  ""output_json"": """ & EscaparJSON(RutaOutput()) & """" & vbCrLf & _
        "}"
    EscribirArchivo RutaConfig(), jsonCfg

    If FileExists(RutaOutput()) Then Kill RutaOutput()

    ' Lanzar en SEGUNDO PLANO: Excel queda libre y el motor muestra su
    ' propia ventana de progreso.
    Dim ok As Boolean
    ok = LanzarPythonAsync("descargar --config """ & EscaparJSON(RutaConfig()) & """")

    If ok Then
        MsgBox "La descarga se está ejecutando en SEGUNDO PLANO." & vbCrLf & vbCrLf & _
               "Verás una ventana de progreso con el avance (X de N)." & vbCrLf & _
               "Puedes seguir usando Excel mientras tanto." & vbCrLf & vbCrLf & _
               "Cuando la ventana de progreso se cierre, los XML y PDF estarán" & vbCrLf & _
               "listos. Entonces genera el reporte con '3. Reporte Recibidos'.", _
               vbInformation, "SRI XML Manager"
    End If
End Sub

'── 3. Descarga Automatica (Etapa 3) ─────────────────────────
Public Sub DescargaAutomatica(control As IRibbonControl)
    If Not VerificarEngine() Then Exit Sub

    Dim cfg As Object
    Set cfg = ObtenerConfig()
    If cfg Is Nothing Then Exit Sub

    ' Validar datos requeridos
    Dim ruc As String, clave As String, desde As String, hasta As String
    ruc   = cfg("ruc")
    clave = cfg("clave")
    desde = cfg("desde")
    hasta = cfg("hasta")

    If ruc = "" Then
        MsgBox "Configure el RUC en la hoja Config (celda C4).", vbExclamation
        Exit Sub
    End If
    If clave = "" Then
        MsgBox "Configure la Clave SRI en la hoja Config (celda C14).", vbExclamation
        Exit Sub
    End If
    If desde = "" Or hasta = "" Then
        MsgBox "Configure las fechas 'Desde' y 'Hasta' en la hoja Config " & _
               "(celdas C20 y C22).", vbExclamation
        Exit Sub
    End If

    ' Confirmacion
    Dim claveMask As String
    claveMask = String(Len(clave), "*")  ' Enmascarar clave
    Dim resp As Integer
    resp = MsgBox("Iniciar descarga automatica?" & vbCrLf & vbCrLf & _
                  "RUC:    " & ruc & vbCrLf & _
                  "Clave:  " & claveMask & vbCrLf & _
                  "Desde:  " & desde & vbCrLf & _
                  "Hasta:  " & hasta & vbCrLf & vbCrLf & _
                  "El proceso abre Chrome, hace login al portal SRI," & vbCrLf & _
                  "descarga el TXT y los XMLs automaticamente." & vbCrLf & vbCrLf & _
                  "Puede tardar varios minutos.", _
                  vbQuestion + vbYesNo, "SRI XML Manager — Descarga Automatica")

    If resp <> vbYes Then Exit Sub

    Dim destDir As String
    destDir = cfg("xml_dir")
    If destDir = "" Then
        destDir = RutaBase() & "Documentos Recibidos\XMLs"
        ActualizarConfig "xml_dir", destDir
    End If

    Dim outputDir As String
    outputDir = cfg("output_dir")
    If outputDir = "" Then
        outputDir = RutaBase() & "reportes"
        ActualizarConfig "output_dir", outputDir
    End If

    Dim outputPath As String
    outputPath = outputDir & "\SRI_Reporte_" & Format(Now, "YYYYMMDD_HHmmss") & ".xlsx"

    ' Construir JSON de configuracion
    Dim jsonCfg As String
    jsonCfg = "{" & vbCrLf & _
        "  ""ruc"": """       & EscaparJSON(ruc)        & """," & vbCrLf & _
        "  ""clave"": """     & EscaparJSON(clave)      & """," & vbCrLf & _
        "  ""desde"": """     & EscaparJSON(desde)      & """," & vbCrLf & _
        "  ""hasta"": """     & EscaparJSON(hasta)      & """," & vbCrLf & _
        "  ""dest"": """      & EscaparJSON(destDir)    & """," & vbCrLf & _
        "  ""output"": """    & EscaparJSON(outputPath) & """," & vbCrLf & _
        "  ""produccion"": "  & IIf(cfg("produccion"), "true", "false") & "," & vbCrLf & _
        "  ""headless"": true," & vbCrLf & _
        "  ""output_json"": """ & EscaparJSON(RutaOutput()) & """" & vbCrLf & _
        "}"
    EscribirArchivo RutaConfig(), jsonCfg

    If FileExists(RutaOutput()) Then Kill RutaOutput()

    Application.StatusBar = "Descarga automatica en progreso... Abriendo Chrome."
    Application.Cursor = xlWait

    Dim ok As Boolean
    ok = EjecutarPython("auto --config """ & EscaparJSON(RutaConfig()) & """", _
                        False)  ' oculto — Chrome maneja su propia ventana

    Application.Cursor = xlDefault
    Application.StatusBar = False

    If ok And FileExists(outputPath) Then
        Dim resp2 As Integer
        resp2 = MsgBox("Descarga automatica completada exitosamente." & vbCrLf & vbCrLf & _
                       "Reporte generado: " & outputPath & vbCrLf & vbCrLf & _
                       "Abrir el reporte ahora?", _
                       vbQuestion + vbYesNo, "SRI XML Manager")
        If resp2 = vbYes Then
            Shell "explorer.exe """ & outputPath & """", vbHide
        End If
    ElseIf ok Then
        MsgBox "Descarga completada. Genere el reporte con '3. Reporte Recibidos'.", _
               vbInformation, "SRI XML Manager"
    Else
        MsgBox "Se produjeron errores en la descarga automatica." & vbCrLf & vbCrLf & _
               "Posibles causas:" & vbCrLf & _
               "  - Credenciales incorrectas (RUC o clave)" & vbCrLf & _
               "  - Portal SRI no disponible" & vbCrLf & _
               "  - No hay comprobantes en el periodo indicado" & vbCrLf & vbCrLf & _
               "Revise el log en la carpeta de trabajo.", _
               vbExclamation, "SRI XML Manager — Error"
    End If
End Sub

'── 4. Descarga Masiva Multi-RUC (Etapa 3+) ─────────────────
Public Sub DescargaMasivaMultiRUC(control As IRibbonControl)
    Const FILA_INICIO As Integer = 5   ' Primera fila de datos en la hoja
    Const FILA_MAX    As Integer = 24  ' Maximo 20 RUCs (filas 5-24)
    Const COL_RUC     As Integer = 2
    Const COL_CLAVE   As Integer = 3
    Const COL_DESDE   As Integer = 4
    Const COL_HASTA   As Integer = 5
    Const COL_ESTADO  As Integer = 6
    Const COL_XMLS    As Integer = 7
    Const COL_REPORTE As Integer = 8

    ' Verificar hoja Multi-RUC
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Multi-RUC")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Hoja 'Multi-RUC' no encontrada.", vbCritical, "SRI XML Manager"
        Exit Sub
    End If

    If Not VerificarEngine() Then Exit Sub

    ' Contar RUCs validos
    Dim totalRUCs As Integer
    Dim i As Integer
    For i = FILA_INICIO To FILA_MAX
        If Trim(ws.Cells(i, COL_RUC).Value) <> "" Then totalRUCs = totalRUCs + 1
    Next i

    If totalRUCs = 0 Then
        MsgBox "No hay RUCs configurados en la hoja 'Multi-RUC'." & vbCrLf & vbCrLf & _
               "Complete al menos una fila con: RUC, Clave, Fecha Desde y Fecha Hasta.", _
               vbExclamation, "SRI XML Manager"
        Exit Sub
    End If

    ' Confirmacion
    Dim resp As Integer
    resp = MsgBox("Iniciar descarga masiva para " & totalRUCs & " RUC(s)?" & vbCrLf & vbCrLf & _
                  "El proceso se ejecutara secuencialmente." & vbCrLf & _
                  "Tiempo estimado: ~2-5 minutos por RUC.", _
                  vbQuestion + vbYesNo, "SRI XML Manager — Descarga Masiva")
    If resp <> vbYes Then Exit Sub

    Application.Cursor = xlWait
    Dim procesados As Integer
    Dim errores    As Integer

    For i = FILA_INICIO To FILA_MAX
        Dim rucVal   As String: rucVal   = Trim(ws.Cells(i, COL_RUC).Value)
        If rucVal = "" Then GoTo SiguienteFila

        Dim claveVal As String: claveVal = Trim(ws.Cells(i, COL_CLAVE).Value)
        Dim desdeVal As String: desdeVal = FechaParaPython(ws.Cells(i, COL_DESDE).Value)
        Dim hastaVal As String: hastaVal = FechaParaPython(ws.Cells(i, COL_HASTA).Value)

        ' Validar fila
        If claveVal = "" Or desdeVal = "" Or hastaVal = "" Then
            ws.Cells(i, COL_ESTADO).Value      = "Incompleto"
            ws.Cells(i, COL_ESTADO).Font.Color = RGB(200, 100, 0)
            GoTo SiguienteFila
        End If

        ' Marcar como en proceso
        ws.Cells(i, COL_ESTADO).Value      = "Procesando..."
        ws.Cells(i, COL_ESTADO).Font.Color = RGB(0, 0, 180)
        ws.Cells(i, COL_XMLS).Value        = ""
        ws.Cells(i, COL_REPORTE).Value     = ""
        Application.StatusBar = "Procesando RUC " & rucVal & _
                                 " (" & (i - FILA_INICIO + 1) & "/" & totalRUCs & ")..."
        DoEvents

        ' Rutas especificas para este RUC
        Dim destDir   As String: destDir   = RutaBase() & "xmls\"   & rucVal
        Dim outputDir As String: outputDir = RutaBase() & "reportes"
        Dim outPath   As String
        outPath = outputDir & "\SRI_" & rucVal & "_" & _
                  Format(Now, "YYYYMMDD_HHmmss") & ".xlsx"

        ' Construir config.json
        Dim jCfg As String
        jCfg = "{" & vbCrLf & _
            "  ""ruc"":         """ & EscaparJSON(rucVal)   & """," & vbCrLf & _
            "  ""clave"":       """ & EscaparJSON(claveVal) & """," & vbCrLf & _
            "  ""desde"":       """ & EscaparJSON(desdeVal) & """," & vbCrLf & _
            "  ""hasta"":       """ & EscaparJSON(hastaVal) & """," & vbCrLf & _
            "  ""dest"":        """ & EscaparJSON(destDir)  & """," & vbCrLf & _
            "  ""output"":      """ & EscaparJSON(outPath)  & """," & vbCrLf & _
            "  ""produccion"":  true,"                              & vbCrLf & _
            "  ""headless"":    true,"                              & vbCrLf & _
            "  ""output_json"": """ & EscaparJSON(RutaOutput()) & """" & vbCrLf & _
            "}"
        EscribirArchivo RutaConfig(), jCfg
        If FileExists(RutaOutput()) Then Kill RutaOutput()

        ' Ejecutar motor
        Dim okRUC As Boolean
        okRUC = EjecutarPython("auto --config """ & EscaparJSON(RutaConfig()) & """")

        ' Parsear output.json para contar XMLs
        Dim jsonOut  As String: jsonOut = LeerArchivo(RutaOutput())
        Dim xmlsDesc As Long
        On Error Resume Next
        Dim pos As Long
        pos = InStr(jsonOut, """descargados"": ")
        If pos > 0 Then xmlsDesc = CLng(Trim(Split(Mid(jsonOut, pos + 16), ",")(0)))
        On Error GoTo 0

        ' Actualizar estado en la hoja
        If okRUC Then
            ws.Cells(i, COL_ESTADO).Value      = ChrW(10004) & " Completado"
            ws.Cells(i, COL_ESTADO).Font.Color = RGB(0, 128, 0)
            ws.Cells(i, COL_XMLS).Value        = xmlsDesc
            ws.Cells(i, COL_REPORTE).Value     = outPath
            procesados = procesados + 1
        Else
            ws.Cells(i, COL_ESTADO).Value      = ChrW(10008) & " Error"
            ws.Cells(i, COL_ESTADO).Font.Color = RGB(180, 0, 0)
            errores = errores + 1
        End If
        DoEvents

SiguienteFila:
    Next i

    Application.Cursor = xlDefault
    Application.StatusBar = False

    MsgBox "Descarga masiva finalizada." & vbCrLf & vbCrLf & _
           ChrW(10004) & " Completados : " & procesados & vbCrLf & _
           ChrW(10008) & " Con errores : " & errores & vbCrLf & vbCrLf & _
           "Reportes guardados en: " & RutaBase() & "reportes\", _
           vbInformation, "SRI XML Manager — Descarga Masiva"
End Sub

'── 3. Generar Reporte Recibidos (facturas + retenciones) ────
Public Sub GenerarReporteRecibidos(control As IRibbonControl)
    If Not VerificarEngine() Then Exit Sub

    Dim cfg As Object
    Set cfg = ObtenerConfig()
    If cfg Is Nothing Then Exit Sub

    Dim xmlDir As String
    xmlDir = cfg("xml_dir")
    If xmlDir = "" Then
        MsgBox "Primero descargue los XMLs (Paso 2 o Descarga Auto).", vbExclamation
        Exit Sub
    End If

    Dim outputDir As String
    outputDir = cfg("output_dir")
    If outputDir = "" Then
        outputDir = RutaBase() & "reportes"
        ActualizarConfig "output_dir", outputDir
    End If

    Dim outputPath As String
    outputPath = outputDir & "\SRI_Reporte_" & Format(Now, "YYYYMMDD_HHmmss") & ".xlsx"

    Dim jsonCfg As String
    jsonCfg = "{" & vbCrLf & _
        "  ""xml_dir"": """ & EscaparJSON(xmlDir)     & """," & vbCrLf & _
        "  ""output"": """  & EscaparJSON(outputPath)  & """," & vbCrLf & _
        "  ""output_json"": """ & EscaparJSON(RutaOutput()) & """" & vbCrLf & _
        "}"
    EscribirArchivo RutaConfig(), jsonCfg

    If FileExists(RutaOutput()) Then Kill RutaOutput()

    Application.StatusBar = "Generando reporte Excel... Por favor espere."
    Application.Cursor = xlWait

    Dim ok As Boolean
    ok = EjecutarPython("reportar --config """ & EscaparJSON(RutaConfig()) & """")

    Application.Cursor = xlDefault
    Application.StatusBar = False

    If ok And FileExists(outputPath) Then
        Dim resp As Integer
        resp = MsgBox("Reporte generado exitosamente." & vbCrLf & vbCrLf & _
                      outputPath & vbCrLf & vbCrLf & "Abrir ahora?", _
                      vbQuestion + vbYesNo, "SRI XML Manager")
        If resp = vbYes Then
            Shell "explorer.exe """ & outputPath & """", vbHide
        End If
    Else
        MsgBox "Error al generar el reporte." & vbCrLf & _
               "Verifique que existan XMLs en: " & xmlDir, _
               vbExclamation, "SRI XML Manager"
    End If
End Sub

' ============================================================
'  DOCUMENTOS EMITIDOS (v2)
' ============================================================

'── 1. Cargar TXT Emitidos ───────────────────────────────────
Public Sub CargarTXTEmitidos(control As IRibbonControl)
    Dim filtro As String
    filtro = "Archivo TXT (*.txt),*.txt,Todos (*.*),*.*"

    Dim ruta As Variant
    ruta = Application.GetOpenFilename(filtro, 1, _
                "Seleccione el TXT de comprobantes EMITIDOS del SRI")
    If ruta = False Then Exit Sub

    ActualizarConfig "txt_emit", CStr(ruta)
    MsgBox "TXT de emitidos cargado." & vbCrLf & vbCrLf & _
           "Ahora haga clic en '2. Descargar Emitidos'.", _
           vbInformation, "SRI XML Manager"
End Sub

'── 1b. Cargar CARPETA de TXT Emitidos (mes completo) ────────
Public Sub CargarCarpetaTXTEmitidos(control As IRibbonControl)
    Dim carpeta As String
    carpeta = SeleccionarCarpeta("Seleccione la CARPETA con los TXT de EMITIDOS (uno por dia)")
    If carpeta = "" Then Exit Sub

    Dim n As Integer
    n = ContarTXT(carpeta)
    If n = 0 Then
        MsgBox "La carpeta seleccionada no contiene archivos .txt:" & vbCrLf & carpeta, _
               vbExclamation, "SRI XML Manager"
        Exit Sub
    End If

    ActualizarConfig "txt_emit", carpeta
    MsgBox "Carpeta de emitidos cargada: " & n & " archivo(s) TXT." & vbCrLf & _
           carpeta & vbCrLf & vbCrLf & _
           "Se procesaran TODOS juntos (las claves repetidas se omiten)." & vbCrLf & _
           "Ahora haga clic en '2. Descargar Emitidos'.", _
           vbInformation, "SRI XML Manager"
End Sub

'── 2. Descargar Emitidos (XMLs + PDFs) ──────────────────────
Public Sub DescargarEmitidos(control As IRibbonControl)
    If Not VerificarEngine() Then Exit Sub

    Dim cfg As Object
    Set cfg = ObtenerConfig()
    If cfg Is Nothing Then Exit Sub

    Dim rutaTXT As String
    rutaTXT = cfg("txt_emit")
    If rutaTXT = "" Then
        MsgBox "Primero cargue el TXT o CARPETA de emitidos (Paso 1).", vbExclamation
        Exit Sub
    End If

    ' La celda puede contener un ARCHIVO .txt o una CARPETA con varios TXT
    Dim esCarpeta As Boolean
    esCarpeta = FolderExists(rutaTXT)
    If Not esCarpeta And Not FileExists(rutaTXT) Then
        MsgBox "El TXT o carpeta no existe:" & vbCrLf & rutaTXT, vbCritical
        Exit Sub
    End If

    Dim destDir As String
    destDir = RutaBase() & "Documentos Emitidos\XMLs"
    ActualizarConfig "xml_dir_emit", destDir

    Dim origenJson As String
    If esCarpeta Then
        origenJson = "  ""carpeta_txt"": """ & EscaparJSON(rutaTXT) & ""","
    Else
        origenJson = "  ""txt"": """ & EscaparJSON(rutaTXT) & ""","
    End If

    Dim jsonCfg As String
    jsonCfg = "{" & vbCrLf & _
        origenJson                                          & vbCrLf & _
        "  ""dest"": """   & EscaparJSON(destDir) & """," & vbCrLf & _
        "  ""emitidos"": true," & vbCrLf & _
        "  ""produccion"": true," & vbCrLf & _
        "  ""output_json"": """ & EscaparJSON(RutaOutput()) & """" & vbCrLf & _
        "}"
    EscribirArchivo RutaConfig(), jsonCfg
    If FileExists(RutaOutput()) Then Kill RutaOutput()

    ' Lanzar en SEGUNDO PLANO: Excel queda libre, ventana de progreso propia.
    Dim ok As Boolean
    ok = LanzarPythonAsync("descargar --config """ & EscaparJSON(RutaConfig()) & """")

    If ok Then
        MsgBox "La descarga de EMITIDOS se está ejecutando en SEGUNDO PLANO." & vbCrLf & vbCrLf & _
               "Verás una ventana de progreso con el avance (X de N)." & vbCrLf & _
               "Puedes seguir usando Excel mientras tanto." & vbCrLf & vbCrLf & _
               "Cuando la ventana de progreso se cierre, genera el reporte" & vbCrLf & _
               "con '3. Reporte Emitidos'.", _
               vbInformation, "SRI XML Manager"
    End If
End Sub

'── 3. Generar Reporte Emitidos ──────────────────────────────
Public Sub GenerarReporteEmitidos(control As IRibbonControl)
    If Not VerificarEngine() Then Exit Sub

    Dim cfg As Object
    Set cfg = ObtenerConfig()
    If cfg Is Nothing Then Exit Sub

    Dim xmlDir As String
    xmlDir = cfg("xml_dir_emit")
    If xmlDir = "" Then xmlDir = RutaBase() & "Documentos Emitidos\XMLs"

    If Not FolderExists(xmlDir) Then
        MsgBox "Primero descargue los emitidos (Paso 2).", vbExclamation
        Exit Sub
    End If

    Dim outputDir As String
    outputDir = cfg("output_dir")
    If outputDir = "" Then
        outputDir = RutaBase() & "reportes"
        ActualizarConfig "output_dir", outputDir
    End If

    Dim outputPath As String
    outputPath = outputDir & "\SRI_Emitidos_" & Format(Now, "YYYYMMDD_HHmmss") & ".xlsx"

    Dim jsonCfg As String
    jsonCfg = "{" & vbCrLf & _
        "  ""xml_dir"": """ & EscaparJSON(xmlDir)      & """," & vbCrLf & _
        "  ""output"": """  & EscaparJSON(outputPath)  & """," & vbCrLf & _
        "  ""output_json"": """ & EscaparJSON(RutaOutput()) & """" & vbCrLf & _
        "}"
    EscribirArchivo RutaConfig(), jsonCfg
    If FileExists(RutaOutput()) Then Kill RutaOutput()

    Application.StatusBar = "Generando reporte de emitidos... Por favor espere."
    Application.Cursor = xlWait

    Dim ok As Boolean
    ok = EjecutarPython("reportar-emitidos --config """ & EscaparJSON(RutaConfig()) & """")

    Application.Cursor = xlDefault
    Application.StatusBar = False

    If ok And FileExists(outputPath) Then
        Dim resp As Integer
        resp = MsgBox("Reporte de emitidos generado exitosamente." & vbCrLf & vbCrLf & _
                      outputPath & vbCrLf & vbCrLf & _
                      "Hojas: FC_Emitidas, NC_Emitidas, LIQ_Emitidas, RET_Emitidas." & vbCrLf & vbCrLf & _
                      "Abrir ahora?", _
                      vbQuestion + vbYesNo, "SRI XML Manager")
        If resp = vbYes Then
            Shell "explorer.exe """ & outputPath & """", vbHide
        End If
    Else
        MsgBox "Error al generar el reporte de emitidos." & vbCrLf & _
               "Verifique que existan XMLs en: " & xmlDir, _
               vbExclamation, "SRI XML Manager"
    End If
End Sub

'── Acerca de ─────────────────────────────────────────────────
Public Sub AcercaDe(control As IRibbonControl)
    MsgBox "SRI XML Manager v" & VERSION_STR & vbCrLf & vbCrLf & _
           "Descarga y procesa comprobantes electronicos del SRI Ecuador." & vbCrLf & vbCrLf & _
           "Comprobantes soportados:" & vbCrLf & _
           "  - Facturas recibidas (FC)" & vbCrLf & _
           "  - Comprobantes de retencion (CR)" & vbCrLf & vbCrLf & _
           "Etapa 3: Descarga automatica via portal SRI (Selenium)" & vbCrLf & vbCrLf & _
           "Motor: sri_engine.exe (Python 3.12)", _
           vbInformation, "Acerca de SRI XML Manager"
End Sub

'── Activar licencia (stub para Etapa 4) ─────────────────────
Public Sub ActivarLicencia(control As IRibbonControl)
    MsgBox "Sistema de licencias disponible en la version comercial.", _
           vbInformation, "Activar"
End Sub

' ============================================================
'  RIBBON CALLBACK - onLoad
' ============================================================
Public Sub RibbonOnLoad(ribbon As Object)
    Set m_ribbon = ribbon
End Sub
