Attribute VB_Name = "modDiagnostico"
' ============================================================
' modDiagnostico — Pega este codigo en un modulo nuevo en el
' VBA Editor (Alt+F11 > Insertar > Modulo) y ejecuta
' DiagnosticoRibbon con F5.
' ============================================================

Sub DiagnosticoRibbon()
    Dim msg As String
    msg = "========== DIAGNOSTICO RIBBON ==========" & vbCrLf

    ' --- 1. Informacion de Excel ---
    msg = msg & vbCrLf & "[1] Excel" & vbCrLf
    msg = msg & "  Version : " & Application.Version & vbCrLf
    msg = msg & "  Build   : " & Application.Build & vbCrLf
    msg = msg & "  OS      : " & Application.OperatingSystem & vbCrLf

    ' --- 2. Informacion del archivo ---
    msg = msg & vbCrLf & "[2] Archivo" & vbCrLf
    msg = msg & "  Nombre  : " & ThisWorkbook.Name & vbCrLf
    msg = msg & "  Ruta    : " & ThisWorkbook.Path & vbCrLf
    msg = msg & "  Tipo    : " & ThisWorkbook.FileFormat & vbCrLf  ' 52=xlsm

    ' --- 3. Estado de seguridad ---
    msg = msg & vbCrLf & "[3] Seguridad" & vbCrLf
    On Error Resume Next
    Dim vbp As Object
    Set vbp = ThisWorkbook.VBProject
    If Err.Number = 0 And Not vbp Is Nothing Then
        msg = msg & "  VBProject accesible : SI (" & vbp.VBComponents.Count & " componentes)" & vbCrLf
    Else
        msg = msg & "  VBProject accesible : NO (error " & Err.Number & ")" & vbCrLf
    End If
    On Error GoTo 0

    ' --- 4. CustomUI - verificar via archivo temporal ---
    msg = msg & vbCrLf & "[4] CustomUI en el ZIP" & vbCrLf
    Dim tmpPath As String
    tmpPath = Environ("TEMP") & "\__diag_xlsm_" & Format(Now, "HHmmss") & ".zip"
    On Error Resume Next
    FileCopy ThisWorkbook.FullName, tmpPath
    If Err.Number <> 0 Then
        msg = msg & "  No se pudo copiar el archivo para inspeccion." & vbCrLf
    Else
        ' Usar Shell + PowerShell para listar contenido del ZIP
        Dim ps As String
        Dim outPath As String
        outPath = Environ("TEMP") & "\__diag_zip_out.txt"
        ps = "powershell -NoProfile -Command """ & _
             "[System.IO.Compression.ZipFile]::OpenRead('" & tmpPath & "').Entries | " & _
             "Select-Object -ExpandProperty FullName | Out-File -Encoding utf8 '" & outPath & "'"""
        Shell ps, vbHide
        Application.Wait Now + TimeValue("00:00:03")

        ' Leer resultado
        Dim ff As Integer
        ff = FreeFile
        Dim zipContents As String
        Dim linea As String
        Open outPath For Input As #ff
        Do While Not EOF(ff)
            Line Input #ff, linea
            If InStr(linea, "customUI") > 0 Or InStr(linea, "LabelInfo") > 0 Or _
               InStr(linea, "persons") > 0 Or InStr(linea, "vbaProject") > 0 Then
                msg = msg & "  " & Trim(linea) & vbCrLf
            End If
        Loop
        Close #ff

        ' Leer _rels/.rels
        Dim relsPath As String
        relsPath = Environ("TEMP") & "\__diag_rels_out.txt"
        ps = "powershell -NoProfile -Command """ & _
             "Add-Type -A System.IO.Compression.FileSystem; " & _
             "$z=[System.IO.Compression.ZipFile]::OpenRead('" & tmpPath & "'); " & _
             "$e=$z.Entries | Where-Object {$_.FullName -eq '_rels/.rels'}; " & _
             "$r=[System.IO.StreamReader]$e.Open(); " & _
             "$r.ReadToEnd() | Out-File -Encoding utf8 '" & relsPath & "'; " & _
             "$r.Close(); $z.Dispose()"""
        Shell ps, vbHide
        Application.Wait Now + TimeValue("00:00:03")

        msg = msg & vbCrLf & "[5] _rels/.rels" & vbCrLf
        ff = FreeFile
        Open relsPath For Input As #ff
        Do While Not EOF(ff)
            Line Input #ff, linea
            If Len(Trim(linea)) > 0 Then
                msg = msg & "  " & Left(Trim(linea), 90) & vbCrLf
            End If
        Loop
        Close #ff

        Kill tmpPath
        Kill outPath
        Kill relsPath
    End If
    On Error GoTo 0

    ' --- 5. Politica de ribbon ---
    msg = msg & vbCrLf & "[6] Politica de ribbon" & vbCrLf
    Dim regOut As String
    regOut = Environ("TEMP") & "\__diag_reg.txt"
    Dim psReg As String
    psReg = "powershell -NoProfile -Command """ & _
            "Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' " & _
            "-ErrorAction SilentlyContinue | Select AccessVBOM,VBAWarnings | " & _
            "Format-List | Out-File -Encoding utf8 '" & regOut & "'; " & _
            "(Get-ItemProperty 'HKLM:\Software\Policies\Microsoft\Office\16.0\Excel\DisabledCmdBarItemsList' " & _
            "-ErrorAction SilentlyContinue) | Out-File -Append -Encoding utf8 '" & regOut & "'; " & _
            "(Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Common\Toolbars' " & _
            "-ErrorAction SilentlyContinue) | Format-List | Out-File -Append -Encoding utf8 '" & regOut & "'; " & _
            """
    Shell psReg, vbHide
    Application.Wait Now + TimeValue("00:00:03")
    ff = FreeFile
    On Error Resume Next
    Open regOut For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, linea
        If Len(Trim(linea)) > 0 Then
            msg = msg & "  " & Trim(linea) & vbCrLf
        End If
    Loop
    Close #ff
    Kill regOut
    On Error GoTo 0

    msg = msg & vbCrLf & "========================================"

    ' Guardar a archivo para no perderlo
    Dim outFile As String
    outFile = ThisWorkbook.Path & "\diagnostico_ribbon.txt"
    ff = FreeFile
    Open outFile For Output As #ff
    Print #ff, msg
    Close #ff

    MsgBox msg & vbCrLf & vbCrLf & "Resultado guardado en: " & outFile, _
           vbInformation, "Diagnostico Ribbon - " & ThisWorkbook.Name

End Sub
