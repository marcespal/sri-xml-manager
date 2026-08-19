' instalar_vba.vbs
' Importa el codigo VBA (modMain.bas) en XML_Manager.xlsm
' Ejecutar con doble clic DESPUES de habilitar el Trust Center en Excel.
'
' Como habilitar Trust Center (una sola vez):
'   Excel > Archivo > Opciones > Centro de confianza >
'   Configuracion del centro de confianza > Configuracion de macros >
'   Marcar: "Confiar en el acceso al modelo de objetos del proyecto de VBA"

Option Explicit

Dim sBase, sXlsm, sVBA

sBase = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
sXlsm = sBase & "XML_Manager.xlsm"
sVBA  = sBase & "vba\modMain.bas"

' Verificar archivos
If Not FileExists(sXlsm) Then
    MsgBox "No se encontro XML_Manager.xlsm en:" & Chr(13) & sBase, 16, "Error"
    WScript.Quit 1
End If

If Not FileExists(sVBA) Then
    MsgBox "No se encontro vba\modMain.bas en:" & Chr(13) & sBase, 16, "Error"
    WScript.Quit 1
End If

Dim xl, wb
On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "No se pudo abrir Excel: " & Err.Description, 16, "Error"
    WScript.Quit 1
End If
On Error GoTo 0

xl.Visible       = False
xl.DisplayAlerts = False

Set wb = xl.Workbooks.Open(sXlsm)

' Eliminar modulo anterior si existe
On Error Resume Next
Dim comp
For Each comp In wb.VBProject.VBComponents
    If comp.Name = "modMain" Then
        wb.VBProject.VBComponents.Remove comp
        Exit For
    End If
Next
On Error GoTo 0

' Importar nuevo modulo
On Error Resume Next
wb.VBProject.VBComponents.Import sVBA
Dim errNum, errDesc
errNum  = Err.Number
errDesc = Err.Description
On Error GoTo 0

If errNum <> 0 Then
    xl.Quit
    Set wb = Nothing : Set xl = Nothing
    MsgBox "Error al importar VBA (codigo " & errNum & "):" & Chr(13) & errDesc & Chr(13) & Chr(13) & _
           "Solucion:" & Chr(13) & _
           "1. Abra Excel" & Chr(13) & _
           "2. Archivo > Opciones > Centro de confianza" & Chr(13) & _
           "3. Configuracion del centro de confianza > Configuracion de macros" & Chr(13) & _
           "4. Habilitar: 'Confiar en el acceso al modelo de objetos del proyecto de VBA'" & Chr(13) & _
           "5. Reinicie Excel y vuelva a ejecutar este script", _
           16, "Error de acceso VBA"
    WScript.Quit 1
End If

' Guardar
wb.Save
wb.Close False
xl.Quit
Set wb = Nothing
Set xl = Nothing

MsgBox "VBA instalado correctamente en XML_Manager.xlsm." & Chr(13) & Chr(13) & _
       "El archivo esta listo para usar.", 64, "SRI XML Manager"

WScript.Quit 0

' ─────────────────────────────────────────────
Function FileExists(ruta)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(ruta)
    Set fso = Nothing
End Function
