Attribute VB_Name = "Module1"
Public cn As ADODB.Connection
Public idUsuario As Integer
Public Usuario As String
Public idArticulo As Integer
Public idCliente As Integer
Public CerroCaja As Integer
Public PidiendoPrecio As Integer
Public Saltar As Integer
Public Estado As String
Public EligiendoCliente As Integer
Public portscan As Integer
Public portfiscal As Integer
Public idUsuarioPermiso As Integer
Public idComprobante As Integer
Public printerTicket As String
Public printerDefault As String


'Para usar archivos ini
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'funciones para que el formulario se vea en la barra de tareas
'Public Const WS_EX_APPWINDOW As Long = &H40000
'Public Const GWL_EXSTYLE As Long = (-20)
'Public Const SW_HIDE As Long = 0
'Public Const SW_SHOW As Long = 5
'Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Public m_bActivated As Boolean


Sub Main()
    'chequear configuracion regional
    'SeparadorDecimal = Format(0.1, "#. #")
    'SeparadorDecimal = IIf(InStr(SeparadorDecimal, ","), ",", ".")
    'If SeparadorDecimal = "," Then
    '   MsgBox ("La configuracion regional no es la recomendada" & Chr(13) & "Debe configurar el punto para separador decimal y la coma para separador de miles"): Exit Sub
    'End If
    
     'For Each x In Printers
      '  MsgBox (x.DeviceName)
    'Next
    
    Dim I As Integer
    Dim Est As String
    On Error GoTo noInicia
    Est = String$(50, " ")
    I = GetPrivateProfileString("Config", "srv", "", Est, Len(Est), "./config.ini")
    srv = Mid(Est, 1, Len(Trim(Est)) - 1)
    Est = String$(50, " ")
    I = GetPrivateProfileString("Config", "db", "", Est, Len(Est), "./config.ini")
    db = Mid(Est, 1, Len(Trim(Est)) - 1)
    Est = String$(50, " ")
    I = GetPrivateProfileString("Config", "us", "", Est, Len(Est), "./config.ini")
    us = Mid(Est, 1, Len(Trim(Est)) - 1)
    Est = String$(50, " ")
    I = GetPrivateProfileString("Config", "pw", "", Est, Len(Est), "./config.ini")
    pw = Mid(Est, 1, Len(Trim(Est)) - 1)
    Est = String$(50, " ")
    I = GetPrivateProfileString("Config", "portscan", "", Est, Len(Est), "./config.ini")
    portscan = Mid(Est, 1, Len(Trim(Est)) - 1)
    Est = String$(50, " ")
    I = GetPrivateProfileString("Config", "portfiscal", "", Est, Len(Est), "./config.ini")
    portfiscal = Mid(Est, 1, Len(Trim(Est)) - 1)

    Est = String$(200, " ")
    I = GetPrivateProfileString("Config", "StringConnection", "", Est, Len(Est), "./config.ini")
    StringConnection = Mid(Est, 1, Len(Trim(Est)) - 1)

    Est = String$(200, " ")
    I = GetPrivateProfileString("Config", "printerTicket", "", Est, Len(Est), "./config.ini")
    printerTicket = Mid(Est, 1, Len(Trim(Est)) - 1)
    
    Est = String$(200, " ")
    I = GetPrivateProfileString("Config", "printerDefault", "", Est, Len(Est), "./config.ini")
    printerDefault = Mid(Est, 1, Len(Trim(Est)) - 1)
    
    'para escribir ini
    'Dim I As Integer
    'Dim Est As String
    'Est = "Ejemplo - Apartado"
    'I = WritePrivateProfileString("Ejemplo", "Nombre", Est, "Ejemplo.ini")
    
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    'esta cadena es para conectar a sqlserver2000
    'cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Pwd=soloitenet;Initial Catalog=Ejemplo;Data Source=JDL\EXPRESS"
    'esta cadena es para sqlserver2008Express
    'cn.ConnectionString = "Provider=SQLNCLI10;Data Source=" & srv & ";Persist Security info=True;Initial Catalog=" & db & ";User ID=" & us & ";Password=" & pw
    
    cn.ConnectionString = StringConnection
    

    
    frmLoguin.Show
    Exit Sub
    
noInicia:
    MsgBox ("Error de configuración" & Chr(13) & "No se puede iniciar la aplicación" & Err.Description)
    Exit Sub
End Sub
