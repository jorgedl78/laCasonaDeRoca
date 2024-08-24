VERSION 5.00
Begin VB.Form frmMenuPrincipal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menú Principal"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   11175
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCaja 
      Caption         =   "Caja"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   360
      Picture         =   "frmMenuPrincipal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdClientes 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3960
      Picture         =   "frmMenuPrincipal.frx":2BB0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   1215
      Left            =   9360
      Picture         =   "frmMenuPrincipal.frx":4740
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Artículos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   360
      Picture         =   "frmMenuPrincipal.frx":500A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdFacturador 
      Caption         =   "Facturación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   7680
      Picture         =   "frmMenuPrincipal.frx":6727
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Menu mnArticulos 
      Caption         =   "Artículos"
      Begin VB.Menu mnAbmArticulos 
         Caption         =   "ABM Artículos"
      End
   End
   Begin VB.Menu mnVentas 
      Caption         =   "Ventas"
      Begin VB.Menu mnClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnFacturador 
         Caption         =   "Facturador"
      End
   End
   Begin VB.Menu mnInformes 
      Caption         =   "Informes"
      Begin VB.Menu mnHistorialDeCajas 
         Caption         =   "Historial de Cajas"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnDetalleDeVentas 
         Caption         =   "Detalle de Ventas"
      End
   End
   Begin VB.Menu mnUsuarios 
      Caption         =   "Usuarios"
      Begin VB.Menu mnUsuariosYpermisos 
         Caption         =   "Usuarios y Permisos"
      End
      Begin VB.Menu mnCambiarClave 
         Caption         =   "Cambiar mi Clave"
      End
   End
End
Attribute VB_Name = "frmMenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdCaja_Click()
    Dim rs As New ADODB.Recordset
    cn.Open
    Set rs = cn.Execute("SELECT Count(Apertura) As Total from Caja where Cerrada=0")
    
    If rs!Total = 0 Then
        cn.Close
        frmCajaApertura.Show 1
    Else
        cn.Close
        frmCajaCierre.Show 1
    End If
End Sub

Private Sub cmdClientes_Click()
    mnClientes_Click
End Sub

Private Sub cmdFacturador_Click()
    mnFacturador_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    mnAbmArticulos_Click
End Sub

Private Sub Form_Load()
    If Date >= DateValue("20/08/2024") Then End


    cn.Open
    Set rs = cn.Execute("VerPermsisosDeUsuario " & idUsuario)
    If rs!EstablecerUsuariosyPermisos = 1 Then
        mnUsuariosYpermisos.Visible = True
    Else
        mnUsuariosYpermisos.Visible = False
    End If
    If rs!VerInformes = 1 Then
        mnInformes.Visible = True
    Else
        mnInformes.Visible = False
    End If
    
    cn.Close
End Sub

Private Sub mnAbmArticulos_Click()
    cn.Open
    Set rs = cn.Execute("VerPermsisosDeUsuario " & idUsuario)
    If rs!ModificarArticulos = 0 Then MsgBox ("Función no permitida"): cn.Close: Exit Sub
    cn.Close
    frmArticulos.Show 1
End Sub

Private Sub mnCambiarClave_Click()
    frmCambiarClave.Show 1
End Sub

Private Sub mnClientes_Click()
    cn.Open
    Set rs = cn.Execute("VerPermsisosDeUsuario " & idUsuario)
    If rs!ModificarClientes = 0 Then MsgBox ("Función no permitida"): cn.Close: Exit Sub
    cn.Close
    frmClientes.Show 1
End Sub

Private Sub mnDetalleDeVentas_Click()
    frmDetalleDeVentas.Show 1
End Sub

Private Sub mnFacturador_Click()
    cn.Open
    Set rs = cn.Execute("VerPermsisosDeUsuario " & idUsuario)
    If rs!Facturar = 0 Then MsgBox ("Función no permitida"): cn.Close: Exit Sub
    cn.Close
    frmFacturador.Show 1
End Sub

Private Sub mnHistorialDeCajas_Click()
    frmHistoricoDeCajas.Show 1
End Sub

Private Sub mnUsuariosYpermisos_Click()
    frmUsuarios.Show 1
End Sub
