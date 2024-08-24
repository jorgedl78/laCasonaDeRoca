VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmClientes 
   Caption         =   "Clientes"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCobrar 
      Caption         =   "Cobrar"
      Height          =   855
      Left            =   6720
      Picture         =   "frmClientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCuentaCorriente 
      Caption         =   "Cta Cte"
      Height          =   855
      Left            =   5520
      Picture         =   "frmClientes.frx":0534
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   855
      Left            =   8280
      Picture         =   "frmClientes.frx":0895
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txtBusca 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grclientes 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   11456
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      GridColorFixed  =   255
      TextStyleFixed  =   3
      HighLight       =   2
      SelectionMode   =   1
      GridLineWidthFixed=   1
      FontWidthFixed  =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   855
      Left            =   9480
      Picture         =   "frmClientes.frx":115F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label lblDescripcion 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   7440
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Total de clientes:"
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblEncontrados 
      Caption         =   "lblEncontrados"
      Height          =   255
      Left            =   9600
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    cn.Open
    Set rs = cn.Execute("VerPermsisosDeUsuario " & idUsuario)
    If rs!ModificarClientes = 0 Then MsgBox ("Función no permitida"): cn.Close: Exit Sub
    cn.Close
    Estado = "Agregando"
    frmFichaCliente.Show 1
End Sub

Private Sub cmdCobrar_Click()
    idCliente = grclientes.TextMatrix(grclientes.Row, 0)
    frmRecibo.txtNombre = grclientes.TextMatrix(grclientes.Row, 1)
    frmRecibo.Show 1
End Sub

Private Sub cmdCuentaCorriente_Click()
    idCliente = grclientes.TextMatrix(grclientes.Row, 0)
    frmCuentaCorrienteCliente.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    BuscarClientes
End Sub

Private Sub Form_Load()
    grclientes.Cols = 7
    grclientes.ColWidth(0) = 700
    grclientes.ColWidth(1) = 3000
    grclientes.ColWidth(2) = 2500
    grclientes.ColWidth(3) = 2500
    grclientes.ColWidth(4) = 2500
    lblEncontrados = 0
    lblDescripcion = ""
End Sub

Private Sub grClientes_DblClick()
    EditarCliente
End Sub

Private Sub grClientes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then EditarCliente
End Sub

Private Sub grClientes_RowColChange()
    lblDescripcion = grclientes.TextMatrix(grclientes.Row, 1)
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then cmdSalir_Click
    If KeyAscii <> 13 Then Exit Sub
    BuscarClientes
End Sub
Sub EditarCliente()
    If grclientes.Rows > 1 Then
        If EligiendoCliente = 1 Then
            frmFacturador.lblIdCliente = grclientes.TextMatrix(grclientes.Row, 0)
            frmFacturador.lblCliente = grclientes.TextMatrix(grclientes.Row, 1)
            frmFacturador.lblCategoria = grclientes.TextMatrix(grclientes.Row, 4)
            frmFacturador.lblTipoDocumento = grclientes.TextMatrix(grclientes.Row, 5)
            frmFacturador.lblNumeroDocumento = grclientes.TextMatrix(grclientes.Row, 6)
            Unload Me
        Else
            If grclientes.TextMatrix(grclientes.Row, 0) = 1 Then Exit Sub 'no permito editar al cliente CONSUMIDOR FINAL
            idCliente = grclientes.TextMatrix(grclientes.Row, 0)
            ClienteSeleccionado = grclientes.Row
            Saltar = 1
            Estado = "Modificando"
            frmFichaCliente.Show 1
            If Saltar = 0 Then
                BuscarClientes
                grclientes.Row = ClienteSeleccionado
                grClientes_RowColChange
            End If
        End If
    End If
End Sub

Sub BuscarClientes()
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("ABMClientes '" & txtBusca & "'")
    lblEncontrados = rs.RecordCount
   ' Set grClientes.DataSource = rs
    With grclientes
    .Rows = 1
    .TextArray(0) = "Codigo"
    .TextArray(1) = "Nombre"
    .TextArray(2) = "Domicilio"
    .TextArray(3) = "Telefonos"
    .TextArray(4) = "Categoría"
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idCliente
        .TextMatrix(.Rows - 1, 1) = rs!Nombre
        .TextMatrix(.Rows - 1, 2) = rs!Domicilio
        .TextMatrix(.Rows - 1, 3) = rs!Telefonos
        .TextMatrix(.Rows - 1, 4) = rs!categoria
        .TextMatrix(.Rows - 1, 5) = rs!TipoDocumento
        .TextMatrix(.Rows - 1, 6) = rs!NumeroDocumento
        rs.MoveNext
        .FixedRows = 1
    Loop
    End With
    If rs.RecordCount > 0 Then
        grclientes.SetFocus
        grClientes_RowColChange
    Else
        txtBusca.SetFocus
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub
