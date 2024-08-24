VERSION 5.00
Begin VB.Form frmFichaArticulo 
   Caption         =   "Ficha del Artículo"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAgregarRubro 
      Height          =   495
      Left            =   4920
      Picture         =   "frmFichaClientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdAgregarMarca 
      Height          =   495
      Left            =   4920
      Picture         =   "frmFichaClientes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2280
      Width           =   495
   End
   Begin VB.CheckBox chkEtiquetar 
      Alignment       =   1  'Right Justify
      Caption         =   "Etiquetar"
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "True"
         FalseValue      =   "False"
         NullValue       =   ""
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   7
      EndProperty
      Height          =   255
      Left            =   280
      TabIndex        =   19
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   6240
      Picture         =   "frmFichaClientes.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   4560
      Picture         =   "frmFichaClientes.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.ComboBox cmMarca 
      Height          =   315
      ItemData        =   "frmFichaClientes.frx":2328
      Left            =   1080
      List            =   "frmFichaClientes.frx":232F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2400
      Width           =   3735
   End
   Begin VB.ComboBox cmRubro 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox txtStock 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtPrecioCosto 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtPrecioVenta 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   375
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   2
      Top             =   960
      Width           =   7095
   End
   Begin VB.TextBox txtCodBarras 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Barras:"
      Height          =   195
      Left            =   480
      TabIndex        =   18
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Stock:"
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Top             =   3120
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Marca:"
      Height          =   195
      Left            =   480
      TabIndex        =   16
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Rubro:"
      Height          =   195
      Left            =   480
      TabIndex        =   15
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Precio de Costo:"
      Height          =   195
      Left            =   5880
      TabIndex        =   14
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Precio de Venta:"
      Height          =   195
      Left            =   5880
      TabIndex        =   13
      Top             =   1920
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Baras:"
      Height          =   195
      Left            =   -240
      TabIndex        =   11
      Top             =   360
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   6360
      TabIndex        =   10
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "frmFichaArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarMarca_Click()
    frmAgregarMarca.Show 1
End Sub

Private Sub cmdAgregarRubro_Click()
    frmAgregarRubro.Show 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If txtCodBarras.Text = "" Then MsgBox ("Debe especificar una codigo de barras"): Exit Sub
    If txtDescripcion.Text = "" Then MsgBox ("Debe especificar una descripción"): Exit Sub
    If cmRubro.ListIndex < 0 Then MsgBox ("Debe elejir un rubro"): Exit Sub
    If cmMarca.ListIndex < 0 Then MsgBox ("Debe elejir una marca"): Exit Sub
    If IsNumeric(txtPrecioVenta) = False Then MsgBox ("El precio de venta no es valido"): Exit Sub
    If IsNumeric(txtPrecioCosto) = False Then MsgBox ("El precio de costo no es valido"): Exit Sub
    If IsNumeric(txtStock) = False Then MsgBox ("El valor del stock no es valido"): Exit Sub
    Respuesta = MsgBox("¿Esta seguro de guardar el artículo?", vbYesNo, "Guardar")
    If Respuesta = vbNo Then Exit Sub
    cn.Open
    If Estado = "Modificando" Then
        cn.Execute ("GuardarArticulo " & Val(txtCodigo) & ",'" & txtCodBarras & "','" & txtDescripcion & "'," & Replace(Val(txtPrecioVenta), ",", ".") & "," & Replace(Val(txtPrecioCosto), ",", ".") & "," & Replace(Val(txtStock), ",", ".") & "," & cmRubro.ItemData(cmRubro.ListIndex) & "," & cmMarca.ItemData(cmMarca.ListIndex) & "," & chkEtiquetar)
    Else
        cn.Execute ("AgregaArticulo '" & txtCodBarras & "','" & txtDescripcion & "'," & Replace(Val(txtPrecioVenta), ",", ".") & "," & Replace(Val(txtPrecioCosto), ",", ".") & "," & Replace(Val(txtStock), ",", ".") & "," & cmRubro.ItemData(cmRubro.ListIndex) & "," & cmMarca.ItemData(cmMarca.ListIndex) & "," & chkEtiquetar)
    End If
    cn.Close
    Saltar = 0
    Unload Me
End Sub



Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    cn.Open
    Set rs = cn.Execute("VerRubros")
    Do While rs.EOF = False
        cmRubro.AddItem (rs!Rubro)
        cmRubro.ItemData(cmRubro.NewIndex) = rs!IdRubro
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set rs = cn.Execute("VerMarcas")
    cmMarca.Clear
    Do While rs.EOF = False
        cmMarca.AddItem (rs!Marca)
        cmMarca.ItemData(cmMarca.NewIndex) = rs!idMarca
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If Estado = "Agregando" Then
        txtCodigo = ""
        txtCodBarras = ""
        txtDescripcion = ""
        txtPrecioVenta = "0.00"
        txtPrecioCosto = "0.00"
        txtStock = "0"
        cmRubro.ListIndex = 0
        cmMarca.ListIndex = 0
        chkEtiquetar = 1
    Else
        Set rs = cn.Execute("VerArticulo " & idArticulo)
        If rs.EOF = False Then
            txtCodigo = rs!idArticulo
            txtCodBarras = rs!CodBar
            txtDescripcion = rs!Descripcion
            txtPrecioVenta = Format(rs!Venta, "0.00")
            txtPrecioCosto = Format(rs!Costo, "0.00")
            txtStock = rs!Stock
            For I = 0 To cmRubro.ListCount - 1
                If cmRubro.ItemData(I) = rs!IdRubro Then cmRubro.ListIndex = I
            Next I
            For I = 0 To cmMarca.ListCount - 1
                If cmMarca.ItemData(I) = rs!idMarca Then cmMarca.ListIndex = I
            Next I
        chkEtiquetar = 1
        End If
    End If
    cn.Close
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPrecioCosto_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrecioVenta_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Public Sub CargarRubros()
    If cn.State = 0 Then cn.Open
    Set rs = cn.Execute("VerRubros")
    cmRubro.Clear
    Do While rs.EOF = False
        cmRubro.AddItem (rs!Rubro)
        cmRubro.ItemData(cmRubro.NewIndex) = rs!IdRubro
        rs.MoveNext
    Loop
    Set rs = Nothing
    cn.Close
End Sub

Public Sub CargarMarcas()
    If cn.State = 0 Then cn.Open
    Set rs = cn.Execute("VerMarcas")
    cmMarca.Clear
    Do While rs.EOF = False
        cmMarca.AddItem (rs!Marca)
        cmMarca.ItemData(cmMarca.NewIndex) = rs!idMarca
        rs.MoveNext
    Loop
    Set rs = Nothing
    cn.Close
End Sub
