VERSION 5.00
Begin VB.Form frmFichaCliente 
   Caption         =   "Ficha del Cliente"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumeroDocumento 
      Height          =   375
      Left            =   6720
      TabIndex        =   17
      Top             =   2400
      Width           =   1575
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   5280
      Picture         =   "frmFichaArticulo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   2880
      Picture         =   "frmFichaArticulo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   975
   End
   Begin VB.ComboBox cmCategorias 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtTelefonos 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox txtMail 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   3000
      Width           =   4335
   End
   Begin VB.TextBox txtDomicilio 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   6135
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "N�mero:"
      Height          =   255
      Left            =   6000
      TabIndex        =   16
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   3360
      TabIndex        =   15
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Telefonos:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Categor�a:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail:"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   3120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "C�digo:"
      Height          =   195
      Left            =   6480
      TabIndex        =   8
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "frmFichaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If txtNombre = "" Then MsgBox ("Debe especificar un nombre"): Exit Sub
    If cmCategorias = "Responsable Inscripto" And cmbTipo <> "CUIT" Then
        MsgBox ("Para esta categor�a el tipo debe ser CUIT")
        Exit Sub
    End If
    If cmCategorias = "Responsable Inscripto" And Len(txtNumeroDocumento) <> 11 Then
        MsgBox ("Es obligatorio un nro de CUIT v�lido para esta categor�a")
        Exit Sub
    End If
    Respuesta = MsgBox("�Esta seguro de guardar el cliente?", vbYesNo, "Guardar")
    If Respuesta = vbNo Then Exit Sub
    cn.Open
    If Estado = "Modificando" Then
        cn.Execute ("GuardarCliente '" & txtNombre & "','" & txtDomicilio & "','" & txtTelefonos & "','" & txtMail & "'," & cmCategorias.ItemData(cmCategorias.ListIndex) & "," & Val(txtCodigo) & "," & cmbTipo.ItemData(cmbTipo.ListIndex) & "," & Val(txtNumeroDocumento))
    Else
        cn.Execute ("AgregaCliente '" & txtNombre & "','" & txtDomicilio & "','" & txtTelefonos & "','" & txtMail & "'," & cmCategorias.ItemData(cmCategorias.ListIndex) & "," & cmbTipo.ItemData(cmbTipo.ListIndex) & "," & Val(txtn))
    End If
    cn.Close
    Saltar = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    cn.Open
    Set rs = cn.Execute("VerCategorias")
    Do While rs.EOF = False
        cmCategorias.AddItem (rs!categoria)
        cmCategorias.ItemData(cmCategorias.NewIndex) = rs!idCategoria
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set rs = cn.Execute("VerTiposDocumentos")
    Do While rs.EOF = False
        cmbTipo.AddItem (rs!TipoDocumento)
        cmbTipo.ItemData(cmbTipo.NewIndex) = rs!idTipoDocumento
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    If Estado = "Modificando" Then
        Set rs = cn.Execute("VerCliente " & idCliente)
        If rs.EOF = False Then
            txtCodigo = rs!idCliente
            txtNombre = rs!Nombre
            txtDomicilio = rs!Domicilio
            txtTelefonos = rs!Telefonos
            txtMail = rs!email
            txtNumeroDocumento = rs!NumeroDocumento
            For I = 0 To cmCategorias.ListCount - 1
                If cmCategorias.ItemData(I) = rs!idCategoria Then cmCategorias.ListIndex = I
            Next I
            For I = 0 To cmbTipo.ListCount - 1
                If cmbTipo.ItemData(I) = rs!idTipoDocumento Then cmbTipo.ListIndex = I
            Next I
        End If
    Else
        cmbTipo.ListIndex = 0
        cmCategorias.ListIndex = 0
        txtNombre = ""
        txtDomicilio = ""
        txtTelefonos = ""
        txtMail = ""
        txtNumeroDocumento = ""
    End If
    cn.Close
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNumeroDocumento_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
