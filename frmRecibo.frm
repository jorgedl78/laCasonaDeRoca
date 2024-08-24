VERSION 5.00
Begin VB.Form frmRecibo 
   Caption         =   "Recibo"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   0
      TabIndex        =   8
      Top             =   3840
      Width           =   8895
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   855
         Left            =   4560
         Picture         =   "frmRecibo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   855
         Left            =   2160
         Picture         =   "frmRecibo.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   5655
      Begin VB.TextBox txtImporte 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importe a cobrar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   2130
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   5175
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00FFFFFF&
         Height          =   855
         Left            =   6840
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo:"
         Height          =   255
         Left            =   6240
         TabIndex        =   3
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblSaldo 
         Alignment       =   2  'Center
         Caption         =   "lblSaldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If txtImporte = "" Then MsgBox ("Debe especificar importe"): Exit Sub
    Respuesta = MsgBox("¿Esta seguro de ingresar el pago?", vbYesNo, "Guardar")
    If Respuesta = vbNo Then Exit Sub
    cn.Open
    Set rs = cn.Execute("SELECT IsNull(MAX(numero),0) + 1 AS NuevoR FROM Recibos")
    NuevoNumero = rs!NuevoR
    cn.Execute ("INSERT INTO Recibos(Fecha,Numero,Importe,idCliente) VALUES ('" & Date & "'," & NuevoNumero & "," & Replace(txtImporte, ",", ".") & "," & idCliente & ")")
    cn.Execute ("AgregarCuentaCorriente '" & Format(Date, "yyyy/mm/dd") & "','Recibo " & Format(NuevoNumero, "00000000") & "',0," & Replace(txtImporte, ",", ".") & "," & idCliente & ",'Rec'," & NuevoNumero)
    cn.Close
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As Recordset
    cn.Open
    Set rs = cn.Execute("SELECT IsNull(sum(Debe) - sum(Haber),0) as saldo FROM CuentaCorriente  where idCliente=" & idCliente)
    txtImporte.Text = Format(rs!saldo, "0.00")
    lblSaldo = Format(rs!saldo, "0.00")
    cn.Close
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
