VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCuentaCorrienteCliente 
   Caption         =   "Cuenta Corriente de Cliente"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdImprimirDetalle 
         Caption         =   "Imprimir"
         Height          =   855
         Left            =   3360
         Picture         =   "frmCuentaCorrienteCliente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3960
         Width           =   1335
      End
      Begin MSACAL.Calendar CalendarHasta 
         Height          =   2655
         Left            =   4080
         TabIndex        =   2
         Top             =   960
         Width           =   3735
         _Version        =   524288
         _ExtentX        =   6588
         _ExtentY        =   4683
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2015
         Month           =   5
         Day             =   23
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSACAL.Calendar CalendarDesde 
         Height          =   2655
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   3735
         _Version        =   524288
         _ExtentX        =   6588
         _ExtentY        =   4683
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2015
         Month           =   5
         Day             =   23
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCuentaCorrienteCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimirDetalle_Click()
    cn.Open
    Set rs = cn.Execute("SELECT IsNull(sum(Debe) - sum(Haber),0) as inicial FROM CuentaCorriente  where idCliente=" & idCliente & " and Fecha < '" & CalendarDesde.Value & "'")
    With CuentaCorrienteCliente.Sections("Sección4")
        .Controls("lblCliente").Caption = frmClientes.grclientes.TextMatrix(frmClientes.grclientes.Row, 1)
        .Controls("lblDesde").Caption = CalendarDesde.Value
        .Controls("lblHasta").Caption = CalendarHasta.Value
        .Controls("lblSaldoInicial").Caption = Format(rs!inicial, "0.00")
    End With
    
    With CuentaCorrienteCliente.Sections("Sección5")
        Set rs = cn.Execute("SELECT IsNull(sum(Debe) - sum(Haber),0) as final FROM CuentaCorriente  where idCliente=" & idCliente & " and Fecha <= '" & CalendarHasta.Value & "'")
        .Controls("lblSaldoFinal").Caption = Format(rs!final, "0.00")
    End With
        
     
  
    Set rs = cn.Execute("SELECT Fecha, Detalle,Debe,Haber FROM CuentaCorriente  where idCliente=" & idCliente & " and (Fecha between '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "')  order by Fecha")
    Set CuentaCorrienteCliente.DataSource = rs
    CuentaCorrienteCliente.WindowState = 2
    
    CuentaCorrienteCliente.Show 1
    
    cn.Close
End Sub

Private Sub Form_Load()
    CalendarDesde.Value = Date
    CalendarHasta.Value = Date
End Sub
