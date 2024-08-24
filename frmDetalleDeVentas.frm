VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmDetalleDeVentas 
   Caption         =   "Detalle de Ventas"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.CommandButton cmdInformedeGastos 
         Caption         =   "Detalle de Gastos"
         Height          =   1095
         Left            =   5160
         Picture         =   "frmDetalleDeVentas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimirDetallePorRubro 
         Caption         =   "Imprimir Total por Rubros"
         Height          =   1095
         Left            =   3240
         Picture         =   "frmDetalleDeVentas.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimirDetalle 
         Caption         =   "Imprimir Detalle"
         Height          =   1095
         Left            =   1200
         Picture         =   "frmDetalleDeVentas.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3840
         Width           =   1335
      End
      Begin MSACAL.Calendar CalendarHasta 
         Height          =   2655
         Left            =   4080
         TabIndex        =   1
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
         TabIndex        =   5
         Top             =   480
         Width           =   975
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
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmDetalleDeVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImprimirDetalle_Click()
    cn.Open
    'Set rs = cn.Execute("SELECT Fecha, Tipo, Numero, Neto, IVA, Total, idCliente, Puesto, pedido FROM Ventas")
    With DetalleDeVenta.Sections("Sección4")
        .Controls("lblDesde").Caption = CalendarDesde.Value
        .Controls("lblHasta").Caption = CalendarHasta.Value
    End With
    
    With DetalleDeVenta.Sections("Sección5")
        Set rs = cn.Execute("SELECT Sum(Neto) as neto, sum(IVA) as iva, sum(Total) as total FROM Ventas WHERE Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
        .Controls("lblNeto").Caption = Format(rs!neto, "#.00")
        .Controls("lblIva").Caption = Format(rs!iva, "#.00")
        .Controls("lblTotal").Caption = Format(rs!Total, "#.00")
    End With
        
        
    Set rs = cn.Execute("SELECT Fecha, Tipo, Numero, Neto, IVA, Total, idCliente, Puesto, pedido FROM Ventas WHERE Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
    Set DetalleDeVenta.DataSource = rs
    DetalleDeVenta.WindowState = 2
    
    DetalleDeVenta.Show 1
    
    cn.Close
End Sub

Private Sub cmdImprimirDetallePorRubro_Click()
    cn.Open
    'Set rs = cn.Execute("SELECT Fecha, Tipo, Numero, Neto, IVA, Total, idCliente, Puesto, pedido FROM Ventas")
    With DetalleDeVentaPorRubro.Sections("Sección4")
        .Controls("lblDesde").Caption = CalendarDesde.Value
        .Controls("lblHasta").Caption = CalendarHasta.Value
    End With
    
    With DetalleDeVentaPorRubro.Sections("Sección5")
        Set rs = cn.Execute("SELECT sum(Total) as total FROM Ventas WHERE Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
        .Controls("lblTotal").Caption = Format(rs!Total, "#.00")
    End With
        
        
    Set rs = cn.Execute("SELECT Rubros.Rubro, SUM(DetalleVenta.PrecioFinal) AS Total FROM Rubros INNER JOIN Articulos ON Rubros.idRubro = Articulos.idRubro INNER JOIN  DetalleVenta ON Articulos.idArticulo = DetalleVenta.idArticulo INNER JOIN  Ventas ON DetalleVenta.idVenta = Ventas.idVenta WHERE Ventas.Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'" & " GROUP BY Rubros.Rubro ORDER BY Rubros.Rubro")
    Set DetalleDeVentaPorRubro.DataSource = rs
    DetalleDeVentaPorRubro.WindowState = 2
    
    DetalleDeVentaPorRubro.Show 1
    
    cn.Close
End Sub

Private Sub cmdInformedeGastos_Click()
    cn.Open
    With ReporteDeGastosEntreFechas.Sections("Sección4")
        .Controls("lblDesde").Caption = CalendarDesde.Value
        .Controls("lblHasta").Caption = CalendarHasta.Value
    End With
    
    With ReporteDeGastosEntreFechas.Sections("Sección5")
        Set rs = cn.Execute("SELECT sum(efectivo) * (- 1) AS Total FROM Detalle_Caja WHERE (idtipodetalle = 4) and fecha between '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
        .Controls("lblTotal").Caption = Format(rs!Total, "#.00")
    End With
        
        
    Set rs = cn.Execute("SELECT fecha, efectivo * (- 1) AS importe, comentario FROM Detalle_Caja WHERE (idtipodetalle = 4) and fecha between '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'" & " ORDER BY fecha")
    Set ReporteDeGastosEntreFechas.DataSource = rs
    ReporteDeGastosEntreFechas.WindowState = 2
    
    ReporteDeGastosEntreFechas.Show 1
    
    cn.Close
End Sub

Private Sub Form_Load()
    CalendarDesde.Value = Date
    CalendarHasta.Value = Date
End Sub
