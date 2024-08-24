VERSION 5.00
Object = "{9C5C9460-5789-11DA-8CFB-0000E856BC17}#1.0#0"; "Fiscal051122.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFacturador 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9165
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMozo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8520
      TabIndex        =   36
      Top             =   8280
      Width           =   2775
   End
   Begin VB.TextBox txtPersonas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7080
      TabIndex        =   35
      Top             =   8280
      Width           =   1215
   End
   Begin VB.TextBox txtMesa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5640
      TabIndex        =   34
      Top             =   8280
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   255
      Left            =   600
      TabIndex        =   30
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   83492865
      CurrentDate     =   42724
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   28
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdBuscarCliente 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      Picture         =   "Ventas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   960
      Width           =   735
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin VB.CommandButton cmdCaja 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      MaskColor       =   &H00000000&
      Picture         =   "Ventas.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Caja"
      Top             =   7920
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5760
      TabIndex        =   1
      Top             =   375
      Width           =   735
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      MaskColor       =   &H00000000&
      Picture         =   "Ventas.frx":101F
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Aceptar Ticket"
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox txtBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6840
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grDetalle 
      Height          =   6135
      Left            =   5760
      TabIndex        =   14
      Top             =   1320
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10821
      _Version        =   393216
      BackColor       =   14737632
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      BackColorUnpopulated=   14737632
      GridColor       =   14737632
      GridColorFixed  =   16776960
      GridLinesUnpopulated=   1
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      Picture         =   "Ventas.frx":18E9
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtVuelto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6840
      Width           =   2535
   End
   Begin VB.TextBox txtCredito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   6000
      Width           =   2535
   End
   Begin VB.TextBox txtDebito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox txtEfectivo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Picture         =   "Ventas.frx":22EB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Mozo"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   8520
      TabIndex        =   33
      Top             =   7920
      Width           =   720
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Personas"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   7080
      TabIndex        =   32
      Top             =   7920
      Width           =   1320
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Mesa"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   5760
      TabIndex        =   31
      Top             =   7920
      Width           =   720
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H00FFFFFF&
      Height          =   1095
      Left            =   5520
      Shape           =   4  'Rounded Rectangle
      Top             =   7800
      Width           =   5895
   End
   Begin VB.Label lblCondicion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONTADO"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   600
      TabIndex        =   29
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   2360
      Width           =   5055
   End
   Begin VB.Label lblNumeroDocumento 
      BackStyle       =   0  'Transparent
      Caption         =   "lblNumeroDocumento"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1920
      TabIndex        =   27
      Top             =   1800
      Width           =   2130
   End
   Begin VB.Label lblTipoDocumento 
      BackStyle       =   0  'Transparent
      Caption         =   "lblTipoDocumento"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   600
      TabIndex        =   26
      Top             =   1800
      Width           =   1170
   End
   Begin VB.Label lblCategoria 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCategoria"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   600
      TabIndex        =   25
      Top             =   1440
      Width           =   3810
   End
   Begin VB.Label lblIdCliente 
      Caption         =   "Label10"
      Height          =   255
      Left            =   600
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCliente"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   600
      TabIndex        =   22
      Top             =   1080
      Width           =   3810
   End
   Begin FiscalPrinterLibCtl.HASAR HASAR1 
      Left            =   4560
      OleObjectBlob   =   "Ventas.frx":2BB5
      Top             =   0
   End
   Begin VB.Label lblCaja 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cerrada"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1920
      TabIndex        =   20
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Caja Nº:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   600
      TabIndex        =   19
      Top             =   480
      Width           =   1200
   End
   Begin VB.Label lblCajero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cajero:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3600
      TabIndex        =   18
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Barras"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   7200
      TabIndex        =   17
      Top             =   45
      Width           =   2400
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   5595
      TabIndex        =   16
      Top             =   45
      Width           =   1200
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   975
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00FFFFFF&
      Height          =   1095
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   7800
      Width           =   5055
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   3855
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vuelto:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   720
      TabIndex        =   7
      Top             =   6840
      Width           =   1470
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   600
      TabIndex        =   6
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Débito:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   720
      TabIndex        =   5
      Top             =   5160
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Efectivo:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   450
      TabIndex        =   4
      Top             =   4320
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   960
      TabIndex        =   3
      Top             =   3480
      Width           =   1260
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   4575
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   5055
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1335
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   6735
      Left            =   5520
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   6015
   End
End
Attribute VB_Name = "frmFacturador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Total As Double
Dim Efectivo As Double
Dim Debito As Double
Dim Credito As Double
Dim Vuelto As Double

Dim Nombre As String
Dim NumeroDoc As String
Dim TipoDoc As TiposDeDocumento
Dim Comprobante As DocumentosFiscales
Dim Responsable As TiposDeResponsabilidades
Dim HayScanner As String
Dim PuestoFiscal As Integer
Dim ComprobanteFiscal As String
Dim Buffer As String
Dim x As Printer

Private Sub cmdAceptar_Click()

If txtMesa.Text = "" Then: MsgBox ("Definir mesa"): txtMesa.SetFocus: Exit Sub
If txtPersonas.Text = "" Then: MsgBox ("Definir personas"): txtPersonas.SetFocus: Exit Sub
If txtMozo.Text = "" Then: MsgBox ("Definir mozo"): txtMozo.SetFocus: Exit Sub

'MsgBox (HASAR1.IndicadorFiscal(2048))
'HASAR1.CapacidadRestante
'Capacidad = Round(100 * (1 - PFiscal.Respuesta(4) / PFiscal.Respuesta(3)), 2)
    ComprobanteFiscal = "NO" 'esta variable la pongo por defecto en NO para que siempre tome venta sin factura
    
    If Total = 0 Then MsgBox ("Comprobante sin movimientos"): Exit Sub
    If (Efectivo - Vuelto + Debito + Credito) <> Total Then
        'MsgBox ("El detalle de pago no coincide con el total")
        'Exit Sub
    End If
    Respuesta = MsgBox("¿Confirma el comprobante?", vbYesNo, "")
    If Respuesta = vbNo Then Exit Sub
    'Confirmo comprobante
    Dim rs As ADODB.Recordset
    If ComprobanteFiscal <> "NO" Then
        Pedido = 0
        'On Error GoTo impresora_apag
'Imprimir:
        Select Case lblCategoria
            Case "Consumidor Final"
                Letra = "B"
                If lblCliente = "CONSUMIDOR FINAL" And Total <= 1000 Then
                    'es un consumidor final y no supera los $1000. Sale ticket
                    Nombre = ""
                    NumeroDoc = ""
                    TipoDoc = TIPO_NINGUNO
                    Comprobante = TICKET_C
                    Responsable = CONSUMIDOR_FINAL
                Else
                    'es un cliente seleccionado consumidor final o un consumidor final pero con importe mayor a $1000
                    Nombre = Mid(lblCliente, 1, 40)
                    NumeroDoc = lblNumeroDocumento
                    TipoDoc = TIPO_NINGUNO
                    Comprobante = TICKET_FACTURA_B
                    Responsable = CONSUMIDOR_FINAL
                End If
            Case "Monotributo"
                'es un cliente seleccionado consumidor final o un consumidor final pero con importe mayor a $1000
                Nombre = Mid(lblCliente, 1, 40)
                NumeroDoc = lblNumeroDocumento
                TipoDoc = TIPO_NINGUNO
                Comprobante = TICKET_FACTURA_B
                Responsable = CONSUMIDOR_FINAL
                Letra = "B"
            Case "Responsable Inscripto"
                Nombre = Mid(lblCliente, 1, 40)
                NumeroDoc = lblNumeroDocumento
                TipoDoc = TIPO_CUIT
                Comprobante = TICKET_FACTURA_A
                Responsable = RESPONSABLE_INSCRIPTO
                Letra = "A"
        End Select
        HASAR1.Encabezado(1) = Chr(244) & "     S  U  M  A"
        HASAR1.Encabezado(2) = "     R.S. Peña 245 - Junìn (Bs. As.)"
        HASAR1.Encabezado(3) = "          Te: (0236) 4443018        "
        
        HASAR1.DatosCliente Nombre, NumeroDoc, TipoDoc, Responsable

        'HayError = PFiscal.HuboErrorFiscal Or PFiscal.HuboErrorMecanico Or PFiscal.HuboFaltaPapel
        'If HayError Then MsgBox "Los datos del cliente son incorrectos o el cuit es inválido": Exit Sub
        
        
        HASAR1.AbrirComprobanteFiscal Comprobante
        'HayError = PFiscal.HuboErrorFiscal Or PFiscal.HuboErrorMecanico Or PFiscal.HuboFaltaPapel
        'If HayError Then MsgBox "No se puede abrir el comprobante fiscal. Se realizó cierre Z?": Exit Sub
    
        'HASAR1.ImprimirTextoFiscal "Texto Fiscal..."
        With grDetalle
        For I = 0 To grDetalle.Rows - 1
            HASAR1.ImprimirItem Mid(.TextMatrix(I, 1), 1, 20), .TextMatrix(I, 0), Val(.TextMatrix(I, 2)) / Val(.TextMatrix(I, 0)), 21, 0
        Next I
        End With
        'HASAR1.ImprimirPago "Efectivo", Val(Total)
        HASAR1.CerrarComprobanteFiscal
        If lblCategoria = "Responsable Inscripto" Then
            NumeroComprobante = HASAR1.UltimoDocumentoFiscalA
        Else
            NumeroComprobante = HASAR1.UltimoDocumentoFiscalBC
        End If
    Else
        cn.Open
        Set rs = cn.Execute("SELECT MAX(numero) + 1 AS NuevoP FROM VENTAS WHERE pedido=1")
        If rs!NuevoP > 0 Then
            NumeroComprobante = rs!NuevoP
        Else
            NumeroComprobante = 1
        End If
        rs.Close
        Set rs = Nothing
        cn.Close
        Letra = "P"
        PuestoFiscal = 0
        Pedido = 1
    End If
    
    Fecha = DTPFecha.Value
    
    
    cn.Open
    cn.Execute ("AgregarVenta '" & Format(Fecha, "yyyy/mm/dd") & "','" & Letra & "'," & PuestoFiscal & "," & NumeroComprobante & "," & Replace((Total / 1.21), ",", ".") & "," & Replace(((Total / 1.21) * 0.21), ",", ".") & "," & Replace(Total, ",", ".") & "," & Val(lblIdCliente) & "," & Pedido)
    Set rs = cn.Execute("SELECT MAX(idventa) AS Nuevoid FROM VENTAS")
    NuevoID = rs!NuevoID
    idComprobante = NuevoID
    cn.Execute ("UPDATE Ventas set mesa=" & txtMesa & ", personas=" & txtPersonas & ", mozo='" & txtMozo & "' WHERE idVenta=" & idComprobante)



    With grDetalle
    For I = 0 To grDetalle.Rows - 1
        cn.Execute ("AgregaDetalleVenta " & Replace((Val(.TextMatrix(I, 2)) / 1.21) / Val(.TextMatrix(I, 0)), ",", ".") & "," & Replace(.TextMatrix(I, 2), ",", ".") & " ," & Replace(.TextMatrix(I, 0), ",", ".") & "," & Replace(.TextMatrix(I, 4), ",", ".") & "," & NuevoID & "," & .TextMatrix(I, 3))
        cn.Execute ("DescuentaStock " & .TextMatrix(I, 3) & "," & .TextMatrix(I, 0))
    Next I
    End With
    

    If lblCondicion = "CUENTA CORRIENTE" Then
        cn.Execute ("AgregarCuentaCorriente '" & Format(Fecha, "yyyy/mm/dd") & "','Venta " & Letra & Format(PuestoFiscal, "0000") & "-" & Format(NumeroComprobante, "00000000") & "'," & Replace(Total, ",", ".") & "," & "0," & Val(lblIdCliente) & ",'Ven'," & NuevoID)
    Else
        cn.Execute ("AgregarDetalleCaja " & lblCaja & ",'" & Format(Fecha, "yyyy/mm/dd") & "'," & Replace(Efectivo - Vuelto, ",", ".") & "," & Replace(Debito, ",", ".") & "," & Replace(Credito, ",", ".") & ",1,'Factura " & txtnumero & "'")
    End If
    cn.Close
    If ComprobanteFiscal = "NO" Then
        cn.Open
        
        Dim obj_Impresora As Object
        Set obj_Impresora = CreateObject("WScript.Network")
        obj_Impresora.setdefaultprinter printerTicket

        'rptComprobanteNoFiscal.WindowState = 2
        'rptComprobanteNoFiscal.Show 1
         rptComprobanteNoFiscal.PrintReport
         
        Unload rptComprobanteNoFiscal
        cn.Close
        obj_Impresora.setdefaultprinter printerDefault
    End If
    

    grDetalle.Rows = 0
    Total = 0
    txtTotal = ""
    txtEfectivo = ""
    txtDebito = ""
    txtCredito = ""
    txtVuelto = ""
    lblIdCliente = 1
    lblCliente = "CONSUMIDOR FINAL"
    lblCategoria = "Consumidor Final"
    lblTipoDocumento = ""
    lblNumeroDocumento = ""
    txtBarras.SetFocus
    Exit Sub
    
'impresora_apag:
'    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
'        Resume Imprimir
'    End If
End Sub
Sub EntregarMercaderia()
    Fecha = DTPFecha.Value
    If (txtEfectivo - txtVuelto + txtDebito + txtCredito) <> Total Then
        MsgBox ("El detalle de pago no coincide con el total")
        Exit Sub
    End If
    Respuesta = MsgBox("¿Confirma la entrega de mercadería?", vbYesNo, "")
    If Respuesta = vbNo Then Exit Sub
    'Confirmo comprobante
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT MAX(numero) + 1 AS NuevoP FROM VENTAS WHERE pedido=1")
    If rs!NuevoP > 0 Then
        NumeroComprobante = rs!NuevoP
    Else
        NumeroComprobante = 1
    End If
    rs.Close
    Set rs = Nothing
    cn.Execute ("AgregarVenta '" & Format(Fecha, "mm/dd/yyyy") & "','P',0," & NumeroComprobante & "," & Replace((Total / 1.21), ",", ".") & "," & Replace(((txtTotal / 1.21) * 0.21), ",", ".") & "," & Replace(Total, ",", ".") & "," & Val(lblIdCliente) & ",1")
    Set rs = cn.Execute("SELECT MAX(idventa) AS Nuevoid FROM VENTAS")
    NuevoID = rs!NuevoID
    With grDetalle
    For I = 0 To grDetalle.Rows - 1
        cn.Execute ("AgregaDetalleVenta " & Replace((Val(.TextMatrix(I, 2)) / 1.21), ",", ".") / Replace(Val(.TextMatrix(I, 0)), ",", ".") & "," & Replace(.TextMatrix(I, 0), ",", ".") & "," & Replace(.TextMatrix(I, 4), ",", ".") & "," & NuevoID & "," & .TextMatrix(I, 3))
        cn.Execute ("DescuentaStock " & .TextMatrix(I, 3) & "," & .TextMatrix(I, 0))
    Next I
    End With
    
    cn.Execute ("AgregarDetalleCaja " & lblCaja & ",'" & Format(Fecha, "mm/dd/yyyy") & "'," & Replace(txtEfectivo - txtVuelto, ",", ".") & "," & Replace(txtDebito, ",", ".") & "," & Replace(Credito, ",", ".") & ",1,'Factura " & txtnumero & "'")
    cn.Close
    grDetalle.Rows = 0
    Total = 0
    txtTotal = ""
    txtEfectivo = ""
    txtDebito = ""
    txtCredito = ""
    txtVuelto = ""
    lblIdCliente = 1
    lblCliente = "CONSUMIDOR FINAL"
    lblCategoria = "Consumidor Final"
    lblTipoDocumento = ""
    lblNumeroDocumento = ""

    txtBarras.SetFocus
    
    Exit Sub
End Sub

Private Sub cmdBuscar_Click()
    frmBuscaArticulos.Show 1
    If Len(txtBarras.Text) > 0 Then CargarDetalle
    txtCantidad.SetFocus
    txtCantidad.SelStart = 0
    txtCantidad.SelLength = Len(txtCantidad.Text)
End Sub

Private Sub cmdBuscarCliente_Click()
    EligiendoCliente = 1
    frmClientes.Show 1
    EligiendoCliente = 0
    txtCantidad.SetFocus
End Sub

Private Sub cmdCaja_Click()
    CerroCaja = 0
    frmCajaCierre.Show 1
    If CerroCaja = 1 Then Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
MsgBox (Efectivo & " - " & Debito & " - " & Credito & " - " & Vuelto)
End Sub

Private Sub Form_Activate()
    If PidiendoPrecio = 0 Then ControlarCaja
    'cn.Close
End Sub

Private Sub Form_Load()
    grDetalle.Cols = 5
    grDetalle.ColWidth(0) = 600
    grDetalle.ColWidth(1) = 3600
    grDetalle.ColWidth(3) = 0
    grDetalle.ColWidth(4) = 0
    txtCantidad = 1
    lblCajero = Usuario
    lblIdCliente = 1
    lblCliente = "CONSUMIDOR FINAL"
    lblCategoria = "Consumidor Final"
    lblTipoDocumento = ""
    lblNumeroDocumento = ""
    ControloScanner
    DTPFecha.Value = Date
    

On Error GoTo impresora_apag
Procesar:

    HASAR1.Puerto = portfiscal
    HASAR1.Modelo = MODELO_PR4
    HASAR1.Comenzar
    HASAR1.PrecioBase = False
    HASAR1.TratarDeCancelarTodo
    HASAR1.ObtenerDatosDeInicializacion cuit, razon, serie, fechainicio, Puesto, fechainicio, codiibb, categoria
    PuestoFiscal = Puesto
    Exit Sub

impresora_apag:

    'If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
    '    Resume Procesar
    'End If
    
End Sub
Sub ControloScanner()
On Error GoTo NoScanner
    MSComm1.CommPort = portscan
    MSComm1.PortOpen = True
    Exit Sub
NoScanner:
    'MsgBox ("No se localizó ningún scanner")
    'HayScanner = "no"
    'Resume Next
End Sub
Sub ControlarCaja()
    cn.Open
    Dim rs As New ADODB.Recordset
    Set rs = cn.Execute("BuscarCajaAbierta " & idUsuario)
    If rs.EOF = True Then 'no existe ninguna caja abierta para este usuario
        MsgBox ("Debe abrir una caja antes de operar"): cn.Close: Unload Me: Exit Sub
        Set rs = Nothing
        Set rs = cn.Execute("VerUltimaCaja " & idUsuario)

        'If cn.State = 0 Then
        '    cn.Open
        'End If
        

        If rs.EOF = True Then 'este usuario nunca tuvo una caja
            frmCajaApertura.txtEfectivo = 0
            frmCajaApertura.txtDebito = 0
            frmCajaApertura.txtCredito = 0
            frmCajaApertura.txtSaldoApertura = 0
        Else
            frmCajaApertura.txtEfectivo = Format(rs!EfectivoFinal, "0.00")
            frmCajaApertura.txtDebito = Format(rs!DebitoFinal, "0.00")
            frmCajaApertura.txtCredito = Format(rs!CreditoFinal, "0.00")
            frmCajaApertura.txtSaldoApertura = Format(rs!EfectivoFinal + rs!DebitoFinal + rs!CreditoFinal, "0.00")
        End If
        frmCajaApertura.lblUsuario = Usuario
        frmCajaApertura.Show 1
    Else
        lblCaja = rs!idCaja
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
    If Val(lblCaja) = 0 Then Unload Me: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HayScanner <> "no" Then
        'MSComm1.PortOpen = False
    End If
    HASAR1.Finalizar
    Total = 0
End Sub

Private Sub grDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        Respuesta = MsgBox("¿Está seguro de borrar el artículo?", vbYesNo, "Borrar")
        If Respuesta = vbNo Then Exit Sub
        Total = Total - grDetalle.TextMatrix(grDetalle.Row, 2)
        txtTotal = Total
        txtEfectivo = Total
        If grDetalle.Rows > 1 Then
            grDetalle.RemoveItem (grDetalle.Row)
        Else
            grDetalle.Rows = 0
            Total = 0
            txtTotal = Total
            txtEfectivo = Total
        End If
    End If
    ControlTeclas (KeyCode)
End Sub

Private Sub lblCondicion_Click()
    If lblCondicion = "CONTADO" Then
        lblCondicion = "CUENTA CORRIENTE"
        lblCondicion.ForeColor = vbBlue
    Else
        lblCondicion = "CONTADO"
        lblCondicion.ForeColor = vbGreen
    End If
End Sub

Private Sub MSComm1_OnComm()
    If MSComm1.CommEvent = comEvReceive And MSComm1.InBufferCount > 0 Then
        Buffer = Buffer & MSComm1.Input
        If Asc(Mid(Buffer, Len(Buffer), 1)) = 10 Then
            txtBarras = Mid(Buffer, 1, Len(Buffer) - 2)
            Buffer = ""
            CargarDetalle
        End If
    End If
End Sub

Private Sub txtBarras_LostFocus()
    If txtBarras.Text <> "" Then CargarDetalle
End Sub


Private Sub txtBarras_GotFocus()
    txtBarras = ""
End Sub

Private Sub txtBarras_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtBarras_KeyPress(KeyAscii As Integer)

    If txtBarras.Text = "" Then Exit Sub
    'If InStr(1, "0123456789" & Chr(9), Chr(KeyAscii)) = 0 Then
    '    KeyAscii = 0
    'End If
    'If KeyAscii = 13 Then
    '    CargarDetalle
    'End If
End Sub
Sub CargarDetalle()
    If IsNumeric(txtCantidad) = False Then MsgBox ("La cantidad no es valida"): txtCantidad.SetFocus: Exit Sub
    cn.Open
    Dim rs As ADODB.Recordset
    Barra = txtBarras.Text
    Set rs = cn.Execute("SELECT idArticulo,Descripcion ,CodBar,Venta, Costo FROM Articulos where CodBar='" & Barra & "'")
    If rs.EOF = True Then
        encontro = "no"
       'busco quitando el ultimo digito
       Barra = Mid(RTrim(txtBarras), 1, Len(RTrim(txtBarras)) - 1)
       Set rs = cn.Execute("SELECT idArticulo,Descripcion ,CodBar,Venta, Costo FROM Articulos where CodBar='" & Barra & "'")
       If rs.EOF = False Then
            encontro = "si"
       End If
    Else
        encontro = "si"
    End If

    If encontro = "si" Then
       grDetalle.Rows = grDetalle.Rows + 1
       grDetalle.TextMatrix(grDetalle.Rows - 1, 0) = txtCantidad
       grDetalle.TextMatrix(grDetalle.Rows - 1, 1) = rs!Descripcion
       grDetalle.TextMatrix(grDetalle.Rows - 1, 2) = Format(rs!Venta * txtCantidad, "0.00")
       If rs!Venta = 0 Or rs!Descripcion = "FOTOCOPIAS" Or rs!Descripcion = "LIBRERIA" Or rs!Descripcion = "VARIOS" Then
            frmPedirPrecio.lblDescripcion = rs!Descripcion
            PidiendoPrecio = 1
            frmPedirPrecio.Show 1
            grDetalle.TextMatrix(grDetalle.Rows - 1, 2) = Format(Val(frmPedirPrecio.txtPrecio) * txtCantidad, "0.00")
            cn.Execute ("ActualizarPrecioVenta " & rs!idArticulo & "," & Replace(Val(frmPedirPrecio.txtPrecio), ",", "."))
       End If
       grDetalle.TextMatrix(grDetalle.Rows - 1, 3) = rs!idArticulo
       grDetalle.TextMatrix(grDetalle.Rows - 1, 4) = Format(rs!Costo * txtCantidad, "0.00")
       Total = (Total + grDetalle.TextMatrix(grDetalle.Rows - 1, 2))
       Total = Total
       Efectivo = Total
       Debito = 0
       Credito = 0
       PasarTotalesAtxt
    Else
       If HayScanner <> "no" Then
            'MSComm1.PortOpen = False
       End If
       MsgBox ("No se encontro el articulo")
       If HayScanner <> "no" Then
            'MSComm1.PortOpen = True
       End If
       txtBarras = "": txtBarras.SetFocus
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
    txtBarras = ""
    txtCantidad = 1
    PidiendoPrecio = 0
    txtBarras.SetFocus
End Sub



Private Sub txtCantidad_GotFocus()
    txtCantidad.SelStart = 0
    txtCantidad.SelLength = Len(txtCantidad.Text)
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBarras.SetFocus
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCredito_Change()
    Credito = Val(txtCredito)
End Sub

Private Sub txtCredito_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtDebito_Change()
    Debito = Val(txtDebito)
End Sub

Private Sub txtDebito_GotFocus()
    'txtDebito = txtTotal
    txtDebito.SelStart = 0
    txtDebito.SelLength = Len(txtDebito.Text)
End Sub

Private Sub txtDebito_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtDebito_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtCredito_GotFocus()
    'txtCredito = txtTotal
    txtCredito.SelStart = 0
    txtCredito.SelLength = Len(txtCredito.Text)
End Sub

Private Sub txtCredito_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEfectivo_Change()
    If Val(txtEfectivo) + Debito + Credito > Total Then
        Efectivo = Val(txtEfectivo)
        Vuelto = Efectivo + Debito + Credito - Total
        txtVuelto = Format(Vuelto, "0.00")
    Else
        txtVuelto = "0.00"
        Vuelto = 0
    End If
    Efectivo = Val(txtEfectivo)
    End Sub

Private Sub txtEfectivo_GotFocus()
    'txtEfectivo = Total
    txtEfectivo.SelStart = 0
    txtEfectivo.SelLength = Len(txtEfectivo.Text)
End Sub

Private Sub txtEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtEfectivo_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Sub ControlTeclas(Tecla As Integer)
    Select Case Tecla
        Case 112
            cmdBuscar_Click
        Case 113
            cmdAceptar_Click
        Case 114
            cmdCaja_Click
        Case 115
            cmdBuscarCliente_Click
        Case 123
            ComprobanteFiscal = "NO"
            cmdAceptar_Click
            ComprobanteFiscal = "SI"
        Case 27
            Respuesta = MsgBox("¿Cancela el comprobante?", vbYesNo, "")
            If Respuesta = vbYes Then cmdSalir_Click
    End Select
End Sub

Private Sub txtTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtVuelto_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub PasarTotalesAtxt()
    txtTotal = Total
    txtEfectivo = Efectivo
    txtDebito = Debito
    txtCredito = Credito
    txtVuelto = Vuelto
End Sub
