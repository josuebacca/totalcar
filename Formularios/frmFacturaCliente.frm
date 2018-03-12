VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFacturaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura de Clientes..."
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   8115
      TabIndex        =   15
      Top             =   7530
      Width           =   990
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10185
      TabIndex        =   17
      Top             =   7530
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   7080
      TabIndex        =   14
      Top             =   7530
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   9150
      TabIndex        =   16
      Top             =   7530
      Width           =   990
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7500
      Left            =   60
      TabIndex        =   34
      Top             =   15
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   13229
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   512
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "frmFacturaCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameFactura"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameRemito"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "freCliente"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmFacturaCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "frameBuscar"
      Tab(1).ControlCount=   2
      Begin VB.Frame freCliente 
         Height          =   1775
         Left            =   4320
         TabIndex        =   74
         Top             =   900
         Width           =   6705
         Begin VB.TextBox txtcodpos 
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   930
            TabIndex        =   96
            Top             =   780
            Width           =   1215
         End
         Begin VB.TextBox txtprovincia 
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   930
            MaxLength       =   50
            TabIndex        =   84
            Top             =   1080
            Width           =   4380
         End
         Begin VB.TextBox txtlocalidad 
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   2250
            MaxLength       =   50
            TabIndex        =   83
            Top             =   780
            Width           =   4380
         End
         Begin VB.TextBox txtCUIT 
            BackColor       =   &H8000000F&
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
            Height          =   300
            Left            =   930
            TabIndex        =   82
            Top             =   1395
            Width           =   1455
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2385
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaCliente.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Agregar Cliente"
            Top             =   133
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaCliente.frx":03C2
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Buscar Cliente"
            Top             =   133
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtCondicionIVA 
            BackColor       =   &H8000000F&
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
            Height          =   300
            Left            =   2415
            TabIndex        =   79
            Top             =   1395
            Width           =   2895
         End
         Begin VB.TextBox TxtCodigoCli 
            Enabled         =   0   'False
            Height          =   300
            Left            =   930
            MaxLength       =   40
            TabIndex        =   78
            Top             =   140
            Width           =   975
         End
         Begin VB.TextBox txtRazSocCli 
            BackColor       =   &H8000000F&
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
            Height          =   300
            Left            =   2835
            MaxLength       =   50
            TabIndex        =   77
            Tag             =   "Descripción"
            Top             =   140
            Width           =   3750
         End
         Begin VB.TextBox txtDomici 
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   930
            MaxLength       =   50
            TabIndex        =   76
            Top             =   465
            Width           =   4620
         End
         Begin VB.TextBox txtIngBrutos 
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   5370
            TabIndex        =   75
            Top             =   1395
            Width           =   1215
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   150
            TabIndex        =   90
            Top             =   1125
            Width           =   705
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   795
            Width           =   735
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   330
            TabIndex        =   88
            Top             =   165
            Width           =   525
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   180
            TabIndex        =   87
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   255
            TabIndex        =   86
            Top             =   1440
            Width           =   600
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos"
            Height          =   195
            Left            =   5730
            TabIndex        =   85
            Top             =   1200
            Width           =   765
         End
      End
      Begin VB.Frame FrameRemito 
         Caption         =   "Remito ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   4320
         TabIndex        =   50
         Top             =   300
         Width           =   6705
         Begin VB.TextBox txtNroRemito 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1515
            TabIndex        =   99
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdBuscarRemito 
            Height          =   315
            Left            =   2640
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaCliente.frx":06CC
            Style           =   1  'Graphical
            TabIndex        =   98
            ToolTipText     =   "Buscar Remito"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtRemSuc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   930
            MaxLength       =   4
            TabIndex        =   5
            Top             =   240
            Width           =   555
         End
         Begin VB.TextBox txtCodigoStock 
            Height          =   300
            Left            =   5760
            TabIndex        =   72
            Top             =   240
            Visible         =   0   'False
            Width           =   465
         End
         Begin MSComCtl2.DTPicker FechaRemito 
            Height          =   315
            Left            =   4080
            TabIndex        =   6
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   54525953
            CurrentDate     =   41098
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   3555
            TabIndex        =   52
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   300
            Width           =   600
         End
      End
      Begin VB.Frame frameBuscar 
         Caption         =   "Buscar por..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   -74595
         TabIndex        =   39
         Top             =   480
         Width           =   10410
         Begin VB.OptionButton optRep 
            Caption         =   "de Repuestos"
            Height          =   195
            Left            =   1560
            TabIndex        =   30
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton optMaq 
            Caption         =   "de Maquinarias"
            Height          =   195
            Left            =   3360
            TabIndex        =   31
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton optTodas 
            Caption         =   "Todas"
            Height          =   195
            Left            =   5280
            TabIndex        =   32
            Top             =   1800
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.CheckBox chkProducto 
            Caption         =   "Producto"
            Height          =   195
            Left            =   210
            TabIndex        =   103
            Top             =   660
            Width           =   1035
         End
         Begin VB.TextBox txtDesProducto 
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   4875
            TabIndex        =   102
            Top             =   615
            Width           =   4620
         End
         Begin VB.TextBox txtProducto 
            Height          =   285
            Left            =   3360
            TabIndex        =   22
            Top             =   600
            Width           =   990
         End
         Begin VB.CommandButton cmdProducto 
            Height          =   315
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaCliente.frx":09D6
            Style           =   1  'Graphical
            TabIndex        =   101
            ToolTipText     =   "Buscar Vendedor"
            Top             =   615
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CheckBox chkanuladas 
            Caption         =   "Ver Anuladas"
            Height          =   255
            Left            =   6600
            TabIndex        =   26
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton cmdBuscarVen 
            Height          =   300
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaCliente.frx":0CE0
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Buscar Vendedor"
            Top             =   840
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.ComboBox cboFactura1 
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1365
            Width           =   2400
         End
         Begin VB.CheckBox chkTipoFactura 
            Caption         =   "Tipo de Facrura"
            Height          =   195
            Left            =   210
            TabIndex        =   27
            Top             =   1200
            Width           =   1485
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaCliente.frx":0FEA
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Buscar Cliente"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3360
            TabIndex        =   20
            Top             =   825
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.TextBox txtDesVen 
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   4845
            TabIndex        =   44
            Top             =   840
            Visible         =   0   'False
            Width           =   4860
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   450
            TabIndex        =   21
            Top             =   885
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1815
            Left            =   9690
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaCliente.frx":12F4
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Buscar "
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   570
         End
         Begin VB.TextBox txtDesCli 
            BackColor       =   &H8000000F&
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
            Height          =   300
            Left            =   4875
            MaxLength       =   50
            TabIndex        =   40
            Tag             =   "Descripción"
            Top             =   255
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3360
            MaxLength       =   40
            TabIndex        =   19
            Top             =   255
            Width           =   975
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   210
            TabIndex        =   25
            Top             =   930
            Width           =   810
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   210
            TabIndex        =   18
            Top             =   345
            Width           =   855
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   3360
            TabIndex        =   23
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   54525953
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   6360
            TabIndex        =   24
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   54525953
            CurrentDate     =   41098
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Producto:"
            Height          =   195
            Left            =   2625
            TabIndex        =   104
            Top             =   630
            Width           =   690
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Factura:"
            Height          =   195
            Left            =   2325
            TabIndex        =   70
            Top             =   1410
            Width           =   945
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   5295
            TabIndex        =   43
            Top             =   975
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2265
            TabIndex        =   42
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   2790
            TabIndex        =   41
            Top             =   300
            Width           =   525
         End
      End
      Begin VB.Frame FrameFactura 
         Caption         =   "Factura..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1890
         Left            =   120
         TabIndex        =   36
         Top             =   300
         Width           =   4140
         Begin VB.CommandButton cmdNuevoVendedor 
            Height          =   315
            Left            =   2700
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaCliente.frx":3A96
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Agregar Vendedor"
            Top             =   1200
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2280
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaCliente.frx":3E20
            Style           =   1  'Graphical
            TabIndex        =   92
            ToolTipText     =   "Buscar Vendedor"
            Top             =   1200
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtNombreVendedor 
            BackColor       =   &H8000000F&
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
            Height          =   315
            Left            =   480
            TabIndex        =   91
            Top             =   1545
            Width           =   3165
         End
         Begin VB.TextBox txtNroVendedor 
            Height          =   315
            Left            =   1380
            TabIndex        =   4
            Top             =   1200
            Width           =   780
         End
         Begin VB.TextBox txtNroSucursal 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            MaxLength       =   4
            TabIndex        =   1
            Top             =   585
            Width           =   435
         End
         Begin VB.ComboBox cboFactura 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   2190
         End
         Begin VB.TextBox txtNroFactura 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1170
            MaxLength       =   8
            TabIndex        =   2
            Top             =   585
            Width           =   945
         End
         Begin MSComCtl2.DTPicker FechaFactura 
            Height          =   315
            Left            =   2760
            TabIndex        =   3
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   54525953
            CurrentDate     =   41098
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   540
            TabIndex        =   94
            Top             =   1245
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   300
            TabIndex        =   56
            Top             =   255
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   2205
            TabIndex        =   53
            Top             =   615
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   60
            TabIndex        =   49
            Top             =   615
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   930
            Width           =   540
         End
         Begin VB.Label lblEstadoFactura 
            AutoSize        =   -1  'True
            Caption         =   "EST. FACTURA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   720
            TabIndex        =   47
            Top             =   945
            Width           =   1350
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4620
         Left            =   -74640
         TabIndex        =   33
         Top             =   2700
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8149
         _Version        =   393216
         Cols            =   17
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame4 
         Height          =   540
         Left            =   105
         TabIndex        =   54
         Top             =   2115
         Width           =   4150
         Begin VB.CommandButton cmdNuevoRubro 
            Height          =   315
            Left            =   3795
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaCliente.frx":412A
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Agregar Condición de Venta"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.ComboBox cboCondicion 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   165
            Width           =   2910
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Condición:"
            Height          =   195
            Left            =   30
            TabIndex        =   67
            Top             =   210
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4830
         Left            =   105
         TabIndex        =   37
         Top             =   2595
         Width           =   10935
         Begin VB.TextBox txtAuxNeto 
            Height          =   285
            Left            =   5040
            TabIndex        =   106
            Top             =   3480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtAuxIva 
            Height          =   285
            Left            =   4440
            TabIndex        =   105
            Top             =   3480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmdredondeo 
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10440
            TabIndex        =   97
            Top             =   3960
            Width           =   255
         End
         Begin VB.CheckBox chkBonificaEnPesos 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en $"
            Height          =   285
            Left            =   390
            TabIndex        =   10
            Top             =   3840
            Width           =   1290
         End
         Begin VB.CheckBox chkBonificaEnPorsentaje 
            Alignment       =   1  'Right Justify
            Caption         =   "Descuento % "
            Height          =   285
            Left            =   390
            TabIndex        =   9
            Top             =   3540
            Width           =   1290
         End
         Begin VB.TextBox txtSubTotalBoni 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   4905
            TabIndex        =   68
            Top             =   3870
            Width           =   1155
         End
         Begin VB.TextBox txtImporteIva 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   6900
            TabIndex        =   64
            Top             =   3870
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeIva 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6900
            TabIndex        =   12
            Top             =   3540
            Width           =   1155
         End
         Begin VB.TextBox txtImporteBoni 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   2850
            TabIndex        =   61
            Top             =   3870
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeBoni 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2850
            TabIndex        =   11
            Top             =   3540
            Width           =   1155
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   8970
            TabIndex        =   58
            Top             =   3870
            Width           =   1350
         End
         Begin VB.TextBox txtSubtotal 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   8970
            TabIndex        =   57
            Top             =   3540
            Width           =   1350
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   540
            Left            =   1455
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   4215
            Width           =   8865
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   1140
            TabIndex        =   38
            Top             =   420
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3375
            Left            =   75
            TabIndex        =   8
            Top             =   120
            Width           =   10725
            _ExtentX        =   18918
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   3
            Cols            =   11
            FixedCols       =   0
            BackColorSel    =   12648447
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            AllowUserResizing=   3
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   4110
            TabIndex        =   69
            Top             =   3930
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   6270
            TabIndex        =   66
            Top             =   3915
            Width           =   570
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "% I.V.A.:"
            Height          =   195
            Left            =   6240
            TabIndex        =   65
            Top             =   3570
            Width           =   600
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   2235
            TabIndex        =   63
            Top             =   3915
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Descuento:"
            Height          =   195
            Left            =   1980
            TabIndex        =   62
            Top             =   3570
            Width           =   825
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   8505
            TabIndex        =   60
            Top             =   3915
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   8175
            TabIndex        =   59
            Top             =   3570
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   210
            TabIndex        =   55
            Top             =   4500
            Width           =   1110
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   35
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vendedor:"
      Height          =   195
      Index           =   2
      Left            =   3045
      TabIndex        =   100
      Top             =   750
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3240
      TabIndex        =   95
      Top             =   7680
      Width           =   2130
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   225
      TabIndex        =   46
      Top             =   7635
      Width           =   750
   End
End
Attribute VB_Name = "frmFacturaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim I As Integer
Dim w As Integer
Dim TipoBusquedaDoc As Integer
Dim VBonificacion As Double
Dim VTotal As Double
Dim VEstadoFactura As Integer
Dim Rec1 As ADODB.Recordset

Private Sub cboFactura_LostFocus()
    Dim I As Integer
    txtNroFactura.Text = Format(BuscoUltimaFactura(cboFactura.ItemData(cboFactura.ListIndex)), "00000000")
    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then ' Factura B
        txtAuxIva = txtPorcentajeIva.Text
        txtAuxNeto = txtSubtotal.Text
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) = "" Then
                Exit For
            End If
            grdGrilla.TextMatrix(I, 3) = CDbl(grdGrilla.TextMatrix(I, 3)) + (CDbl(grdGrilla.TextMatrix(I, 3)) * CDbl(txtPorcentajeIva.Text)) / 100
            grdGrilla.TextMatrix(I, 6) = CDbl(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3))
            grdGrilla.TextMatrix(I, 3) = Format(grdGrilla.TextMatrix(I, 3), "#0.00")
            grdGrilla.TextMatrix(I, 6) = Format(grdGrilla.TextMatrix(I, 6), "#0.00")
        Next I
        txtPorcentajeIva.Text = "0,00"
        txtPorcentajeIva_LostFocus
        txtSubtotal.Text = Valido_Importe(SumaBonificacion)
        txtTotal.Text = txtSubtotal.Text
    End If
End Sub

Private Sub chkBonificaEnPesos_Click()
    If chkBonificaEnPesos.Value = Checked Then
        chkBonificaEnPorsentaje.Value = Unchecked
        chkBonificaEnPorsentaje.Enabled = False
    Else
        chkBonificaEnPorsentaje.Enabled = True
    End If
    txtPorcentajeBoni.Text = ""
    txtImporteBoni.Text = ""
    txtSubTotalBoni.Text = ""
End Sub

Private Sub chkBonificaEnPorsentaje_Click()
    If chkBonificaEnPorsentaje.Value = Checked Then
        chkBonificaEnPesos.Value = Unchecked
        chkBonificaEnPesos.Enabled = False
    Else
        chkBonificaEnPesos.Enabled = True
    End If
    txtPorcentajeBoni.Text = ""
    txtImporteBoni.Text = ""
    txtSubTotalBoni.Text = ""
End Sub

Private Sub chkCliente_Click()
    If chkCliente.Value = Checked Then
        txtCliente.Enabled = True
        cmdBuscarCli.Enabled = True
    Else
        txtCliente.Enabled = False
        cmdBuscarCli.Enabled = False
    End If
End Sub

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        FechaDesde.Enabled = True
        FechaHasta.Enabled = True
    Else
        FechaDesde.Enabled = False
        FechaHasta.Enabled = False
    End If
End Sub

Private Sub chkProducto_Click()
 If chkProducto.Value = Checked Then
        txtproducto.Enabled = True
        cmdProducto.Enabled = True
    Else
        txtproducto.Enabled = False
        cmdProducto.Enabled = False
        txtproducto.Text = ""
        txtdesProducto.Text = ""
    End If
End Sub

Private Sub chkTipoFactura_Click()
    If chkTipoFactura.Value = Checked Then
        cboFactura1.Enabled = True
    Else
        cboFactura1.Enabled = False
    End If
End Sub

Private Sub chkTipoFactura_LostFocus()
    If chkTipoFactura.Value = Checked And chkCliente.Value = Unchecked _
        And chkVendedor.Value = Unchecked _
        And chkFecha.Value = Unchecked Then cboFactura1.SetFocus
End Sub

Private Sub chkVendedor_Click()
    If chkVendedor.Value = Checked Then
        txtVendedor.Enabled = True
        cmdBuscarVen.Enabled = True
    Else
        txtVendedor.Enabled = False
        cmdBuscarVen.Enabled = False
    End If
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Select Case TipoBusquedaDoc
    
    Case 1 'BUSCA FACTURA
        chkVendedor.Enabled = True
        txtVendedor.Enabled = True
        cmdBuscarVen.Enabled = False
        
        sql = "SELECT DISTINCT FC.*,"
        sql = sql & " C.CLI_CODIGO,C.CLI_RAZSOC, C.CLI_DOMICI, TC.TCO_ABREVIA "
        If txtproducto.Text <> "" Then
            sql = sql & ",DT.PTO_CODIGO"
        End If
        sql = sql & " FROM FACTURA_CLIENTE FC, REMITO_CLIENTE RC, CLIENTE C,"
        sql = sql & " TIPO_COMPROBANTE TC"
        If txtproducto.Text <> "" Then
            sql = sql & ", DETALLE_FACTURA_CLIENTE DT"
        End If
        sql = sql & " WHERE"
        sql = sql & " FC.RCL_NUMERO=RC.RCL_NUMERO"
        sql = sql & " AND FC.RCL_SUCURSAL=RC.RCL_SUCURSAL"
        sql = sql & " AND FC.RCL_FECHA=RC.RCL_FECHA"
        sql = sql & " AND FC.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
        If txtproducto.Text <> "" Then
            sql = sql & " AND FC.TCO_CODIGO = DT.TCO_CODIGO"
            sql = sql & " AND FC.FCL_NUMERO = DT.FCL_NUMERO"
            sql = sql & " AND FC.FCL_SUCURSAL = DT.FCL_SUCURSAL"
            sql = sql & " AND FC.FCL_FECHA = DT.FCL_FECHA"
            'sql = sql & " AND DT.PTO_CODIGO=P.PTO_CODIGO"
        End If
        
        If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
        If txtVendedor.Text <> "" Then sql = sql & " AND FC.VEN_CODIGO=" & XN(txtVendedor)
        If Not IsNull(FechaDesde) Then sql = sql & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta)
        If chkTipoFactura.Value = Checked Then sql = sql & " AND FC.TCO_CODIGO=" & XN(cboFactura1.ItemData(cboFactura1.ListIndex))
        'CAMBIOS
        If txtproducto.Text <> "" Then sql = sql & " AND DT.PTO_CODIGO LIKE '" & txtproducto.Text & "'"
        
        If chkanuladas.Value = Unchecked Then
            sql = sql & " AND FC.EST_CODIGO <> 2"
        Else
            sql = sql & " AND FC.EST_CODIGO = 2"
        End If
        'esto lo uso para filtrar por facturas de maquinarias o repuestos
        If optMaq.Value = True Then
            sql = sql & " AND FC.FCL_IVA = 21 "
        End If
        If optRep.Value = True Then
            sql = sql & " AND FC.FCL_IVA = 21 "
        End If
        
        
        
        
        sql = sql & " ORDER BY FC.FCL_FECHA,FC.FCL_SUCURSAL,FC.FCL_NUMERO"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") _
                                & Chr(9) & rec!FCL_FECHA & Chr(9) & Format(rec!FCL_TOTAL, "#0.00") & Chr(9) & rec!CLI_RAZSOC _
                                & Chr(9) & rec!CLI_DOMICI & Chr(9) & rec!VEN_CODIGO _
                                & Chr(9) & rec!EST_CODIGO & Chr(9) & Format(rec!RCL_SUCURSAL, "0000") & "-" & Format(rec!RCL_NUMERO, "00000000") _
                                & Chr(9) & rec!RCL_FECHA & Chr(9) & rec!FCL_BONIFICA _
                                & Chr(9) & rec!FCL_IVA & Chr(9) & rec!FCL_OBSERVACION _
                                & Chr(9) & rec!TCO_CODIGO & Chr(9) & rec!FPG_CODIGO _
                                & Chr(9) & rec!FCL_BONIPESOS & Chr(9) & rec!CLI_CODIGO _
                                & Chr(9) & rec!FCL_TOTAL
                rec.MoveNext
            Loop
            GrdModulos.SetFocus
        Else
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        End If
        
    Case 2 'BUSCA REMITO
        chkVendedor.Enabled = False
        txtVendedor.Enabled = False
        cmdBuscarVen.Enabled = False
                        
        sql = "SELECT RC.RCL_NUMERO, RC.RCL_SUCURSAL, RC.RCL_FECHA, C.CLI_RAZSOC, C.CLI_DOMICI"
        sql = sql & " FROM REMITO_CLIENTE RC,CLIENTE C"
        sql = sql & " WHERE"
        sql = sql & " RC.CLI_CODIGO=C.CLI_CODIGO"
        sql = sql & " AND RC.EST_CODIGO = 1"
        'sql = sql & " AND NP.VEN_CODIGO=V.VEN_CODIGO"
        If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
        'If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
        If Not IsNull(FechaDesde) Then sql = sql & " AND RC.RCL_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND RC.RCL_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY RC.RCL_SUCURSAL,RC.RCL_NUMERO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem "" & Chr(9) & Format(rec!RCL_SUCURSAL, "0000") & "-" & Format(rec!RCL_NUMERO, "00000000") _
                                & Chr(9) & rec!RCL_FECHA & Chr(9) & rec!CLI_RAZSOC _
                                & Chr(9) & rec!CLI_DOMICI & Chr(9) & ""
                rec.MoveNext
            Loop
            GrdModulos.SetFocus
        Else
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        End If
    End Select
    
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    rec.Close
End Sub

Private Sub cmdBuscarCli_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCliente.Text = frmBuscar.grdBuscar.Text
        txtCliente.SetFocus
        txtCliente_LostFocus
    Else
        txtCliente.SetFocus
    End If
End Sub

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtRazSocCli.Text = frmBuscar.grdBuscar.Text
        TxtCodigoCli_LostFocus
    Else
        TxtCodigoCli.SetFocus
    End If
End Sub

Private Sub cmdBuscarRemito_Click()
    TipoBusquedaDoc = 2 'BUSCA REMITOS
    GrdModulos.ColWidth(0) = 0 'TIPO FACTURA
    tabDatos.Tab = 1
    chkVendedor.Enabled = False
    
End Sub
Private Sub cmdBuscarVen_Click()
    frmBuscar.TipoBusqueda = 4
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtVendedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtDesVen.Text = frmBuscar.grdBuscar.Text
        txtVendedor.SetFocus
    Else
        txtVendedor.SetFocus
    End If
End Sub

Private Sub cmdGrabar_Click()
    
    If ValidarFactura = False Then Exit Sub
    If MsgBox("¿Confirma Factura?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayErrorFactura
    
    DBConn.BeginTrans
    sql = "SELECT * FROM FACTURA_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
    sql = sql & " AND FCL_NUMERO = " & XN(txtNroFactura)
    sql = sql & " AND FCL_SUCURSAL=" & XN(txtNroSucursal)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = True Then
        'NUEVA FACTURA
        sql = "INSERT INTO FACTURA_CLIENTE"
        sql = sql & " (TCO_CODIGO,FCL_NUMERO,FCL_SUCURSAL,FCL_FECHA,"
        sql = sql & "RCL_NUMERO,RCL_SUCURSAL,RCL_FECHA,FCL_BONIFICA,FCL_IVA,FPG_CODIGO,FCL_OBSERVACION,"
        sql = sql & "FCL_BONIPESOS,FCL_NUMEROTXT,FCL_SUBTOTAL,FCL_TOTAL,FCL_SALDO,EST_CODIGO,VEN_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & cboFactura.ItemData(cboFactura.ListIndex) & ","
        sql = sql & XN(txtNroFactura) & ","
        sql = sql & XN(txtNroSucursal) & ","
        sql = sql & XDQ(FechaFactura) & ","
        sql = sql & XN(txtNroRemito) & ","
        sql = sql & XN(txtRemSuc) & ","
        sql = sql & XDQ(FechaRemito) & ","
        sql = sql & XN(txtPorcentajeBoni) & ","
        If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then
            sql = sql & XN(txtAuxIva.Text) & ","
        Else
            sql = sql & XN(txtPorcentajeIva) & ","
        End If
        sql = sql & cboCondicion.ItemData(cboCondicion.ListIndex) & ","
        sql = sql & XS(txtObservaciones) & ","
        If chkBonificaEnPesos.Value = Checked Then
            sql = sql & "'S'" & "," 'BONIFICA EN PESOS
        ElseIf chkBonificaEnPorsentaje.Value = Checked Then
            sql = sql & "'N'" & "," 'BONIFICA EN PORCENTAJE
        Else
            sql = sql & "NULL" & "," 'NO HAY BONIFICACION
        End If
        sql = sql & XS(Format(txtNroFactura.Text, "00000000")) & ","
        If txtSubTotalBoni.Text <> "" Then 'SUBTOTAL
            sql = sql & XN(txtSubTotalBoni) & ","
        Else
            If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then
                sql = sql & XN(txtAuxNeto.Text) & ","
            Else
                sql = sql & XN(txtSubtotal) & ","
            End If
        End If
        sql = sql & XN(txtTotal) & ","
        If cboCondicion.ItemData(cboCondicion.ListIndex) = 1 Then
            sql = sql & XN("0") & "," 'SALDO FACTURA
        Else
            sql = sql & XN(txtTotal) & "," 'SALDO FACTURA
        End If
'        If cboCondicion.ItemData(cboCondicion.ListIndex) = 1 Then ' FORMA PAGO DE CONTADO
'            sql = sql & XN("0") & ","
'        Else
'            sql = sql & XN(txtTotal) & "," 'SALDO FACTURA
'        End If
        
        sql = sql & "3" & "," 'ESTADO DEFINITIVO
        sql = sql & XN(txtNroVendedor) & ")" 'SALDO FACTURA
        DBConn.Execute sql
           
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                sql = "INSERT INTO DETALLE_FACTURA_CLIENTE"
                sql = sql & " (TCO_CODIGO,FCL_NUMERO,FCL_SUCURSAL,FCL_FECHA,"
                sql = sql & "DFC_NROITEM,PTO_CODIGO,DFC_CANTIDAD,DFC_PRECIO,DFC_BONIFICA,DFC_DETALLE)"
                sql = sql & " VALUES ("
                sql = sql & cboFactura.ItemData(cboFactura.ListIndex) & ","
                sql = sql & XN(txtNroFactura) & ","
                sql = sql & XN(txtNroSucursal) & ","
                sql = sql & XDQ(FechaFactura) & ","
                sql = sql & I & ","
                sql = sql & XS(grdGrilla.TextMatrix(I, 0), True) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 4)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(I, 1)) & ")"
                DBConn.Execute sql
            End If
        Next
        
        'ACTUALIZO EL STOCK CUANDO EL REMITO ES DEFINITIVO (STOCK PENDIENTE)
         For I = 1 To grdGrilla.Rows - 1
             If grdGrilla.TextMatrix(I, 0) <> "" Then
                     sql = "UPDATE DETALLE_STOCK"
                     sql = sql & " SET"
                     sql = sql & " DST_STKPEN = DST_STKPEN - " & XN(grdGrilla.TextMatrix(I, 2))
                     sql = sql & " ,DST_STKFIS = DST_STKFIS - " & XN(grdGrilla.TextMatrix(I, 2))
                     sql = sql & " WHERE STK_CODIGO= 1"
                     sql = sql & " AND PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "'"
                     DBConn.Execute sql
             End If
         Next
         
        'ACTUALIZO ESTADO DE REMITOS MULTIPLES
        If txtNroRemito.Text > 90000000 Then
            sql = "UPDATE REMITOS_FACTURA SET "
            sql = sql & " FCL_NUMERO =" & XN(txtNroFactura.Text)
            sql = sql & " ,FCL_SUCURSAL =" & XN(txtNroSucursal.Text)
            sql = sql & " WHERE REF_REMITOM =" & XN(txtNroRemito.Text)
            DBConn.Execute sql
            
            sql = "SELECT * FROM REMITOS_FACTURA "
            sql = sql & " WHERE REF_REMITOM = " & XN(txtNroRemito)
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
             Do While Rec1.EOF = False
                sql = "UPDATE REMITO_CLIENTE SET EST_CODIGO=3"
                sql = sql & " WHERE"
                sql = sql & " RCL_NUMERO=" & Rec1!RCL_NUMERO
                sql = sql & " AND RCL_SUCURSAL=" & Rec1!RCL_SUCURSAL
                DBConn.Execute sql
                Rec1.MoveNext
             Loop
            
            End If
            Rec1.Close
       End If
         
        'CAMBIO ESTADO DEL REMITO (LE PONGO DEFINITIVO) // EL REMITO UNIFICADO QUEDA COMO DEFINITIVO
        
        sql = "UPDATE REMITO_CLIENTE SET EST_CODIGO=3"
        sql = sql & " WHERE"
        sql = sql & " RCL_NUMERO=" & XN(txtNroRemito)
        sql = sql & " AND RCL_SUCURSAL=" & XN(txtRemSuc)
        DBConn.Execute sql
        
        
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO A LA FACTURA QUE CORRESPONDE
        sql = "SELECT * FROM PARAMETROS"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            'If Rec1!REP_CODIGO = cboRep.ItemData(cboRep.ListIndex) Then
                Select Case cboFactura.ItemData(cboFactura.ListIndex)
                    Case 1
                        sql = "UPDATE PARAMETROS SET FACTURA_A=" & XN(txtNroFactura)
                    Case 2
                        sql = "UPDATE PARAMETROS SET FACTURA_B=" & XN(txtNroFactura)
                End Select
                    DBConn.Execute sql
            'End If
        End If
        Rec1.Close
        
        'ACTUALIZO LA CUENTA CORRIENTE DEL CLIENTE
        DBConn.Execute AgregoCtaCteCliente(TxtCodigoCli, CStr(cboFactura.ItemData(cboFactura.ListIndex)) _
                                            , txtNroFactura, txtNroSucursal, _
                                            FechaFactura, txtTotal, "D", CStr(Date))
        DBConn.CommitTrans
    Else
        MsgBox "La Factura ya fue Registrada", vbCritical, TIT_MSGBOX
        DBConn.CommitTrans
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    cmdImprimir_Click
        
    
    If cboCondicion.ItemData(cboCondicion.ListIndex) = 2 Then
        'preguntar si quiere abrir el recibo y abrirlo
        If MsgBox("¿Desea realizar el Recibo para el cliente?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then
            CmdNuevo_Click
            Exit Sub
        End If
    
        frmReciboCliente.txtCodCliente.Text = TxtCodigoCli.Text
        frmReciboCliente.Show vbModal
    End If
    
    CmdNuevo_Click
    Exit Sub
    
HayErrorFactura:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarFactura() As Boolean
    If IsNull(FechaFactura.Value) Then
        MsgBox "La Fecha de la Factura es requerida", vbExclamation, TIT_MSGBOX
        FechaFactura.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If txtNroSucursal.Text = "" Then
        MsgBox "El número completo de la Factura es requerido", vbExclamation, TIT_MSGBOX
        txtNroSucursal.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If txtNroFactura.Text = "" Then
        MsgBox "El número de la Factura es requerido", vbExclamation, TIT_MSGBOX
        txtNroFactura.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If txtNroRemito.Text = "" Then
        MsgBox "El número del Remito es requerido", vbExclamation, TIT_MSGBOX
        txtNroRemito.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If IsNull(FechaRemito.Value) Then
        MsgBox "La Fecha del Remito es requerida", vbExclamation, TIT_MSGBOX
        FechaRemito.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If cboCondicion.ListIndex = -1 Then
        MsgBox "La Condición de pago es requerida", vbExclamation, TIT_MSGBOX
        cboCondicion.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If chkBonificaEnPesos.Value = Checked Or chkBonificaEnPorsentaje.Value = Checked Then
        If txtPorcentajeBoni.Text = "" Then
            MsgBox "Debe ingresar la Bonificación", vbExclamation, TIT_MSGBOX
            txtPorcentajeBoni.SetFocus
            ValidarFactura = False
            Exit Function
        End If
    End If
    If cboCondicion.ItemData(cboCondicion.ListIndex) = 1 Then ' FORMA PAGO DE CONTADO
        'If CDbl(txtTotal.Text) >= 1000 Then
            If MsgBox("Confirma pago en EFECTIVO?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
                ValidarFactura = True
            Else
                cboCondicion.SetFocus
                ValidarFactura = False
                Exit Function
            End If
        'End If
    End If
    ValidarFactura = True
End Function

Private Sub cmdImprimir_Click()
    If MsgBox("¿Confirma Impresión Factura?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'PONE A LA IMPRESORA  COMO PREDETERMINADA
    Dim X As Printer
    Dim mDriver As String
    mDriver = IMPRESORA
    For Each X In Printers
        If X.DeviceName = mDriver Then
            ' La define como predeterminada del sistema.
            Set Printer = X
            Exit For
        End If
    Next
'-----------------------------------
    Set_Impresora
    ImprimirFactura
End Sub

Public Sub ImprimirFactura()
    Dim Renglon As Double
    Dim canttxt As Integer
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Imprimiendo..."
    
    For w = 1 To 2 'SE IMPRIME POR DUPLICADO
      '-----IMPRESION DEL ENCABEZADO------------------
        ImprimirEncabezado
        
      '---- IMPRESION DE LA FACTURA ------------------
        'If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then
        '    Renglon = 10.4
        'Else
            Renglon = 9.4
        'End If
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                Imprimir 0.6, Renglon, False, grdGrilla.TextMatrix(I, 0)  'codigo
                canttxt = 0
                If Len(grdGrilla.TextMatrix(I, 1)) <= 36 Then
                    Imprimir 3.8, Renglon, False, grdGrilla.TextMatrix(I, 1) 'descripcion
                Else
                     'CortarCadena 3.8, Renglon, grdGrilla.TextMatrix(I, 1)
                    justifica_printer 3.8, 12.5, Renglon, grdGrilla.TextMatrix(I, 1)
                    
'                    Imprimir 3.2, renglon, False, Left(grdGrilla.TextMatrix(I, 1), 36) 'descripcion
'                    Imprimir 3.2, renglon + 0.5, False, Mid(grdGrilla.TextMatrix(I, 1), 37, 36) 'descripcion
'                    Imprimir 3.2, renglon + 1, False, Mid(grdGrilla.TextMatrix(I, 1), 74, 36) 'descripcion
'                    Imprimir 3.2, renglon + 1.5, False, Mid(grdGrilla.TextMatrix(I, 1), 111, 36) 'descripcion
'                    Imprimir 3.2, renglon + 2, False, Mid(grdGrilla.TextMatrix(I, 1), 148, 36) 'descripcion
'                    Imprimir 3.2, renglon + 2.5, False, Mid(grdGrilla.TextMatrix(I, 1), 185, 36) 'descripcion
'                    Imprimir 3.2, renglon + 3, False, Mid(grdGrilla.TextMatrix(I, 1), 222, 36) 'descripcion
                    canttxt = Len(grdGrilla.TextMatrix(I, 1))
                    canttxt = canttxt / 31 'es para sacar la cantidad de renglones
                    canttxt = Int(canttxt)
                    
                End If
                'If (grdGrilla.TextMatrix(I, 10) <> 6) Then 'SI ES UNA MAQUINARIA NO IMPRIMO ESTOS DATOS
                    Imprimir 13.15, Renglon, False, grdGrilla.TextMatrix(I, 2) 'cantidad
                    Imprimir 15.3, Renglon, False, grdGrilla.TextMatrix(I, 3) 'precio
                    Imprimir 15.7, Renglon, False, grdGrilla.TextMatrix(I, 4) 'bonificacion
                'End If
                Imprimir 17.6, Renglon, False, Valido_Importe(grdGrilla.TextMatrix(I, 6)) 'importe
                Renglon = Renglon + (canttxt * 0.5) + 0.5
            
                'IMPRIMO DATOS DE MAQUINARIA
                If grdGrilla.TextMatrix(I, 10) = 6 Then  'SI LINEA MAQUINARIA (hay que ver si es tractor o sembradora)
                    sql = "SELECT * FROM PRODUCTO "
                    sql = sql & "WHERE PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "'"
                    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                    If Rec1.EOF = False Then
                        If Not IsNull(Rec1!PTO_TIPO) Then
                            Imprimir 2, Renglon + 3, False, "Tipo.............: " & IIf(IsNull(Rec1!PTO_TIPO), "", Rec1!PTO_TIPO)
                        End If
                        If Not IsNull(Rec1!PTO_TIPMOD) Then
                            Imprimir 11, Renglon + 3, False, "Modelo.........: " & IIf(IsNull(Rec1!PTO_TIPMOD), "", Rec1!PTO_TIPMOD)
                        End If
                        If Not IsNull(Rec1!PTO_TRACCI) Then
                            Imprimir 2, Renglon + 3.4, False, "Tracción......: " & IIf(IsNull(Rec1!PTO_TRACCI), "", Rec1!PTO_TRACCI)
                        End If
                        If Not IsNull(Rec1!PTO_TIPO) Then
                            Imprimir 11, Renglon + 3.4, False, IIf(Rec1!PTO_CABINA = 1, "Con Cabina", "Sin Cabina")
                        End If
                        If Not IsNull(Rec1!PTO_MOTMAR) Then
                            Imprimir 2, Renglon + 3.8, False, "Motor Marca: " & IIf(IsNull(Rec1!PTO_MOTMAR), "", Rec1!PTO_MOTMAR)
                        End If
                        If Not IsNull(Rec1!PTO_MOTMOD) Then
                            Imprimir 11, Renglon + 3.8, False, "Modelo.........: " & IIf(IsNull(Rec1!PTO_MOTMOD), "", Rec1!PTO_MOTMOD)
                        End If
                        If Not IsNull(Rec1!PTO_ASPIRA) Then
                            Imprimir 2, Renglon + 4.2, False, "Aspiración...: " & IIf(IsNull(Rec1!PTO_ASPIRA), "", Rec1!PTO_ASPIRA)
                        End If
                        If Not IsNull(Rec1!PTO_MOTNRO) Then
                            Imprimir 11, Renglon + 4.2, False, "Motor Nro.....: " & IIf(IsNull(Rec1!PTO_MOTNRO), "", Rec1!PTO_MOTNRO)
                        End If
                        If Not IsNull(Rec1!PTO_CHASIS) Then
                            Imprimir 2, Renglon + 4.6, False, "Chasis Nro..: " & IIf(IsNull(Rec1!PTO_CHASIS), "", Rec1!PTO_CHASIS)
                        End If
                        If Not IsNull(Rec1!PTO_SERIE) Then
                            Imprimir 11, Renglon + 4.6, False, "Serie.............: " & IIf(IsNull(Rec1!PTO_SERIE), "", Rec1!PTO_SERIE)
                        End If
                        If Not IsNull(Rec1!PTO_NEUMDE) Then
                            Imprimir 2, Renglon + 5, False, "Neum. Del...: " & IIf(IsNull(Rec1!PTO_NEUMDE), "", Rec1!PTO_NEUMDE)
                        End If
                        If Not IsNull(Rec1!PTO_NEDECA) Then
                            Imprimir 11, Renglon + 5, False, "Cantidad.......: " & IIf(IsNull(Rec1!PTO_NEDECA), "", Rec1!PTO_NEDECA)
                        End If
                        If Not IsNull(Rec1!PTO_NEUMTR) Then
                            Imprimir 2, Renglon + 5.4, False, "Neum. Tra...: " & IIf(IsNull(Rec1!PTO_NEUMTR), "", Rec1!PTO_NEUMTR)
                        End If
                        If Not IsNull(Rec1!PTO_NETRCA) Then
                            Imprimir 11, Renglon + 5.4, False, "Cantidad.......: " & IIf(IsNull(Rec1!PTO_NETRCA), "", Rec1!PTO_NETRCA)
                        End If
                        If Not IsNull(Rec1!PTO_TIPO) Then
                            Imprimir 2, Renglon + 5.8, False, IIf(Rec1!PTO_KITCON = 1, "Con Kit Confort", "Sin Kit Confort")
                        End If
                        If Not IsNull(Rec1!PTO_SALHID) Then
                            Imprimir 11, Renglon + 5.8, False, "Salida Hidr....: " & IIf(IsNull(Rec1!PTO_SALHID), "", Rec1!PTO_SALHID)
                        End If
                        'If Not IsNull(Rec1!PTO_POSARA) Then
                        '    Imprimir 2, Renglon + 6.2, False, "Posic. Aran..: " & IIf(IsNull(Rec1!PTO_POSARA), "", Rec1!PTO_POSARA)
                        'End If
                        If Not IsNull(Rec1!PTO_CERFAB) Then
                            Imprimir 2, Renglon + 6.2, False, "Cert. Fabrica: " & IIf(IsNull(Rec1!PTO_CERFAB), "", Rec1!PTO_CERFAB)
                        End If
                        If Not IsNull(Rec1!PTO_OPCION1) Then
                            Imprimir 2, Renglon + 6.6, False, IIf(IsNull(Rec1!PTO_OPCION1), "", Rec1!PTO_OPCION1)
                        End If
                        If Not IsNull(Rec1!PTO_OPCION2) Then
                            Imprimir 11, Renglon + 6.6, False, IIf(IsNull(Rec1!PTO_OPCION2), "", Rec1!PTO_OPCION2)
                        End If
                    End If
                    Rec1.Close
                End If
            End If
        Next I
            '-----OBSERVACIONES---------------------
            If txtObservaciones.Text <> "" Then
                Imprimir 1, Renglon + 9.5, True, "Observ.: "
                'CortarCadena 3.2, Renglon + 9.5, Trim(txtObservaciones)
                justifica_printer 3.2, 14, Renglon + 9.5, Trim(txtObservaciones.Text)
                'Imprimir 5, & Trim(txtObservaciones.Text)
            End If
            'Imprimir 0, 16.5, True, "texto de bajo del detalle"
            '-------------IMPRIMO TOTALES--------------------
'            If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then
'               Imprimir 17.2, 20.4, True, txtSubtotal.Text
'
'                 'If txtPorcentajeBoni.Text <> "" Then
'                 '    Imprimir 17.2, 22.3, True, txtSubTotalBoni.Text
'                 'Else
'                 '    Imprimir 17.2, 22.3, True, txtSubtotal.Text
'                 'End If
'
'                Imprimir 14.8, 22.2, True, "    " & txtPorcentajeIva.Text
'                Imprimir 17.2, 22.2, True, txtImporteIva.Text
'                If IMPRIMO = "HP" Then
'                    Imprimir 16.7, 25, True, txtTotal.Text
'                Else
'                    Imprimir 16.7, 25.4, True, txtTotal.Text
'                End If
'            Else
            
                Imprimir 17.2, 22, True, txtSubtotal.Text
    
                If txtPorcentajeBoni.Text <> "" Then
                    Imprimir 14.8, 22.9, True, "    " & txtPorcentajeBoni.Text
                    Imprimir 17.2, 22.9, True, txtImporteBoni.Text
                    Imprimir 17.2, 23.8, True, txtSubTotalBoni.Text
                Else
                    Imprimir 17.2, 23.8, True, txtSubtotal.Text
                End If
                
                Imprimir 14.8, 24.7, True, "    " & txtPorcentajeIva.Text
                Imprimir 17.2, 24.7, True, txtImporteIva.Text
                If IMPRIMO = "HP" Then
                    Imprimir 16.7, 24.65, True, txtTotal.Text
                Else
                    Imprimir 16.7, 26.4, True, txtTotal.Text
                End If
           ' End If
        Printer.EndDoc
    Next w
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Public Sub ImprimirEncabezado()
 '-----------IMPRIME EL ENCABEZADO DE LA FACTURA-------------------
    Dim año As String
    'año = String(4, Year(FechaFactura))
'    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then
'        año = Year(FechaFactura)
'        Imprimir 15.6, 4.1, False, Format(Day(FechaFactura), "00")
'        Imprimir 16.7, 4.1, False, Format(Month(FechaFactura), "00")
'        Imprimir 17.75, 4.1, False, Mid(año, 3, 2)
'
'    Else
        año = Year(FechaFactura)
        Imprimir 14.5, 3, False, Format(Day(FechaFactura), "00")
        Imprimir 16.1, 3, False, Format(Month(FechaFactura), "00")
        Imprimir 17.8, 3, False, Mid(año, 3, 2)
    'End If
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CUIT,C.CLI_INGBRU, L.LOC_DESCRI"
    sql = sql & ", P.PRO_DESCRI,CI.IVA_DESCRI,C.IVA_CODIGO"
    sql = sql & " FROM CLIENTE C, LOCALIDAD L, REMITO_CLIENTE RC,"
    sql = sql & " PROVINCIA P, CONDICION_IVA CI"
    sql = sql & " WHERE RC.RCL_NUMERO=" & XN(txtNroRemito)
    sql = sql & " AND RC.RCL_SUCURSAL=" & XN(txtRemSuc)
    sql = sql & " AND RC.RCL_FECHA=" & XDQ(FechaRemito)
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        
'        If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then
'            If Len(Trim(Rec1!CLI_RAZSOC)) < 36 Then
'                Imprimir 2.3, 6.9, False, Trim(Rec1!CLI_RAZSOC)
'            Else
'                CortarCadena 2.3, 6.5, Trim(Rec1!CLI_RAZSOC)
'            End If
'            Imprimir 10.5, 6.7, False, Trim(IIf(IsNull(Rec1!CLI_DOMICI), "", Rec1!CLI_DOMICI))
'            'REMITO
'            'Imprimir 13.8, 7.2, True, Format(txtNroRemito.Text, "00000000") & " del " & Format(FechaRemito.value, "dd/mm/yyyy")
'            Imprimir 10.5, 7.1, False, Trim(Rec1!LOC_DESCRI) & " - " & Trim(Rec1!PRO_DESCRI)
'            'Imprimir 1, 6.3, False, Trim(Rec1!IVA_DESCRI)
'            If Rec1!IVA_CODIGO = 1 Then
'                Imprimir 4.2, 7.75, False, "X"
'            Else
'                If Rec1!IVA_CODIGO = 4 Then 'Exento
'                    Imprimir 10.6, 7.75, False, "X"
'                Else
'                    Imprimir 7.5, 7.75, False, "X"
'                End If
'            End If
'            'Imprimir 1, 6.3, False, Trim(Rec1!IVA_DESCRI)
'            Imprimir 14.5, 7.75, False, IIf(IsNull(Rec1!CLI_CUIT), "", Format(Rec1!CLI_CUIT, "##-########-#"))
'            Imprimir 15, 8.1, False, IIf(IsNull(Rec1!CLI_INGBRU), "", Format(Rec1!CLI_INGBRU, "###-#####-##"))
'            If cboCondicion.ItemData(cboCondicion.ListIndex) = 1 Then
'                Imprimir 7.2, 8.4, False, "X"
'            Else
'                Imprimir 10.8, 8.4, False, "X"
'            End If
            
        'Else
            If Len(Trim(Rec1!CLI_RAZSOC)) < 36 Then
                Imprimir 2.3, 5.8, False, Trim(Rec1!CLI_RAZSOC)
            Else
                CortarCadena 2.3, 5.4, Trim(Rec1!CLI_RAZSOC)
            End If
            Imprimir 12.3, 5.4, False, Trim(IIf(IsNull(Rec1!CLI_DOMICI), "", Rec1!CLI_DOMICI))
            'REMITO
            'Imprimir 13.8, 7.2, True, Format(txtNroRemito.Text, "00000000") & " del " & Format(FechaRemito.value, "dd/mm/yyyy")
            Imprimir 12.3, 5.8, False, Trim(Rec1!LOC_DESCRI) & " - " & Trim(Rec1!PRO_DESCRI)
            'Imprimir 1, 6.3, False, Trim(Rec1!IVA_DESCRI)
            Select Case Rec1!IVA_CODIGO
            Case 1
                Imprimir 3.75, 6.85, False, "X"
            Case 3
                Imprimir 10.1, 6.85, False, "X"
            Case 4
                Imprimir 7.2, 6.85, False, "X"
            Case 5
                Imprimir 5.2, 6.85, False, "X"
            End Select
            
            
'            If Rec1!IVA_CODIGO = 1 Then
'                Imprimir 3.75, 6.85, False, "X"
'            Else
'                If Rec1!IVA_CODIGO = 4 Then 'Exento
'                    Imprimir 10.1, 6.85, False, "X"
'                Else
'                    Imprimir 7, 6.85, False, "X"
'                End If
'            End If
            'Imprimir 1, 6.3, False, Trim(Rec1!IVA_DESCRI)
            Imprimir 13, 6.85, False, IIf(IsNull(Rec1!CLI_CUIT), "", Format(Rec1!CLI_CUIT, "##-########-#"))
            Imprimir 13.8, 7.5, False, IIf(IsNull(Rec1!CLI_INGBRU), "", Format(Rec1!CLI_INGBRU, "###-#####-##"))
            If cboCondicion.ItemData(cboCondicion.ListIndex) = 1 Then
                Imprimir 6.2, 7.8, False, "X"
            Else
                Imprimir 9.8, 7.8, False, "X"
            End If
        
        'End If
    End If
    Rec1.Close
     
    
    
    
'    Imprimir 0, 8, False, "Código"
'    Imprimir 2.5, 8, False, "Descripción"
'    Imprimir 10, 8, False, "Cantidad"
'    Imprimir 13, 8, False, "Precio"
'    Imprimir 15, 8, False, "Bonof."
'    Imprimir 17, 8, False, "Importe"
End Sub

Private Sub CmdNuevo_Click()
   For I = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(I, 0) = ""
        grdGrilla.TextMatrix(I, 1) = ""
        grdGrilla.TextMatrix(I, 2) = ""
        grdGrilla.TextMatrix(I, 3) = ""
        grdGrilla.TextMatrix(I, 4) = ""
        grdGrilla.TextMatrix(I, 5) = ""
        grdGrilla.TextMatrix(I, 6) = ""
        grdGrilla.TextMatrix(I, 7) = ""
        grdGrilla.TextMatrix(I, 8) = ""
        grdGrilla.TextMatrix(I, 9) = I
   Next
   LimpiarRemito
   txtNroFactura.Text = ""
   txtNroSucursal.Text = ""
   FechaFactura.Value = Date
   lblEstadoFactura.Caption = ""
   txtSubtotal.Text = ""
   txtTotal.Text = ""
   txtCodigoStock.Text = ""
   txtPorcentajeBoni.Text = ""
   txtPorcentajeIva.Text = ""
   txtImporteBoni.Text = ""
   txtSubTotalBoni.Text = ""
   txtImporteIva.Text = ""
   txtObservaciones.Text = ""
   cboCondicion.ListIndex = 0
   lblEstado.Caption = ""
   cmdGrabar.Enabled = True
   cmdImprimir.Enabled = False
   'BUSCO IVA
   BuscoIva
   'CARGO ESTADO
     Call BuscoEstado(1, lblEstadoFactura) 'ESTADO PENDIENTE
    VEstadoFactura = 1
    '--------------
    chkBonificaEnPorsentaje.Value = Unchecked
    chkBonificaEnPesos.Value = Unchecked
    FrameFactura.Enabled = True
    FrameRemito.Enabled = True
    tabDatos.Tab = 0
    FechaFactura.Value = Date
    cboFactura.ListIndex = 0
    TxtCodigoCli.Text = ""
    TxtCodigoCli_Change
    
End Sub

Private Sub cmdNuevoCliente_Click()
    ABMCliente.Show vbModal
    TxtCodigoCli.SetFocus
End Sub

Private Sub cmdNuevoRubro_Click()
    ABMFormaPago.Show vbModal
    cboCondicion.Clear
    LlenarComboFormaPago
    cboCondicion.SetFocus
End Sub

Private Sub cmdNuevoVendedor_Click()
    ABMVendedor.Show vbModal
    txtNroVendedor.SetFocus
End Sub

Private Sub cmdProducto_Click()
    frmBuscar.TipoBusqueda = 2
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtproducto.Text = frmBuscar.grdBuscar.Text
        txtproducto.SetFocus
        txtProducto_LostFocus
    Else
        txtproducto.SetFocus
    End If
End Sub

Private Sub cmdredondeo_Click()
    txtTotal.Enabled = True
    If txtTotal.Text <> "" Then
        txtTotal.Text = Round(txtTotal.Text, 0)
        txtTotal.Text = Valido_Importe(txtTotal.Text)
        txtImporteIva.Text = txtTotal.Text - txtSubtotal
        txtImporteIva.Text = Valido_Importe(txtImporteIva)
        'update en a BD
        'sql = "UPDATE FACTURA_CLIENTE SET"
        'SQL = SQL ""
    
    
    End If
    
    'If txtImporteIva.Text <> "" Then
    '    txtImporteIva.Text = Int(txtImporteIva)
    '    txtImporteIva.Text = Valido_Importe(txtImporteIva)
    'End If
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmFacturaCliente = Nothing
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
    frmBuscar.TipoBusqueda = 4
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtNroVendedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtNombreVendedor.Text = frmBuscar.grdBuscar.Text
        TxtCodigoCli.SetFocus
    Else
        txtNroVendedor.SetFocus
    End If
End Sub

Private Sub Form_Activate()
'FACTURACION AUTOMATICA
    If frmRemitoCliente.TxtCodigoCli.Text <> "" Then
        sql = "SELECT IVA_CODIGO FROM CLIENTE "
        sql = sql & " WHERE CLI_CODIGO = " & frmRemitoCliente.TxtCodigoCli
        sql = sql & " AND IVA_CODIGO IN (4,5)"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            'If rec!IVA_CODIGO = 4 Or rec!IVA_CODIGO = 5 Then
            BuscaCodigoProxItemData 2, cboFactura
        End If
        rec.Close
        If txtNroRemito.Text <> "" Then
            txtNroSucursal_LostFocus
            txtNroFactura_LostFocus
            txtNroVendedor_LostFocus
            txtNroRemito_LostFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        TipoBusquedaDoc = 1 'BUSCA FACTURAS
        GrdModulos.ColWidth(0) = 1300 'TIPO FACTURA
        tabDatos.Tab = 1
        frameBuscar.Caption = "Buscar Facturas por...."
        chkVendedor.Enabled = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl.Name <> "grdGrilla" And _
        Me.ActiveControl.Name <> "txtEdit" And _
        KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)
           
    grdGrilla.FormatString = "Código|Descripción|Cantidad|Precio|Bonif.|Pre.Bonif.|Importe|Rubro|Linea|Orden"
    grdGrilla.ColWidth(0) = 1500  'CODIGO
    grdGrilla.ColWidth(1) = 4300 'DESCRIPCION
    grdGrilla.ColWidth(2) = 1000 'CANTIDAD
    grdGrilla.ColWidth(3) = 1100 'PRECIO
    grdGrilla.ColWidth(4) = 800 'BONOFICACION
    grdGrilla.ColWidth(5) = 800 'PRE BONIFICACION
    grdGrilla.ColWidth(6) = 1100 'IMPORTE
    grdGrilla.ColWidth(7) = 2100 'RUBRO
    grdGrilla.ColWidth(8) = 2100 'LINEA
    grdGrilla.ColWidth(9) = 0    'ORDEN
    grdGrilla.ColWidth(10) = 0   'LNA_CODIGO
    grdGrilla.Cols = 11
    grdGrilla.Rows = 1
    For I = 2 To 25
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                             & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & (I - 1)
    Next
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "Tipo Fac|^Número|^Fecha|Importe|Cliente|Domicilio|Vendedor|Cod_Estado|" _
                              & "REMITO_NUMERO|REMITO_FECHA|PORCENTAJE BONIFICA|PORCENTAJE IVA|" _
                              & "OBSERVACIONES|COD TIPO COMPROBANTE|COD CONDICION VENTA|" _
                              & "BONIFICA EN PESOS|CLI_CODIGO|FCL_TOTAL"
    GrdModulos.ColWidth(0) = 900  'TIPO FACTURA
    GrdModulos.ColWidth(1) = 1500 'NUMERO
    GrdModulos.ColWidth(2) = 1100 'FECHA
    GrdModulos.ColWidth(3) = 1100 'IMPORTE
    GrdModulos.ColWidth(4) = 3500 'CLIENTE
    GrdModulos.ColWidth(5) = 3400 'Domicilio
    GrdModulos.ColWidth(6) = 0    'VENDEDOR
    GrdModulos.ColWidth(7) = 0    'COD_ESTADO
    GrdModulos.ColWidth(8) = 0    'REMITO_NUMERO
    GrdModulos.ColWidth(9) = 0    'REMITO_FECHA
    GrdModulos.ColWidth(10) = 0    'PORCENTAJE BONIFICA
    GrdModulos.ColWidth(11) = 0   'PORCENTAJE IVA
    GrdModulos.ColWidth(12) = 0   'OBSERVACIONES
    GrdModulos.ColWidth(13) = 0   'COD TIPO COMPROBANTE
    GrdModulos.ColWidth(14) = 0   'COD CONDICION VENTA
    GrdModulos.ColWidth(15) = 0   'BONIFICA EN PESOS
    GrdModulos.ColWidth(16) = 0   'REPRESENTADA
    GrdModulos.ColWidth(17) = 0   'Total
    GrdModulos.Rows = 1
    GrdModulos.Cols = 18
    '------------------------------------
   
    '------------------------------------
    lblEstado.Caption = ""
    'CARGO COMBO CON LOS TIPOS DE FACTURA
    LlenarComboFactura
    'CARGO COMBO CON LAS CONDICIONES DE VENTA
    LlenarComboFormaPago
    
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoFactura) 'ESTADO PENDIENTE
    If lblEstadoFactura.Caption = "PENDIENTE" Then
        cmdImprimir.Enabled = False
        cmdGrabar.Enabled = True
    End If
    If lblEstadoFactura.Caption = "ANULADO" Then
        cmdImprimir.Enabled = False
        cmdGrabar.Enabled = False
    End If
    If lblEstadoFactura.Caption = "DEFINITIVO" Then
       cmdImprimir.Enabled = True
       cmdGrabar.Enabled = False
    End If
    
    
    VEstadoFactura = 1
    FechaFactura.Value = Date
    TipoBusquedaDoc = 1 'ESTO ES PARA BUSCAR FACTURA(1), (2)PARA BUSCAR REMITOS
    tabDatos.Tab = 0
    'BUSCO IVA
    BuscoIva
   
    'sql = "DELETE FROM REMITO_CLIENTE WHERE RCL_NUMERO =185 "
    'DBConn.Execute sql
    
    sql = "SELECT RC.RCL_NUMERO,RC.RCL_SUCURSAL"
    sql = sql & " FROM FACTURA_CLIENTE FC, REMITO_CLIENTE RC"
    sql = sql & " WHERE FC.RCL_NUMERO = RC.RCL_NUMERO "
    sql = sql & " AND FC.RCL_SUCURSAL = RC.RCL_SUCURSAL "
    sql = sql & " AND FC.TCO_CODIGO = 2 AND RC.EST_CODIGO = 1"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            sql = "UPDATE REMITO_CLIENTE "
            sql = sql & " SET EST_CODIGO = 3"
            sql = sql & " WHERE RCL_NUMERO = " & rec!RCL_NUMERO
            sql = sql & " AND RCL_SUCURSAL = " & rec!RCL_SUCURSAL
            DBConn.Execute sql
            
            rec.MoveNext
        Loop
    End If
    rec.Close
    
    
End Sub
Private Sub BuscoIva()
    sql = "SELECT IVA FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtPorcentajeIva.Text = IIf(IsNull(rec!IVA), "", Format(rec!IVA, "0.00"))
    End If
    rec.Close
End Sub

Private Sub LlenarComboFactura()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'FACT%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboFactura.AddItem rec!TCO_DESCRI
            cboFactura.ItemData(cboFactura.NewIndex) = rec!TCO_CODIGO
            cboFactura1.AddItem rec!TCO_DESCRI
            cboFactura1.ItemData(cboFactura.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboFactura.ListIndex = 0
        cboFactura1.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboFormaPago()
    sql = "SELECT * FROM FORMA_PAGO"
    sql = sql & " ORDER BY FPG_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboCondicion.AddItem rec!FPG_DESCRI
            cboCondicion.ItemData(cboCondicion.NewIndex) = rec!FPG_CODIGO
            rec.MoveNext
        Loop
        cboCondicion.ListIndex = 0
    End If
    rec.Close
End Sub

Private Function BuscoUltimaFactura(TipoFac As Integer) As String
    'ACA BUSCA EL NUMERO DE REMITO SIGUIENTE AL ULTIMO CARGADO
    sql = "SELECT (FACTURA_A) + 1 AS FAC_A, (FACTURA_B) + 1 AS FAC_B, SUCURSAL"
    sql = sql & " FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtNroSucursal.Text = Format(rec!Sucursal, "0000")
        Select Case TipoFac
            Case 1
                BuscoUltimaFactura = IIf(IsNull(rec!FAC_A), 1, rec!FAC_A)
            Case 2
                BuscoUltimaFactura = IIf(IsNull(rec!FAC_B), 1, rec!FAC_B)
            Case 3
                MsgBox "No hay Facturas del tipo C", vbExclamation, TIT_MSGBOX
                cboFactura.SetFocus
        End Select
    End If
    rec.Close
End Function

Private Sub grdGrilla_Click()
    If grdGrilla.Rows > 1 Then
        grdGrilla.ToolTipText = grdGrilla.TextMatrix(grdGrilla.RowSel, 1)
    End If
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 4
            VBonificacion = 0
            grdGrilla.Text = ""
            grdGrilla.Col = 5
            grdGrilla.Text = ""
            VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)))
            grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Valido_Importe(CStr(VBonificacion))
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            txtTotal.Text = txtSubtotal.Text
            grdGrilla.Col = 4
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
            Case 4
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = "" Then
                    txtObservaciones.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
       (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Or (grdGrilla.Col = 4) Then
        If KeyAscii = vbKeyReturn Then
            If grdGrilla.Col = 4 Then
                If grdGrilla.row < grdGrilla.Rows - 1 Then
                    grdGrilla.row = grdGrilla.row + 1
                    grdGrilla.Col = 4
                Else
                    SendKeys "{TAB}"
                End If
            Else
                grdGrilla.Col = grdGrilla.Col + 1
            End If
        Else
            
            If grdGrilla.Col = 4 Then
                If KeyAscii > 47 And KeyAscii < 58 Then
                    EDITAR grdGrilla, txtEdit, KeyAscii
                End If
            End If
        End If
    End If
End Sub

Private Sub grdGrilla_LeaveCell()
    If txtEdit.Visible = False Then Exit Sub
    grdGrilla = txtEdit.Text
    txtEdit.Visible = False
End Sub

Private Sub grdGrilla_GotFocus()
    If grdGrilla.Rows > 1 Then
        If txtEdit.Visible = False Then
            grdGrilla.Col = 4
            Exit Sub
        End If
        grdGrilla = txtEdit.Text
        txtEdit.Visible = False
    End If
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        CmdNuevo_Click
        Select Case TipoBusquedaDoc
        Case 1 'BUSCA FACTURA
            Set Rec1 = New ADODB.Recordset
            lblEstado.Caption = "Buscando..."
            Screen.MousePointer = vbHourglass
            'CABEZA FACTURA
            Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 13)), cboFactura)
            txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
            txtNroFactura.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
            FechaFactura.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
            Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 7)), lblEstadoFactura)
            VEstadoFactura = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 7))
            If VEstadoFactura = 1 Then ' Pendiente
                cmdImprimir.Enabled = False
                cmdGrabar.Enabled = True
            End If
            If VEstadoFactura = 2 Then 'Anulada
                cmdImprimir.Enabled = False
                cmdGrabar.Enabled = False
            End If
            If VEstadoFactura = 3 Then ' definitiva
                cmdImprimir.Enabled = True
                cmdGrabar.Enabled = False
            End If

            
            
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 12) <> "" Then
                txtObservaciones.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 12))
            End If
            'CABEZA REMITO
            txtRemSuc.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 8), 4)
            txtNroRemito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 8), 8)
            FechaRemito.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 9)
            
            
            TxtCodigoCli.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 16)
            TxtCodigoCli_LostFocus
            
            txtNroVendedor.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
            txtNroVendedor_LostFocus
            
            
            
            'CONDICION VENTA
            Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 14)), cboCondicion)
            '----BUSCO DETALLE DE LA FACTURA------------------
            sql = "SELECT DFC.*, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI,P.LNA_CODIGO"
            sql = sql & " FROM DETALLE_FACTURA_CLIENTE DFC, PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE DFC.FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8))
            sql = sql & " AND DFC.FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
            sql = sql & " AND DFC.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 13))
            sql = sql & " AND DFC.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " ORDER BY DFC.DFC_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(I, 1) = Rec1!DFC_DETALLE
                    grdGrilla.TextMatrix(I, 2) = Rec1!DFC_CANTIDAD
                    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DFC_PRECIO)
                    If IsNull(Rec1!DFC_BONIFICA) Then
                        grdGrilla.TextMatrix(I, 4) = ""
                    Else
                        grdGrilla.TextMatrix(I, 4) = Valido_Importe(Rec1!DFC_BONIFICA)
                    End If
                    VBonificacion = 0
                    If Not IsNull(Rec1!DFC_BONIFICA) Then
                        VBonificacion = (((CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO)) * CDbl(Rec1!DFC_BONIFICA)) / 100)
                        VBonificacion = ((CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO)) - VBonificacion)
                        grdGrilla.TextMatrix(I, 5) = Valido_Importe(CStr(VBonificacion))
                        grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
                    Else
                        VBonificacion = (CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO))
                        grdGrilla.TextMatrix(I, 5) = ""
                        grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
                    End If
                    grdGrilla.TextMatrix(I, 7) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 8) = Rec1!LNA_DESCRI
                    grdGrilla.TextMatrix(I, 9) = Rec1!DFC_NROITEM
                    grdGrilla.TextMatrix(I, 10) = Rec1!LNA_CODIGO
                    I = I + 1
                    Rec1.MoveNext
                Loop
                VBonificacion = 0
            End If
            Rec1.Close
            '--CARGO LOS TOTALES----
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            'txtTotal.Text = txtSubtotal.Text
            
            
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 15) = "S" Then
                chkBonificaEnPesos.Value = Checked
            ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 15) = "N" Then
                chkBonificaEnPorsentaje.Value = Checked
            Else
                chkBonificaEnPesos.Value = Unchecked
                chkBonificaEnPorsentaje.Value = Unchecked
            End If
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 10) <> "" Then
                txtPorcentajeBoni.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 10)
                txtPorcentajeBoni_LostFocus
            End If
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 11) <> "" Then
                If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then
                    txtPorcentajeIva = "0,00"
                Else
                    txtPorcentajeIva = GrdModulos.TextMatrix(GrdModulos.RowSel, 11)
                End If
                txtPorcentajeIva_LostFocus
            End If
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            '--------------
            FrameFactura.Enabled = False
            FrameRemito.Enabled = False
            '--------------
            tabDatos.Tab = 0
            cboCondicion.SetFocus
            txtTotal.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 17))
        '----------------------------------------------------------
        Case 2 'BUSCA REMITO
        
            lblEstado.Caption = "Buscando..."
            Screen.MousePointer = vbHourglass
            
            txtRemSuc.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
            txtNroRemito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
            FechaRemito.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
            
            'grillaRemito.TextMatrix(0, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 3)
            'grillaRemito.TextMatrix(1, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 4)
            'grillaRemito.TextMatrix(2, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
        
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            tabDatos.Tab = 0
            txtNroRemito_LostFocus
            cboCondicion.SetFocus
        End Select
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    txtCliente.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    txtVendedor.Enabled = False
    cboFactura1.Enabled = False
    cmdBuscarCli.Enabled = False
    cmdGrabar.Enabled = False
    cmdBuscarVen.Enabled = False
    'LimpiarBusqueda
    If Me.Visible = True Then chkCliente.SetFocus
    If TipoBusquedaDoc = 1 Then
        frameBuscar.Caption = "Buscar Factura por..."
        chkTipoFactura.Enabled = True
        chkVendedor.Enabled = True
        chkanuladas.Enabled = True
        
    Else
        frameBuscar.Caption = "Buscar Remitos Pendientes por..."
        chkTipoFactura.Enabled = False
        cboFactura1.Enabled = False
        chkVendedor.Enabled = False
        chkanuladas.Enabled = False
    End If
  Else
    If VEstadoFactura = 1 Then
        cmdGrabar.Enabled = True
    Else
        cmdGrabar.Enabled = False
    End If
  End If
End Sub

Private Sub LimpiarBusqueda()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    txtVendedor.Text = ""
    txtDesVen.Text = ""
    cboFactura1.ListIndex = 0
    GrdModulos.Rows = 1
    chkCliente.Value = Unchecked
    chkFecha.Value = Unchecked
    chkVendedor.Value = Unchecked
    chkTipoFactura.Value = Unchecked
End Sub



Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(txtCliente)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtDesCli.Text = ""
            txtCliente.SetFocus
        End If
        rec.Close
    End If
'    If chkSucursal.Value = Unchecked And chkFecha.Value = Unchecked _
'        And chkVendedor.Value = Unchecked And chkTipoFactura.Value = Unchecked _
'        And ActiveControl.Name <> "cmdBuscarCli" _
'        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub

Private Function BuscoCondicionIVA(IVACodigo As String) As String
    sql = "SELECT * FROM CONDICION_IVA"
    sql = sql & " WHERE IVA_CODIGO=" & XN(IVACodigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscoCondicionIVA = rec!IVA_DESCRI
    Else
        BuscoCondicionIVA = ""
    End If
    rec.Close
End Function

Private Sub TxtCodigoCli_Change()
    If TxtCodigoCli.Text = "" Then
        TxtCodigoCli.Text = ""
        txtRazSocCli.Text = ""
        txtCUIT.Text = ""
        txtIngBrutos.Text = ""
        txtCondicionIVA.Text = ""
        txtDomici.Text = ""
        txtlocalidad.Text = ""
        txtProvincia.Text = ""
        txtcodpos.Text = ""
    End If
End Sub

Private Sub TxtCodigoCli_LostFocus()
If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Or ActiveControl.Name = "CmdSalir" Then Exit Sub
    If TxtCodigoCli.Text <> "" Then
        sql = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CUIT,C.IVA_CODIGO,C.CLI_INGBRU,"
        sql = sql & "L.LOC_DESCRI,P.PRO_DESCRI,L.LOC_CODPOS"
        sql = sql & " FROM CLIENTE C, LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE "
        sql = sql & "C.LOC_CODIGO = L.LOC_CODIGO AND "
        sql = sql & "C.PRO_CODIGO = P.PRO_CODIGO AND "
        sql = sql & "L.PRO_CODIGO = P.PRO_CODIGO AND "
        sql = sql & "C.CLI_CODIGO =" & XN(TxtCodigoCli)
        'sql = sql & " AND CLI_ESTADO=1"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtRazSocCli.Text = Rec1!CLI_RAZSOC
            txtDomici.Text = IIf(IsNull(Rec1!CLI_DOMICI), "", Rec1!CLI_DOMICI)
            txtlocalidad.Text = Rec1!LOC_DESCRI
            txtProvincia.Text = Rec1!PRO_DESCRI
            txtCondicionIVA.Text = BuscoCondicionIVA(Rec1!IVA_CODIGO)
            txtCUIT.Text = IIf(IsNull(Rec1!CLI_CUIT), "NO INFORMADO", Format(Rec1!CLI_CUIT, "##-########-#"))
            txtIngBrutos.Text = IIf(IsNull(Rec1!CLI_INGBRU), "NO INFORMADO", Format(Rec1!CLI_INGBRU, "###-#####-##"))
            txtcodpos.Text = IIf(IsNull(Rec1!LOC_CODPOS), "", Rec1!LOC_CODPOS)
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtRazSocCli.Text = ""
            TxtCodigoCli.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    If grdGrilla.Col = 4 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    CarTexto KeyAscii
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    VBonificacion = 0
    If KeyCode = vbKeyF1 Then
        frmBuscar.TipoBusqueda = 2
        frmBuscar.CodListaPrecio = 0
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        frmBuscar.Show vbModal
    End If

    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
            Case 4
                If Trim(txtEdit) <> "" Then
                    If txtEdit.Text = ValidarPorcentaje(txtEdit) = False Then
                        Exit Sub
                    End If
                    VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6)) * CDbl(txtEdit.Text)) / 100)
                    VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6)) - VBonificacion)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Valido_Importe(CStr(VBonificacion))
                    txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                    txtTotal.Text = txtSubtotal.Text
                Else
                    MsgBox "Debe ingresar el Importe", vbExclamation, TIT_MSGBOX
                    grdGrilla.Col = 4
                End If
        End Select
        grdGrilla.SetFocus
    End If
    If KeyCode = vbKeyEscape Then
       txtEdit.Visible = False
       grdGrilla.SetFocus
    End If
End Sub

Private Function BuscoRepetetidos(Codigo As Long, Linea As Integer) As Boolean
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            If Codigo = CLng(grdGrilla.TextMatrix(I, 0)) And (I <> Linea) Then
                MsgBox "El producto ya fue elegido anteriormente", vbExclamation, TIT_MSGBOX
                BuscoRepetetidos = False
                Exit Function
            End If
        End If
    Next
    BuscoRepetetidos = True
End Function

Private Sub LimpiarRemito()
    txtRemSuc.Text = ""
    txtNroRemito.Text = ""
    FechaRemito.Value = Null
    txtCodigoStock.Text = ""
    'grillaRemito.TextMatrix(0, 1) = ""
    'grillaRemito.TextMatrix(1, 1) = ""
    'grillaRemito.TextMatrix(2, 1) = ""
End Sub

Private Function BuscoVendedor(Codigo As String) As String
    sql = "SELECT VEN_NOMBRE"
    sql = sql & " FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO=" & XN(Codigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        BuscoVendedor = Trim(rec!VEN_NOMBRE)
    Else
        BuscoVendedor = "No se encontro el Vendedor"
    End If
    rec.Close
End Function

Private Function BuscoCliente(Codigo As String) As String
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(Codigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            BuscoCliente = rec!CLI_RAZSOC
        Else
            BuscoCliente = "No se encontro el Cliente"
        End If
        rec.Close
End Function

Private Function BuscoSucursal(CodigoSuc As String, CodigoCli As String) As String
        sql = "SELECT * FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(CodigoSuc)
        sql = sql & " AND CLI_CODIGO=" & XN(CodigoCli)
        
        Set Rec1 = New ADODB.Recordset
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            BuscoSucursal = Rec1!SUC_DESCRI
        Else
            BuscoSucursal = "No se encontro la Sucursal"
        End If
        Rec1.Close
End Function

Private Sub txtNroFactura_GotFocus()
    SelecTexto txtNroFactura
End Sub

Private Sub txtNroFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroFactura_LostFocus()
    'BUSCO EL NUMERO DE FACTURA QUE CORRESPONDE
    txtNroFactura.Text = Format(BuscoUltimaFactura(cboFactura.ItemData(cboFactura.ListIndex)), "00000000")
End Sub

Private Sub txtNroRemito_LostFocus()
    If txtNroRemito.Text <> "" Then
        txtNroRemito.Text = Format(txtNroRemito.Text, "00000000")
        sql = "SELECT RC.RCL_NUMERO, RC.RCL_SUCURSAL, RC.RCL_FECHA, RC.EST_CODIGO, RC.STK_CODIGO, E.EST_DESCRI"
        sql = sql & " ,RC.CLI_CODIGO, RC.VEN_CODIGO"
        sql = sql & " FROM REMITO_CLIENTE RC,ESTADO_DOCUMENTO E,CLIENTE C"
        sql = sql & " WHERE RC.RCL_NUMERO=" & XN(txtNroRemito)
        sql = sql & " AND RC.RCL_SUCURSAL=" & XN(txtRemSuc.Text)
        sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
        sql = sql & " AND RC.EST_CODIGO=E.EST_CODIGO"
        
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec2.EOF = False Then
            If Rec2.RecordCount > 1 Then
                MsgBox "Hay mas de un Remito con el Número: " & txtNroRemito.Text, vbInformation, TIT_MSGBOX
                Rec2.Close
                cmdBuscarRemito_Click
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Buscando..."
            
            'CARGO CABECERA DEL REMITO
            FechaRemito.Value = Rec2!RCL_FECHA
            TxtCodigoCli.Text = Rec2!CLI_CODIGO
            TxtCodigoCli_LostFocus
            'grillaRemito.TextMatrix(0, 1) = BuscoCliente(Rec2!CLI_CODIGO)
            'grillaRemito.TextMatrix(1, 1) = BuscoSucursal(Rec2!SUC_CODIGO, Rec2!CLI_CODIGO)
            'grillaRemito.TextMatrix(2, 1) = BuscoVendedor(Rec2!VEN_CODIGO)
            'grillaRemito.TextMatrix(0, 2) = Rec2!CLI_CODIGO 'guardo el codigo del cliente
            txtCodigoStock.Text = Rec2!STK_CODIGO
            
            If Rec2!EST_CODIGO <> 1 Then
                MsgBox "El Remito número: " & txtNroRemito.Text & Chr(13) & Chr(13) & _
                       "No puede ser asignado a la Factura por su estado (" & Rec2!EST_DESCRI & ")", vbExclamation, TIT_MSGBOX
                cmdGrabar.Enabled = False
                Screen.MousePointer = vbNormal
                lblEstado.Caption = ""
                Rec2.Close
                LimpiarRemito
                If txtRemSuc.Enabled = True Then
                    txtRemSuc.SetFocus
                End If
                Exit Sub
            Else
                cmdGrabar.Enabled = True
            End If
            'SI EN LA NOTA DE PEDIDO SE ELIGIO UNA CONDICION DE PAGO LE MUESTRO LA MISMA AQUI
            'If Not IsNull(Rec2!FPG_CODIGO) Then
            '    Call BuscaCodigoProxItemData(Rec2!FPG_CODIGO, cboCondicion)
            'End If
            Rec2.Close
            
        '-----BUSCO LOS DATOS DEL DETALLE DEL REMITO----------------------------------
            sql = "SELECT DRC.*,P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI,P.LNA_CODIGO"
            sql = sql & " FROM DETALLE_REMITO_CLIENTE DRC, PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE DRC.RCL_NUMERO=" & XN(txtNroRemito)
            sql = sql & " AND DRC.RCL_SUCURSAL=" & XN(txtRemSuc)
            sql = sql & " AND DRC.RCL_FECHA=" & XDQ(FechaRemito)
            sql = sql & " AND DRC.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " ORDER BY DRC.DRC_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(I, 1) = IIf(IsNull(Rec1!DRC_DETALLE), Rec1!PTO_DESCRI, Rec1!DRC_DETALLE)
                    grdGrilla.TextMatrix(I, 2) = IIf(IsNull(Rec1!DRC_CANTIDAD), "1", Rec1!DRC_CANTIDAD)
                    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DRC_PRECIO)
                    grdGrilla.TextMatrix(I, 4) = ""
                    grdGrilla.TextMatrix(I, 5) = ""
                    grdGrilla.TextMatrix(I, 6) = Valido_Importe(grdGrilla.TextMatrix(I, 2) * (Rec1!DRC_PRECIO))
                    grdGrilla.TextMatrix(I, 7) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 8) = Rec1!LNA_DESCRI
                    'grdGrilla.TextMatrix(I, 9) = Rec1!DRC_NROITEM
                    grdGrilla.TextMatrix(I, 10) = Rec1!LNA_CODIGO
                    I = I + 1
                    Rec1.MoveNext
                Loop
                txtSubtotal.Text = Valido_Importe(SumaTotal)
                txtTotal.Text = txtSubtotal.Text
            End If
            Rec1.Close
            '--------------------------------------------------
            If grdGrilla.TextMatrix(1, 8) = "MAQUINARIA" Then 'pregunta si la linea es Maquinaria
                txtPorcentajeIva.Text = "10,50"
            Else
                txtPorcentajeIva.Text = "21,00"
            End If
            cboCondicion.SetFocus
            txtPorcentajeIva_LostFocus
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
        Else
            MsgBox "El Remito no existe", vbExclamation, TIT_MSGBOX
            If Rec2.State = 1 Then Rec2.Close
            LimpiarRemito
            txtNroRemito.SetFocus
        End If
    End If
End Sub

Private Function SumaTotal() As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 6) <> "" Then
            VTotal = VTotal + (CDbl(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3)))
        End If
    Next
    SumaTotal = Valido_Importe(CStr(VTotal))
End Function

Private Function SumaBonificacion() As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 6) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(I, 6))
        End If
    Next
    SumaBonificacion = Valido_Importe(CStr(VTotal))
End Function

Private Sub txtNroSucursal_GotFocus()
    SelecTexto txtNroSucursal
End Sub

Private Sub txtNroSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroSucursal_LostFocus()
    If txtNroSucursal.Text = "" Then
        txtNroSucursal.Text = Sucursal
    Else
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    End If
End Sub

Private Sub txtNroVendedor_Change()
    If txtNroVendedor.Text = "" Then
        txtNombreVendedor.Text = ""
    End If
End Sub

Private Sub txtNroVendedor_GotFocus()
    SelecTexto txtNroVendedor
End Sub

Private Sub txtNroVendedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroVendedor_LostFocus()
    If txtNroVendedor.Text = "" Then
        txtNroVendedor.Text = 1
    End If
    sql = "SELECT VEN_NOMBRE"
    sql = sql & " FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO=" & XN(txtNroVendedor)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        txtNombreVendedor.Text = Trim(rec!VEN_NOMBRE)
    Else
        MsgBox "El Vendedor no existe", vbExclamation, TIT_MSGBOX
        txtNombreVendedor.Text = ""
        txtNroVendedor.SetFocus
    End If
    rec.Close
End Sub

Private Sub txtProducto_Change()
    If txtproducto.Text = "" Then
        txtdesProducto.Text = ""
    End If
End Sub

Private Sub txtProducto_GotFocus()
    SelecTexto txtproducto
End Sub

Private Sub txtproducto_KeyPress(KeyAscii As Integer)
'        KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtProducto_LostFocus()
    If txtproducto.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT PTO_DESCRI FROM PRODUCTO"
        sql = sql & " WHERE PTO_CODIGO LIKE '" & txtproducto & "'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtdesProducto.Text = rec!PTO_DESCRI
        Else
            MsgBox "El Producto no existe", vbExclamation, TIT_MSGBOX
            txtdesProducto.Text = ""
            txtproducto.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtRemSuc_GotFocus()
    txtRemSuc.Text = Sucursal
    SelecTexto txtRemSuc
End Sub

Private Sub txtRemSuc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtRemSuc_LostFocus()
    If txtRemSuc.Text = "" Then
        txtRemSuc.Text = Sucursal
    Else
        txtRemSuc.Text = Format(txtRemSuc, "0000")
    End If
End Sub

Private Sub txtObservaciones_GotFocus()
    SelecTexto txtObservaciones
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtPorcentajeBoni_GotFocus()
    SelecTexto txtPorcentajeBoni
End Sub

Private Sub txtPorcentajeBoni_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorcentajeBoni, KeyAscii)
End Sub

Private Sub txtPorcentajeBoni_LostFocus()
    If txtPorcentajeBoni.Text <> "" And txtSubtotal.Text <> "" Then
        If chkBonificaEnPorsentaje.Value = Checked Then
            If ValidarPorcentaje(txtPorcentajeBoni) = False Then
                txtPorcentajeBoni.SetFocus
                Exit Sub
            End If
            txtImporteBoni.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeBoni.Text)) / 100
            txtImporteBoni.Text = Valido_Importe(txtImporteBoni.Text)
            txtTotal.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
            txtSubTotalBoni.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtSubTotalBoni.Text = Valido_Importe(txtSubTotalBoni.Text)
            txtPorcentajeIva_LostFocus
        ElseIf chkBonificaEnPesos.Value = Checked Then
            txtPorcentajeBoni.Text = Valido_Importe(txtPorcentajeBoni.Text)
            txtImporteBoni.Text = Valido_Importe(txtPorcentajeBoni.Text)
            txtTotal.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
            txtSubTotalBoni.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtSubTotalBoni.Text = Valido_Importe(txtSubTotalBoni.Text)
        Else
            txtPorcentajeBoni.Text = ""
            txtImporteBoni.Text = ""
            MsgBox "Debe elegir como bonifica", vbExclamation, TIT_MSGBOX
            chkBonificaEnPorsentaje.SetFocus
        End If
    End If
End Sub

Private Sub txtPorcentajeIva_GotFocus()
    SelecTexto txtPorcentajeIva
End Sub

Private Sub txtPorcentajeIva_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorcentajeIva, KeyAscii)
End Sub

Private Sub txtPorcentajeIva_LostFocus()
    If txtPorcentajeIva.Text <> "" And txtSubtotal.Text <> "" Then
        If ValidarPorcentaje(txtPorcentajeIva) = False Then
            txtPorcentajeIva.SetFocus
            Exit Sub
        End If
        If txtImporteBoni.Text <> "" Then
            txtImporteIva.Text = (CDbl(txtSubTotalBoni.Text) * CDbl(txtPorcentajeIva.Text)) / 100
            txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
            txtTotal.Text = CDbl(txtSubTotalBoni.Text) + CDbl(txtImporteIva.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
        Else
            txtImporteIva.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
            txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
            txtTotal.Text = CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
        End If
    End If
End Sub

Private Sub txtTotal_LostFocus()
    txtTotal.Enabled = False
End Sub

Private Sub txtVendedor_Change()
    If txtVendedor.Text = "" Then
        txtDesVen.Text = ""
    End If
End Sub

Private Sub txtVendedor_GotFocus()
    SelecTexto txtVendedor
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtVendedor_LostFocus()
    If txtVendedor.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT VEN_NOMBRE"
        sql = sql & " FROM VENDEDOR"
        sql = sql & " WHERE VEN_CODIGO=" & XN(txtVendedor)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            txtDesVen.Text = Trim(rec!VEN_NOMBRE)
        Else
            MsgBox "El Vendedor no existe", vbExclamation, TIT_MSGBOX
            txtDesVen.Text = ""
            txtVendedor.SetFocus
        End If
        rec.Close
    End If
'    If chkFecha.Value = Unchecked And chkTipoFactura.Value = Unchecked _
'    And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub
Private Sub CortarCadena(COLUMNA As Double, Renglon As Double, Cadena As String)
    Dim salto, Max, inf, I, leer, leerb As Integer
    Dim salto1, salto2, salto3, salto4, salto5, salto6, salto7 As Integer
    Dim salto1b, salto2b, salto3b, salto4b, salto5b, salto6b, salto7b As Integer
    Dim cadena1 As String
    Dim cadena2 As String
    Dim cadena3 As String
    Dim cadena4 As String
    Dim cadena5 As String
    Dim cadena6 As String
    Dim cadena7 As String
    Dim cadena8 As String
    
    cadena1 = ""
    cadena2 = ""
    cadena3 = ""
    cadena4 = ""
    cadena5 = ""
    cadena6 = ""
    cadena7 = ""
    
    
    salto = 1
    Max = 36 * salto
    inf = Max - 10
    'falta = 0
    'If Len(cadena) > 35 Then
        For I = 1 To Len(Cadena)
            If (Mid(Cadena, I, 1) = " ") And (I > inf) And (I < Max) Or (I > Max) Then
                
                    If salto = 1 Then
                    salto1 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        cadena1 = Left(Cadena, I)
                        cadena2 = Mid(Cadena, salto1, Max)
                        'Imprimir 3.2, renglon, False, Left(grdGrilla.TextMatrix(I, 1), 36) 'descripcion
                        'Imprimir 3.2, renglon + 0.5, False, Mid(grdGrilla.TextMatrix(I, 1), 37, 36) 'descripcion
                    
                    Else
                        cadena1 = Left(Cadena, I)
                    End If
                      'descripcion
                End If
                If salto = 2 Then
                    leer = I - salto1
                    salto2 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto1b = I
                        leerb = Len(Cadena) + 1 - salto1b
                        cadena2 = Mid(Cadena, salto1, leer)
                        cadena3 = Mid(Cadena, salto1b, leerb)  'descripcion
                    Else
                        cadena2 = Mid(Cadena, salto1, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 3 Then
                    Max = 36 + I
                    inf = Max - 10
                    leer = I - salto2
                    salto3 = I
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto2b = I
                        leerb = Len(Cadena) + 1 - salto2b
                        cadena3 = Mid(Cadena, salto2, leer)
                        cadena4 = Mid(Cadena, salto2b, leerb)
                    Else
                        cadena3 = Mid(Cadena, salto2, leer)  'descripcion
                    End If
                    
                End If
                If salto = 4 Then
                    leer = I - salto3
                    salto4 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto3b = I
                        leerb = Len(Cadena) + 1 - salto3b
                        cadena4 = Mid(Cadena, salto3, leer)
                        cadena5 = Mid(Cadena, salto3b, leerb)  'descripcion
                    Else
                         cadena4 = Mid(Cadena, salto3, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 5 Then
                    leer = I - salto4
                    salto5 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto4b = I
                        leerb = Len(Cadena) + 1 - salto4b
                        cadena5 = Mid(Cadena, salto4, leer)
                        cadena6 = Mid(Cadena, salto4b, leerb)  'descripcion
                    Else
                        cadena5 = Mid(Cadena, salto4, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 6 Then
                    leer = I - salto5
                    salto6 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto5b = I
                        leerb = Len(Cadena) + 1 - salto5b
                        cadena6 = Mid(Cadena, salto5, leer)
                        cadena7 = Mid(Cadena, salto5b, leerb)  'descripcion
                        
                    Else
                        cadena6 = Mid(Cadena, salto5, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 7 Then
                    leer = I - salto6
                    salto7 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto6b = I
                        leerb = Len(Cadena) + 1 - salto6b
                        cadena7 = Mid(Cadena, salto6, leer)
                        cadena8 = Mid(Cadena, salto6b, leerb)  'descripcion
                        
                    Else
                        cadena7 = Mid(Cadena, salto6, leer)  'descripcion
                    End If
                    
                End If
                
                salto = salto + 1
                'Max = valor * salto
                'inf = Max - 10
                
            End If
        Next
    
        Imprimir COLUMNA, Renglon, False, cadena1
        Imprimir COLUMNA, Renglon + 0.5, False, Trim(cadena2)
        Imprimir COLUMNA, Renglon + 1, False, Trim(cadena3)
        Imprimir COLUMNA, Renglon + 1.5, False, Trim(cadena4)
        Imprimir COLUMNA, Renglon + 2, False, Trim(cadena5)
        Imprimir COLUMNA, Renglon + 2.5, False, Trim(cadena6)
        Imprimir COLUMNA, Renglon + 3, False, Trim(cadena7)
        Imprimir COLUMNA, Renglon + 3.5, False, Trim(cadena8)
    'Else
    '    cadena1 = cadena
    '    MsgBox cadena1
    'End If
    
End Sub
