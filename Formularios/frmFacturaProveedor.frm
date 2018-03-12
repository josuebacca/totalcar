VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmfacturaproveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura de Proveedores...."
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   570
      Left            =   10185
      TabIndex        =   15
      Top             =   7575
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   570
      Left            =   8160
      TabIndex        =   13
      Top             =   7575
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   570
      Left            =   9172
      TabIndex        =   14
      Top             =   7575
      Width           =   990
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7530
      Left            =   60
      TabIndex        =   32
      Top             =   15
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   13282
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmFacturaProveedor.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(2)=   "FrameFactura"
      Tab(0).Control(3)=   "FrameRemito"
      Tab(0).Control(4)=   "freCliente"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmFacturaProveedor.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameBuscar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame freCliente 
         Height          =   1775
         Left            =   -70920
         TabIndex        =   73
         Top             =   900
         Width           =   6950
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
            TabIndex        =   100
            Top             =   465
            Width           =   4620
         End
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
            TabIndex        =   93
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
            TabIndex        =   82
            Top             =   1080
            Width           =   4620
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
            TabIndex        =   81
            Top             =   780
            Width           =   4620
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
            TabIndex        =   80
            Top             =   1395
            Width           =   1455
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2385
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedor.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Agregar Cliente"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedor.frx":03C2
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Buscar Cliente"
            Top             =   120
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
            TabIndex        =   77
            Top             =   1395
            Width           =   3135
         End
         Begin VB.TextBox TxtCodigoCli 
            Enabled         =   0   'False
            Height          =   300
            Left            =   930
            MaxLength       =   40
            TabIndex        =   76
            Top             =   120
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
            TabIndex        =   75
            Tag             =   "Descripción"
            Top             =   120
            Width           =   3990
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
            Left            =   5610
            TabIndex        =   74
            Top             =   1395
            Width           =   1215
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   180
            TabIndex        =   101
            Top             =   510
            Width           =   675
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   150
            TabIndex        =   87
            Top             =   1125
            Width           =   705
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   795
            Width           =   735
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   85
            Top             =   165
            Width           =   780
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   255
            TabIndex        =   84
            Top             =   1440
            Width           =   600
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos"
            Height          =   195
            Left            =   5730
            TabIndex        =   83
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
         Height          =   575
         Left            =   -70920
         TabIndex        =   49
         Top             =   330
         Width           =   6945
         Begin VB.TextBox txtNroRemito 
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
            Left            =   1515
            TabIndex        =   96
            Top             =   225
            Width           =   1065
         End
         Begin VB.CommandButton cmdBuscarRemito 
            Height          =   315
            Left            =   2640
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedor.frx":06CC
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Buscar Remito"
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtRemSuc 
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
            Left            =   930
            MaxLength       =   4
            TabIndex        =   94
            Top             =   225
            Width           =   555
         End
         Begin VB.TextBox txtCodigoStock 
            Height          =   300
            Left            =   5760
            TabIndex        =   70
            Top             =   220
            Visible         =   0   'False
            Width           =   465
         End
         Begin MSComCtl2.DTPicker FechaRemito 
            Height          =   315
            Left            =   4080
            TabIndex        =   97
            Top             =   225
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   56623105
            CurrentDate     =   41098
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   240
            TabIndex        =   98
            Top             =   285
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   3555
            TabIndex        =   50
            Top             =   285
            Width           =   495
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
         Height          =   2250
         Left            =   405
         TabIndex        =   37
         Top             =   360
         Width           =   10410
         Begin VB.CheckBox chkProducto 
            Caption         =   "Producto"
            Height          =   195
            Left            =   210
            TabIndex        =   19
            Top             =   900
            Width           =   1035
         End
         Begin VB.TextBox txtdesProducto 
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
            TabIndex        =   103
            Top             =   855
            Width           =   4620
         End
         Begin VB.TextBox txtProducto 
            Height          =   300
            Left            =   3390
            TabIndex        =   20
            Top             =   840
            Width           =   990
         End
         Begin VB.CommandButton cmdProducto 
            Height          =   300
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedor.frx":09D6
            Style           =   1  'Graphical
            TabIndex        =   102
            ToolTipText     =   "Buscar Vendedor"
            Top             =   855
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.OptionButton optTodas 
            Caption         =   "Todas"
            Height          =   195
            Left            =   3960
            TabIndex        =   29
            Top             =   1920
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optMaq 
            Caption         =   "de Maquinarias"
            Height          =   195
            Left            =   2040
            TabIndex        =   28
            Top             =   1920
            Width           =   1455
         End
         Begin VB.OptionButton optRep 
            Caption         =   "de Repuestos"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CheckBox chkanuladas 
            Caption         =   "Ver Anuladas"
            Height          =   255
            Left            =   6600
            TabIndex        =   99
            Top             =   1560
            Width           =   1695
         End
         Begin VB.ComboBox cboFactura1 
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1485
            Width           =   2400
         End
         Begin VB.CheckBox chkTipoFactura 
            Caption         =   "Tipo de Facrura"
            Height          =   195
            Left            =   210
            TabIndex        =   25
            Top             =   1440
            Width           =   1485
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedor.frx":0CE0
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Buscar Cliente"
            Top             =   495
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1815
            Left            =   9690
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedor.frx":0FEA
            Style           =   1  'Graphical
            TabIndex        =   30
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
            TabIndex        =   38
            Tag             =   "Descripción"
            Top             =   495
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3390
            MaxLength       =   40
            TabIndex        =   18
            Top             =   495
            Width           =   975
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   210
            TabIndex        =   24
            Top             =   1170
            Width           =   810
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   210
            TabIndex        =   16
            Top             =   585
            Width           =   1095
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Empleado"
            Height          =   195
            Left            =   210
            TabIndex        =   17
            Top             =   645
            Visible         =   0   'False
            Width           =   1035
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
            TabIndex        =   42
            Top             =   600
            Visible         =   0   'False
            Width           =   4620
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3360
            TabIndex        =   23
            Top             =   585
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.CommandButton cmdBuscarVen 
            Height          =   300
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedor.frx":378C
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Buscar Vendedor"
            Top             =   600
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   405
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   3360
            TabIndex        =   21
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   56623105
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   6360
            TabIndex        =   22
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   56623105
            CurrentDate     =   41098
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Producto:"
            Height          =   195
            Index           =   2
            Left            =   2535
            TabIndex        =   104
            Top             =   840
            Width           =   690
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Factura:"
            Height          =   195
            Left            =   2325
            TabIndex        =   68
            Top             =   1530
            Width           =   945
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   5175
            TabIndex        =   41
            Top             =   1215
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2265
            TabIndex        =   40
            Top             =   1200
            Width           =   1005
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   2505
            TabIndex        =   39
            Top             =   540
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Empleado:"
            Height          =   195
            Index           =   0
            Left            =   2535
            TabIndex        =   43
            Top             =   615
            Visible         =   0   'False
            Width           =   750
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
         Height          =   2335
         Left            =   -74880
         TabIndex        =   34
         Top             =   330
         Width           =   3920
         Begin VB.CommandButton cmdNuevoVendedor 
            Height          =   315
            Left            =   2460
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedor.frx":3A96
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "Agregar Vendedor"
            Top             =   1560
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2040
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedor.frx":3E20
            Style           =   1  'Graphical
            TabIndex        =   89
            ToolTipText     =   "Buscar Vendedor"
            Top             =   1560
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
            Height          =   330
            Left            =   240
            TabIndex        =   88
            Top             =   1905
            Width           =   3165
         End
         Begin VB.TextBox txtNroVendedor 
            Height          =   300
            Left            =   1140
            TabIndex        =   4
            Top             =   1560
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
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   1
            Top             =   585
            Width           =   555
         End
         Begin VB.ComboBox cboFactura 
            Height          =   315
            Left            =   1200
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
            Left            =   1770
            MaxLength       =   8
            TabIndex        =   2
            Top             =   585
            Width           =   1065
         End
         Begin MSComCtl2.DTPicker FechaFactura 
            Height          =   315
            Left            =   1200
            TabIndex        =   3
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   56623105
            CurrentDate     =   41098
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Empleado:"
            Height          =   195
            Left            =   300
            TabIndex        =   91
            Top             =   1605
            Width           =   750
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   780
            TabIndex        =   54
            Top             =   255
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   645
            TabIndex        =   51
            Top             =   975
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   540
            TabIndex        =   48
            Top             =   615
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   600
            TabIndex        =   47
            Top             =   1290
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
            Left            =   1200
            TabIndex        =   46
            Top             =   1305
            Width           =   1350
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4620
         Left            =   360
         TabIndex        =   31
         Top             =   2700
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8149
         _Version        =   393216
         Cols            =   16
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame4 
         Height          =   540
         Left            =   -74895
         TabIndex        =   52
         Top             =   2615
         Width           =   10935
         Begin VB.CommandButton cmdNuevoRubro 
            Height          =   315
            Left            =   6195
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedor.frx":412A
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Agregar Condición de Venta"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboCondicion 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   165
            Width           =   4950
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "<F1> Buscar Facturas"
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
            Left            =   8400
            TabIndex        =   72
            Top             =   210
            Width           =   2250
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Condición:"
            Height          =   195
            Left            =   390
            TabIndex        =   65
            Top             =   210
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4350
         Left            =   -74895
         TabIndex        =   35
         Top             =   3075
         Width           =   10935
         Begin VB.CheckBox chkBonificaEnPesos 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en $"
            Height          =   285
            Left            =   390
            TabIndex        =   8
            Top             =   3600
            Width           =   1290
         End
         Begin VB.CheckBox chkBonificaEnPorsentaje 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en % "
            Height          =   285
            Left            =   390
            TabIndex        =   7
            Top             =   3300
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
            TabIndex        =   66
            Top             =   3630
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
            TabIndex        =   62
            Top             =   3630
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeIva 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6900
            TabIndex        =   11
            Top             =   3300
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
            TabIndex        =   59
            Top             =   3630
            Width           =   1035
         End
         Begin VB.TextBox txtPorcentajeBoni 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2850
            TabIndex        =   9
            Top             =   3300
            Width           =   1035
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
            TabIndex        =   56
            Top             =   3630
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
            TabIndex        =   55
            Top             =   3300
            Width           =   1350
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   1455
            MaxLength       =   60
            TabIndex        =   12
            Top             =   3975
            Width           =   8865
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   1140
            TabIndex        =   36
            Top             =   420
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3135
            Left            =   75
            TabIndex        =   6
            Top             =   120
            Width           =   10725
            _ExtentX        =   18918
            _ExtentY        =   5530
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
         Begin VB.TextBox txtIIBB 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4905
            TabIndex        =   10
            Top             =   3300
            Width           =   1155
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Percepcion:"
            Height          =   195
            Left            =   4005
            TabIndex        =   105
            Top             =   3330
            Width           =   855
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   4110
            TabIndex        =   67
            Top             =   3690
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   6270
            TabIndex        =   64
            Top             =   3675
            Width           =   570
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "% I.V.A.:"
            Height          =   195
            Left            =   6240
            TabIndex        =   63
            Top             =   3330
            Width           =   600
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   2235
            TabIndex        =   61
            Top             =   3675
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Bonificación:"
            Height          =   195
            Left            =   1890
            TabIndex        =   60
            Top             =   3330
            Width           =   915
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   8505
            TabIndex        =   58
            Top             =   3675
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   8175
            TabIndex        =   57
            Top             =   3330
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   210
            TabIndex        =   53
            Top             =   4020
            Width           =   1110
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   33
         Top             =   570
         Width           =   1065
      End
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
      TabIndex        =   92
      Top             =   7800
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
      TabIndex        =   45
      Top             =   7755
      Width           =   750
   End
End
Attribute VB_Name = "frmfacturaproveedor"
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
        txtProducto.Enabled = True
        cmdProducto.Enabled = True
    Else
        txtProducto.Enabled = False
        cmdProducto.Enabled = False
        txtProducto.Text = ""
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
        
        sql = "SELECT FC.*,"
        sql = sql & " C.PROV_CODIGO,C.PROV_RAZSOC, C.PROV_DOMICI, TC.TCO_ABREVIA,FC.FPR_IIBB"
        If txtProducto.Text <> "" Then
            sql = sql & ",DT.PTO_CODIGO"
        End If
        sql = sql & " FROM FACTURA_PROVEEDOR FC, REMITO_PROVEEDOR RC, PROVEEDOR C,"
        sql = sql & " TIPO_COMPROBANTE TC"
        If txtProducto.Text <> "" Then
            sql = sql & ", DETALLE_FACTURA_PROVEEDOR DT"
        End If
        sql = sql & " WHERE"
        sql = sql & " FC.RPR_NUMERO=RC.RPR_NUMERO"
        sql = sql & " AND FC.RPR_SUCURSAL=RC.RPR_SUCURSAL"
        sql = sql & " AND FC.RPR_FECHA=RC.RPR_FECHA"
        sql = sql & " AND FC.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND RC.PROV_CODIGO=C.PROV_CODIGO"
         If txtProducto.Text <> "" Then
            sql = sql & " AND FC.TCO_CODIGO = DT.TCO_CODIGO"
            sql = sql & " AND FC.FPR_NUMERO = DT.FPR_NUMERO"
            sql = sql & " AND FC.FPR_NROSUC = DT.FPR_NROSUC"
            sql = sql & " AND FC.FPR_FECHA = DT.FPR_FECHA"
            'sql = sql & " AND DT.PTO_CODIGO=P.PTO_CODIGO"
        End If
        
        If txtCliente.Text <> "" Then sql = sql & " AND RC.PROV_CODIGO=" & XN(txtCliente)
        If txtVendedor.Text <> "" Then sql = sql & " AND FC.VEN_CODIGO=" & XN(txtVendedor)
        If Not IsNull(FechaDesde) Then sql = sql & " AND FC.FPR_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND FC.FPR_FECHA<=" & XDQ(FechaHasta)
        If chkTipoFactura.Value = Checked Then sql = sql & " AND FC.TCO_CODIGO=" & XN(cboFactura1.ItemData(cboFactura1.ListIndex))
        If chkanuladas.Value = Unchecked Then
            sql = sql & " AND FC.EST_CODIGO <> 2"
        End If
        
        If optMaq.Value = True Then
            sql = sql & " AND FC.FPR_IVA = 10.5 "
        End If
        If optRep.Value = True Then
            sql = sql & " AND FC.FPR_IVA = 21 "
        End If
        'CAMBIOS
        If txtProducto.Text <> "" Then sql = sql & " AND DT.PTO_CODIGO LIKE '" & txtProducto.Text & "'"
        
        sql = sql & " ORDER BY FC.FPR_FECHA,FC.FPR_NROSUC,FC.FPR_NUMERO"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FPR_NROSUC, "0000") & "-" & Format(rec!FPR_NUMERO, "00000000") _
                                & Chr(9) & rec!FPR_FECHA & Chr(9) & rec!PROV_RAZSOC _
                                & Chr(9) & rec!PROV_DOMICI & Chr(9) & rec!VEN_CODIGO _
                                & Chr(9) & rec!EST_CODIGO & Chr(9) & Format(rec!RPR_SUCURSAL, "0000") & "-" & Format(rec!RPR_NUMERO, "00000000") _
                                & Chr(9) & rec!RPR_FECHA & Chr(9) & rec!FPR_BONIFICA _
                                & Chr(9) & rec!FPR_IVA & Chr(9) & rec!FPR_OBSERVACION _
                                & Chr(9) & rec!TCO_CODIGO & Chr(9) & rec!FPG_CODIGO _
                                & Chr(9) & rec!FPR_BONIPESOS & Chr(9) & rec!PROV_CODIGO & Chr(9) & _
                                IIf(IsNull(rec!FPR_IIBB), "0", rec!FPR_IIBB)
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
                        
        sql = "SELECT RC.RPR_NUMERO, RC.RPR_SUCURSAL, RC.RPR_FECHA, C.PROV_RAZSOC, C.PROV_DOMICI,C.PROV_CODIGO"
        sql = sql & " FROM REMITO_PROVEEDOR RC,PROVEEDOR C"
        sql = sql & " WHERE"
        sql = sql & " RC.PROV_CODIGO=C.PROV_CODIGO"
        sql = sql & " AND RC.EST_CODIGO = 1"
        'sql = sql & " AND NP.VEN_CODIGO=V.VEN_CODIGO"
        If txtCliente.Text <> "" Then sql = sql & " AND RC.PROV_CODIGO=" & XN(txtCliente)
        'If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
        If Not IsNull(FechaDesde) Then sql = sql & " AND RC.RPR_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND RC.RPR_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY RC.RPR_SUCURSAL,RC.RPR_NUMERO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem "" & Chr(9) & Format(rec!RPR_SUCURSAL, "0000") & "-" & Format(rec!RPR_NUMERO, "00000000") _
                                & Chr(9) & rec!RPR_FECHA & Chr(9) & rec!PROV_RAZSOC _
                                & Chr(9) & rec!PROV_DOMICI & Chr(9) & rec!PROV_CODIGO
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
    frmBuscar.TipoBusqueda = 5
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 1
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
    
    txtNroFactura_LostFocus
    
    DBConn.BeginTrans
    sql = "SELECT * FROM FACTURA_PROVEEDOR"
    sql = sql & " WHERE TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
    sql = sql & " AND FPR_NUMERO = " & XN(txtNroFactura)
    sql = sql & " AND FPR_NROSUC=" & XN(txtNroSucursal)
    'AGREGUE ESTA LINEA
    sql = sql & " AND PROV_CODIGO=" & XN(TxtCodigoCli)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = True Then
        'NUEVA FACTURA
        sql = "INSERT INTO FACTURA_PROVEEDOR"
        sql = sql & " (TCO_CODIGO,FPR_NUMERO,FPR_NROSUC,FPR_FECHA,"
        sql = sql & "RPR_NUMERO,RPR_SUCURSAL,RPR_FECHA,FPR_BONIFICA,FPR_IVA,FPG_CODIGO,FPR_OBSERVACION,"
        sql = sql & "FPR_BONIPESOS,FPR_NROSUCTXT,FPR_NUMEROTXT,FPR_SUBTOTAL,FPR_TOTAL,FPR_SALDO,EST_CODIGO,"
        sql = sql & "VEN_CODIGO,TPR_CODIGO,PROV_CODIGO,TGT_CODIGO,FPR_IIBB)"
        sql = sql & " VALUES ("
        sql = sql & cboFactura.ItemData(cboFactura.ListIndex) & ","
        sql = sql & XN(txtNroFactura) & ","
        sql = sql & XN(txtNroSucursal) & ","
        sql = sql & XDQ(FechaFactura) & ","
        sql = sql & XN(txtNroRemito) & ","
        sql = sql & XN(txtRemSuc) & ","
        sql = sql & XDQ(FechaRemito) & ","
        sql = sql & XN(txtPorcentajeBoni) & ","
        sql = sql & XN(txtPorcentajeIva) & ","
        sql = sql & cboCondicion.ItemData(cboCondicion.ListIndex) & ","
        sql = sql & XS(txtObservaciones) & ","
        If chkBonificaEnPesos.Value = Checked Then
            sql = sql & "'S'" & "," 'BONIFICA EN PESOS
        ElseIf chkBonificaEnPorsentaje.Value = Checked Then
            sql = sql & "'N'" & "," 'BONIFICA EN PORCENTAJE
        Else
            sql = sql & "NULL" & "," 'NO HAY BONIFICACION
        End If
        sql = sql & XS(Format(txtNroSucursal.Text, "0000")) & ","
        sql = sql & XS(Format(txtNroFactura.Text, "00000000")) & ","
        If txtSubTotalBoni.Text <> "" Then 'SUBTOTAL
            sql = sql & XN(txtSubTotalBoni) & ","
        Else
            sql = sql & XN(txtSubtotal) & ","
        End If
        sql = sql & XN(txtTotal) & ","
        sql = sql & XN(txtTotal) & "," 'SALDO FACTURA
        sql = sql & "3" & "," 'ESTADO DEFINITIVO
        sql = sql & XN(txtNroVendedor) & ","
        sql = sql & XN(cmdBuscarCliente.Tag) & ","
        sql = sql & XN(TxtCodigoCli) & ","
        sql = sql & 5 & ","
        sql = sql & XN(txtIIBB.Text) & ")"
        DBConn.Execute sql
           
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                sql = "INSERT INTO DETALLE_FACTURA_PROVEEDOR"
                sql = sql & " (TCO_CODIGO,FPR_NUMERO,FPR_NROSUC,FPR_FECHA,"
                sql = sql & "DFP_NROITEM,PTO_CODIGO,DFP_CANTIDAD,DFP_PRECIO,DFP_BONIFICA,DFP_DETALLE"
                'AGRUEGUE ACA
                sql = sql & ",PROV_CODIGO)"
                sql = sql & " VALUES ("
                sql = sql & cboFactura.ItemData(cboFactura.ListIndex) & ","
                sql = sql & XN(txtNroFactura) & ","
                sql = sql & XN(txtNroSucursal) & ","
                sql = sql & XDQ(FechaFactura) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 9)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(I, 0), True) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 4)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(I, 1)) & ","
                'AGREGUE ACA
                sql = sql & XN(TxtCodigoCli) & ")"
                DBConn.Execute sql
            End If
        Next
         
        'CAMBIO ESTADO DEL REMITO (LE PONGO DEFINITIVO)
        sql = "UPDATE REMITO_PROVEEDOR SET EST_CODIGO=3"
        sql = sql & " WHERE"
        sql = sql & " RPR_NUMERO=" & XN(txtNroRemito)
        sql = sql & " AND RPR_SUCURSAL=" & XN(txtRemSuc)
        DBConn.Execute sql
        
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO A LA FACTURA QUE CORRESPONDE
'        sql = "SELECT * FROM PARAMETROS"
'        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If Rec1.EOF = False Then
'            'If Rec1!REP_CODIGO = cboRep.ItemData(cboRep.ListIndex) Then
'                Select Case cboFactura.ItemData(cboFactura.ListIndex)
'                    Case 1
'                        sql = "UPDATE PARAMETROS SET FACTURA_A=" & XN(txtNroFactura)
'                    Case 2
'                        sql = "UPDATE PARAMETROS SET FACTURA_B=" & XN(txtNroFactura)
'                End Select
'                    DBConn.Execute sql
'            'End If
'        End If
'        Rec1.Close
        
        'ACTUALIZO LA CUENTA CORRIENTE DEL PROVEEDOR
        'DBConn.Execute AgregoCtaCteCliente(TxtCodigoCli, CStr(cboFactura.ItemData(cboFactura.ListIndex)) _
                                            , txtNroFactura, txtNroSucursal, _
                                            FechaFactura, txtTotal, "D", CStr(Date))
        
        DBConn.Execute AgregoCtaCteProveedores(cmdBuscarCliente.Tag _
                                              , TxtCodigoCli, CStr(cboFactura.ItemData(cboFactura.ListIndex)) _
                                        , txtNroSucursal, txtNroFactura, FechaFactura, txtTotal, "D", CStr(Date))
        DBConn.CommitTrans
    Else
        MsgBox "La Factura ya fue Registrada", vbCritical, TIT_MSGBOX
        txtNroFactura.SetFocus
        Screen.MousePointer = vbNormal
        lblEstado.Caption = ""
        rec.Close
        DBConn.CommitTrans
        Exit Sub
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
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
    If txtNroSucursal.Text = "" Then
        MsgBox "El numero de Sucursal es requerido", vbExclamation, TIT_MSGBOX
        txtNroSucursal.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If txtNroFactura.Text = "" Then
        MsgBox "El numero de Factura es requerido", vbExclamation, TIT_MSGBOX
        txtNroFactura.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If FechaFactura.Value = "" Then
        MsgBox "La Fecha de la Factura es requerida", vbExclamation, TIT_MSGBOX
        FechaFactura.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If txtNroRemito.Text = "" Then
        MsgBox "El numero del Remito es requerido", vbExclamation, TIT_MSGBOX
        txtNroRemito.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If FechaRemito.Value = "" Then
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
    
    ValidarFactura = True
End Function
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
   CmdBuscAprox.Tag = ""
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
   txtIIBB.Text = ""
   txtObservaciones.Text = ""
   cboCondicion.ListIndex = 0
   lblEstado.Caption = ""
   cmdGrabar.Enabled = True
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
    TipoBusquedaDoc = 1
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
        txtProducto.Text = frmBuscar.grdBuscar.Text
        txtProducto.SetFocus
        txtProducto_LostFocus
    Else
        txtProducto.SetFocus
    End If
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmfacturaproveedor = Nothing
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

Private Sub Command2_Click()
    
End Sub

Private Sub Form_Activate()
    If txtNroRemito.Text <> "" Then
        txtNroSucursal_LostFocus
        txtNroFactura_LostFocus
        txtNroVendedor_LostFocus
        txtNroRemito_LostFocus
        'txtNroFactura.SetFocus
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
    grdGrilla.Cols = 10
    grdGrilla.Rows = 1
    For I = 2 To 25
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                             & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & (I - 1)
    Next
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "Tipo Fac|^Número|^Fecha|Proveedor|Domicilio|Empleado|Cod_Estado|" _
                              & "REMITO_NUMERO|REMITO_FECHA|PORCENTAJE BONIFICA|PORCENTAJE IVA|" _
                              & "OBSERVACIONES|COD TIPO COMPROBANTE|COD CONDICION VENTA|" _
                              & "BONIFICA EN PESOS|CLI_CODIGO|IIBB"
    GrdModulos.ColWidth(0) = 900  'TIPO FACTURA
    GrdModulos.ColWidth(1) = 1500 'NUMERO
    GrdModulos.ColWidth(2) = 1100 'FECHA
    GrdModulos.ColWidth(3) = 4000 'PROVEEDOR
    GrdModulos.ColWidth(4) = 4000 'Domicilio
    GrdModulos.ColWidth(5) = 0    'VENDEDOR
    GrdModulos.ColWidth(6) = 0    'COD_ESTADO
    GrdModulos.ColWidth(7) = 0    'REMITO_NUMERO
    GrdModulos.ColWidth(8) = 0    'REMITO_FECHA
    GrdModulos.ColWidth(9) = 0    'PORCENTAJE BONIFICA
    GrdModulos.ColWidth(10) = 0   'PORCENTAJE IVA
    GrdModulos.ColWidth(11) = 0   'OBSERVACIONES
    GrdModulos.ColWidth(12) = 0   'COD TIPO COMPROBANTE
    GrdModulos.ColWidth(13) = 0   'COD CONDICION VENTA
    GrdModulos.ColWidth(14) = 0   'BONIFICA EN PESOS
    GrdModulos.ColWidth(15) = 0   'REPRESENTADA
    GrdModulos.ColWidth(16) = 0   'IIBB
    GrdModulos.Rows = 1
    '------------------------------------
   
    '------------------------------------
    lblEstado.Caption = ""
    'CARGO COMBO CON LOS TIPOS DE FACTURA
    LlenarComboFactura
    'CARGO COMBO CON LAS CONDICIONES DE VENTA
    LlenarComboFormaPago
    
    'Llenotabla DETALLE FACTURA PROVEEDOR
'    sql = " SELECT * FROM FACTURA_PROVEEDOR"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            sql = "UPDATE DETALLE_FACTURA_PROVEEDOR "
'            sql = sql & "SET PROV_CODIGO = " & XN(rec!PROV_CODIGO)
'            sql = sql & " WHERE TCO_CODIGO = " & rec!TCO_CODIGO & ""
'            sql = sql & " AND FPR_NUMERO = " & rec!FPR_NUMERO
'            sql = sql & " and FPR_NROSUC = " & rec!FPR_NROSUC
'            DBConn.Execute sql
'            rec.MoveNext
'        Loop
'    End If
'    rec.Close
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoFactura) 'ESTADO PENDIENTE
    VEstadoFactura = 1
    FechaFactura.Value = Date
    TipoBusquedaDoc = 1 'ESTO ES PARA BUSCAR FACTURA(1), (2)PARA BUSCAR REMITOS
    tabDatos.Tab = 0
    'BUSCO IVA
    BuscoIva
'    If reciboFac = 1 Then
'        tabDatos.Tab = 1
'        CmdBuscAprox_Click
'    End If
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

'Private Function BuscoUltimaFactura(TipoFac As Integer) As String
'    'ACA BUSCA EL NUMERO DE REMITO SIGUIENTE AL ULTIMO CARGADO
'    sql = "SELECT (FACTURA_A) + 1 AS FAC_A, (FACTURA_B) + 1 AS FAC_B, SUCURSAL"
'    sql = sql & " FROM PARAMETROS"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        txtNroSucursal.Text = Format(rec!Sucursal, "0000")
'        Select Case TipoFac
'            Case 1
'                BuscoUltimaFactura = IIf(IsNull(rec!FAC_A), 1, rec!FAC_A)
'            Case 2
'                BuscoUltimaFactura = IIf(IsNull(rec!FAC_B), 1, rec!FAC_B)
'            Case 3
'                MsgBox "No hay Facturas del tipo C", vbExclamation, TIT_MSGBOX
'                cboFactura.SetFocus
'        End Select
'    End If
'    rec.Close
'End Function

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
            Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 12)), cboFactura)
            txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
            txtNroFactura.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
            FechaFactura.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
            Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6)), lblEstadoFactura)
            VEstadoFactura = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6))
            
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 11) <> "" Then
                txtObservaciones.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 11))
            End If
            'CABEZA REMITO
            txtRemSuc.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 7), 4)
            txtNroRemito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 7), 8)
            FechaRemito.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
            
            
            TxtCodigoCli.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 15)
            TxtCodigoCli_LostFocus
            
            txtNroVendedor.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
            txtNroVendedor_LostFocus
            
            'IIBB
            txtIIBB.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 16))
            
            'CONDICION VENTA
            Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 13)), cboCondicion)
            '----BUSCO DETALLE DE LA FACTURA------------------
            sql = "SELECT DFP.*, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM DETALLE_FACTURA_PROVEEDOR DFP, PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE DFP.FPR_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8))
            sql = sql & " AND DFP.FPR_NROSUC=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
            sql = sql & " AND DFP.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 12))
            'agregue aca
            sql = sql & " AND DFP.PROV_CODIGO=" & XN(TxtCodigoCli.Text)
            sql = sql & " AND DFP.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " ORDER BY DFP.DFP_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(I, 1) = IIf(IsNull(Rec1!DFP_DETALLE), Rec1!PTO_DESCRI, Rec1!DFP_DETALLE)
                    grdGrilla.TextMatrix(I, 2) = Rec1!DFP_CANTIDAD
                    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DFP_PRECIO)
                    If IsNull(Rec1!DFP_BONIFICA) Then
                        grdGrilla.TextMatrix(I, 4) = ""
                    Else
                        grdGrilla.TextMatrix(I, 4) = Valido_Importe(Rec1!DFP_BONIFICA)
                    End If
                    VBonificacion = 0
                    If Not IsNull(Rec1!DFP_BONIFICA) Then
                        VBonificacion = (((CDbl(Rec1!DFP_CANTIDAD) * CDbl(Rec1!DFP_PRECIO)) * CDbl(Rec1!DFP_BONIFICA)) / 100)
                        VBonificacion = ((CDbl(Rec1!DFP_CANTIDAD) * CDbl(Rec1!DFP_PRECIO)) - VBonificacion)
                        grdGrilla.TextMatrix(I, 5) = Valido_Importe(CStr(VBonificacion))
                        grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
                    Else
                        VBonificacion = (CDbl(Rec1!DFP_CANTIDAD) * CDbl(Rec1!DFP_PRECIO))
                        grdGrilla.TextMatrix(I, 5) = ""
                        grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
                    End If
                    grdGrilla.TextMatrix(I, 7) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 8) = Rec1!LNA_DESCRI
                    grdGrilla.TextMatrix(I, 9) = Rec1!DFP_NROITEM
                    I = I + 1
                    Rec1.MoveNext
                Loop
                VBonificacion = 0
            End If
            Rec1.Close
            '--CARGO LOS TOTALES----
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            txtTotal.Text = txtSubtotal.Text
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 14) = "S" Then
                chkBonificaEnPesos.Value = Checked
            ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 14) = "N" Then
                chkBonificaEnPorsentaje.Value = Checked
            Else
                chkBonificaEnPesos.Value = Unchecked
                chkBonificaEnPorsentaje.Value = Unchecked
            End If
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 9) <> "" Then
                txtPorcentajeBoni.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 9)
                txtPorcentajeBoni_LostFocus
            End If
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 10) <> "" Then
                txtPorcentajeIva = GrdModulos.TextMatrix(GrdModulos.RowSel, 10)
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
        '----------------------------------------------------------
        Case 2 'BUSCA REMITO
        
            lblEstado.Caption = "Buscando..."
            Screen.MousePointer = vbHourglass
            
            txtRemSuc.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
            txtNroRemito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
            FechaRemito.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
            
            'grillaRemito.TextMatrix(0, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 3)
            'grillaRemito.TextMatrix(1, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 4)
            TxtCodigoCli.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
        
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
    Else
        frameBuscar.Caption = "Buscar Remitos Pendientes por..."
        chkTipoFactura.Enabled = False
        cboFactura1.Enabled = False
        chkVendedor.Enabled = False
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



Private Sub Text1_Change()

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
        sql = "SELECT PROV_RAZSOC FROM PROVEEDOR"
        sql = sql & " WHERE PROV_CODIGO=" & XN(txtCliente)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!PROV_RAZSOC
        Else
            MsgBox "El PROVEEDOR no existe", vbExclamation, TIT_MSGBOX
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
        txtprovincia.Text = ""
        txtcodpos.Text = ""
    End If
End Sub

Private Sub TxtCodigoCli_LostFocus()
Dim Rec1 As New ADODB.Recordset
If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Or ActiveControl.Name = "CmdSalir" Then Exit Sub
    If TxtCodigoCli.Text <> "" Then
        sql = "SELECT C.PROV_RAZSOC,C.PROV_DOMICI,C.PROV_CUIT,C.IVA_CODIGO,C.PROV_INGBRU,C.TPR_CODIGO, "
        sql = sql & "L.LOC_DESCRI,P.PRO_DESCRI,L.LOC_CODPOS,C.TPR_CODIGO"
        sql = sql & " FROM PROVEEDOR C, LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE "
        sql = sql & "C.LOC_CODIGO = L.LOC_CODIGO AND "
        sql = sql & "C.PRO_CODIGO = P.PRO_CODIGO AND "
        sql = sql & "L.PRO_CODIGO = P.PRO_CODIGO AND "
        sql = sql & "C.PROV_CODIGO =" & XN(TxtCodigoCli)
        'sql = sql & " AND PROV_ESTADO=1"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtRazSocCli.Text = Rec1!PROV_RAZSOC
            txtDomici.Text = Rec1!PROV_DOMICI
            txtlocalidad.Text = Rec1!LOC_DESCRI
            txtprovincia.Text = Rec1!PRO_DESCRI
            txtCondicionIVA.Text = BuscoCondicionIVA(Rec1!IVA_CODIGO)
            txtCUIT.Text = IIf(IsNull(Rec1!PROV_CUIT), "NO INFORMADO", Format(Rec1!PROV_CUIT, "##-########-#"))
            txtIngBrutos.Text = IIf(IsNull(Rec1!PROV_INGBRU), "NO INFORMADO", Format(Rec1!PROV_INGBRU, "###-#####-##"))
            txtcodpos.Text = IIf(IsNull(Rec1!LOC_CODPOS), "", Rec1!LOC_CODPOS)
            cmdBuscarCliente.Tag = Rec1!TPR_CODIGO 'GUARDO EL CODIGO DEL TIPO DE PROVEEDOR
        Else
            MsgBox "El PROVEEDOR no existe", vbExclamation, TIT_MSGBOX
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
    FechaRemito.Value = ""
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
        sql = "SELECT PROV_RAZSOC FROM PROVEEDOR"
        sql = sql & " WHERE PROV_CODIGO=" & XN(Codigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            BuscoCliente = rec!PROV_RAZSOC
        Else
            BuscoCliente = "No se encontro el PROVEEDOR"
        End If
        rec.Close
End Function

Private Function BuscoSucursal(CodigoSuc As String, CodigoCli As String) As String
        sql = "SELECT * FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(CodigoSuc)
        sql = sql & " AND PROV_CODIGO=" & XN(CodigoCli)
        
        Set Rec1 = New ADODB.Recordset
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            BuscoSucursal = Rec1!SUC_DESCRI
        Else
            BuscoSucursal = "No se encontro la Sucursal"
        End If
        Rec1.Close
End Function

Private Sub txtIIBB_GotFocus()
    SelecTexto txtIIBB
End Sub

Private Sub txtIIBB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtIIBB, KeyAscii)
End Sub

Private Sub txtIIBB_LostFocus()
'    Dim Percepcion As Double
'    Dim total As Double
'    txtIIBB.Text = Valido_Importe(txtIIBB.Text)
'    Percepcion = Valido_Importe(txtIIBB.Text)
'    total = Valido_Importe(txtTotal.Text)
'    txtTotal = Percepcion + total
'    txtTotal = Format(txtTotal, "#.##")
    txtPorcentajeIva_LostFocus
End Sub

Private Sub txtNroFactura_GotFocus()
    SelecTexto txtNroFactura
End Sub

Private Sub txtNroFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroFactura_LostFocus()
    If txtNroFactura.Text = "" Then
        'BUSCO EL NUMERO DE FACTURA QUE CORRESPONDE
        'txtNroFactura.Text = Format(BuscoUltimaFactura(cboFactura.ItemData(cboFactura.ListIndex)), "00000000")
    Else
        txtNroFactura.Text = Format(txtNroFactura.Text, "00000000")
    End If
End Sub

Private Sub txtNroRemito_LostFocus()
    If txtNroRemito.Text <> "" Then
        txtNroRemito.Text = Format(txtNroRemito.Text, "00000000")
        sql = "SELECT RC.RPR_NUMERO, RC.RPR_SUCURSAL, RC.RPR_FECHA, RC.EST_CODIGO, RC.STK_CODIGO, E.EST_DESCRI"
        sql = sql & " ,RC.PROV_CODIGO, RC.VEN_CODIGO"
        sql = sql & " FROM REMITO_PROVEEDOR RC,ESTADO_DOCUMENTO E,PROVEEDOR C"
        sql = sql & " WHERE RC.RPR_NUMERO=" & XN(txtNroRemito)
        sql = sql & " AND RC.RPR_SUCURSAL=" & XN(txtRemSuc.Text)
        sql = sql & " AND RC.PROV_CODIGO= " & XN(TxtCodigoCli.Text)
        sql = sql & " AND RC.PROV_CODIGO=C.PROV_CODIGO"
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
            FechaRemito.Value = Rec2!RPR_FECHA
            TxtCodigoCli.Text = Rec2!PROV_CODIGO
            TxtCodigoCli_LostFocus
            'grillaRemito.TextMatrix(0, 1) = BuscoCliente(Rec2!PROV_CODIGO)
            'grillaRemito.TextMatrix(1, 1) = BuscoSucursal(Rec2!SUC_CODIGO, Rec2!PROV_CODIGO)
            'grillaRemito.TextMatrix(2, 1) = BuscoVendedor(Rec2!VEN_CODIGO)
            'grillaRemito.TextMatrix(0, 2) = Rec2!PROV_CODIGO 'guardo el codigo del PROVEEDOR
            txtCodigoStock.Text = Rec2!STK_CODIGO
            
            If Rec2!EST_CODIGO <> 1 Then
                MsgBox "El Remito número: " & txtNroRemito.Text & Chr(13) & Chr(13) & _
                       "No puede ser asignado a la Factura por su estado (" & Rec2!EST_DESCRI & ")", vbExclamation, TIT_MSGBOX
                cmdGrabar.Enabled = False
                Screen.MousePointer = vbNormal
                lblEstado.Caption = ""
                Rec2.Close
                LimpiarRemito
                'txtRemSuc.SetFocus
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
            sql = "SELECT DRPR.*,P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM DETALLE_REMITO_PROVEEDOR DRPR, PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE DRPR.RPR_NUMERO=" & XN(txtNroRemito)
            sql = sql & " AND DRPR.RPR_SUCURSAL=" & XN(txtRemSuc)
            sql = sql & " AND DRPR.PROV_CODIGO=" & XN(TxtCodigoCli)
            sql = sql & " AND DRPR.RPR_FECHA=" & XDQ(FechaRemito)
            sql = sql & " AND DRPR.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " ORDER BY DRPR.DRPR_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(I, 1) = IIf(IsNull(Rec1!DRPR_DETALLE), Rec1!PTO_DESCRI, Rec1!DRPR_DETALLE)
                    grdGrilla.TextMatrix(I, 2) = IIf(IsNull(Rec1!DRPR_CANTIDAD), 0, Rec1!DRPR_CANTIDAD)
                    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DRPR_PRECIO)
                    grdGrilla.TextMatrix(I, 4) = ""
                    grdGrilla.TextMatrix(I, 5) = ""
                    grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(CDbl(IIf(IsNull(Rec1!DRPR_CANTIDAD), 0, Rec1!DRPR_CANTIDAD) * CDbl(Rec1!DRPR_PRECIO))))
                    grdGrilla.TextMatrix(I, 7) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 8) = Rec1!LNA_DESCRI
                    grdGrilla.TextMatrix(I, 9) = Rec1!DRPR_NROITEM
                    I = I + 1
                    Rec1.MoveNext
                Loop
                txtSubtotal.Text = Valido_Importe(SumaTotal)
                txtTotal.Text = txtSubtotal.Text
            End If
            Rec1.Close
            If grdGrilla.TextMatrix(1, 8) = "MAQUINARIA" Then 'pregunta si la linea es Maquinaria
                txtPorcentajeIva.Text = "10,50"
            Else
                txtPorcentajeIva.Text = "21,00"
            End If
            '--------------------------------------------------
            cboCondicion.SetFocus
            txtPorcentajeIva_LostFocus
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
        Else
            MsgBox "El Remito no existe", vbExclamation, TIT_MSGBOX
            If Rec2.State = 1 Then Rec2.Close
            LimpiarRemito
            'txtNroRemito.SetFocus
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
 If txtProducto.Text = "" Then
        txtdesProducto.Text = ""
    End If
End Sub

Private Sub txtProducto_GotFocus()
    SelecTexto txtProducto
End Sub

Private Sub txtProducto_LostFocus()
     If txtProducto.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT PTO_DESCRI FROM PRODUCTO"
        sql = sql & " WHERE PTO_CODIGO LIKE '" & txtProducto & "'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtdesProducto.Text = rec!PTO_DESCRI
        Else
            MsgBox "El Producto no existe", vbExclamation, TIT_MSGBOX
            txtdesProducto.Text = ""
            txtProducto.SetFocus
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
            txtTotal.Text = CDbl(txtSubTotalBoni.Text) + CDbl(txtImporteIva.Text) + CDbl(Chk0(txtIIBB.Text))
            txtTotal.Text = Valido_Importe(txtTotal.Text)
        Else
            txtImporteIva.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
            txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
            txtTotal.Text = CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text) + CDbl(IIf(txtIIBB.Text = "", 0, txtIIBB))
            txtTotal.Text = Valido_Importe(txtTotal.Text)
        End If
    End If
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
