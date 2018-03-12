VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmOrdenCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Compra"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7965
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   555
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7335
      Width           =   990
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      Height          =   555
      Left            =   9135
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7335
      Width           =   990
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   555
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7335
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   555
      Left            =   7125
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7335
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   555
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7335
      Width           =   990
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7155
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   12621
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
      TabPicture(0)   =   "frmOrdenCompra.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FramePedido"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmOrdenCompra.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Height          =   4470
         Left            =   105
         TabIndex        =   55
         Top             =   2505
         Width           =   10920
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   270
            TabIndex        =   5
            Top             =   495
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.CommandButton cmdQuitarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmOrdenCompra.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   58
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Producto"
            Top             =   885
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdAgregarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmOrdenCompra.frx":0DBA
            Style           =   1  'Graphical
            TabIndex        =   57
            TabStop         =   0   'False
            ToolTipText     =   "Agregar Producto"
            Top             =   540
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdBuscarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmOrdenCompra.frx":10C4
            Style           =   1  'Graphical
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "Buscar Producto"
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   4005
            Left            =   120
            TabIndex        =   4
            Top             =   285
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   7064
            _Version        =   393216
            Rows            =   3
            Cols            =   6
            FixedCols       =   0
            BackColorSel    =   12648447
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483633
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FramePedido 
         Caption         =   "Compra..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   105
         TabIndex        =   46
         Top             =   420
         Width           =   3360
         Begin VB.TextBox txtNroNotaPedido 
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
            Left            =   1020
            MaxLength       =   8
            TabIndex        =   0
            Top             =   315
            Width           =   1155
         End
         Begin VB.TextBox txtNroVendedor 
            Height          =   300
            Left            =   1020
            TabIndex        =   2
            Top             =   1005
            Width           =   780
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
            Left            =   120
            TabIndex        =   49
            Top             =   1350
            Width           =   3165
         End
         Begin VB.CommandButton cmdBuscarVendedor 
            Height          =   315
            Left            =   1845
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenCompra.frx":13CE
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Buscar Vendedor"
            Top             =   1005
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoVendedor 
            Height          =   315
            Left            =   2265
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenCompra.frx":16D8
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Agregar Vendedor"
            Top             =   1005
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin FechaCtl.Fecha FechaNotaPedido 
            Height          =   285
            Left            =   1020
            TabIndex        =   1
            Top             =   675
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   360
            TabIndex        =   54
            Top             =   345
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   465
            TabIndex        =   53
            Top             =   690
            Width           =   495
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Empleado:"
            Height          =   195
            Left            =   225
            TabIndex        =   52
            Top             =   1050
            Width           =   750
         End
         Begin VB.Label lblEstadoNota 
            AutoSize        =   -1  'True
            Caption         =   "EST. ORDEN COMPRA"
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
            Left            =   1020
            TabIndex        =   51
            Top             =   1800
            Width           =   1995
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   420
            TabIndex        =   50
            Top             =   1785
            Width           =   540
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   " Datos del Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   3480
         TabIndex        =   25
         Top             =   420
         Width           =   7545
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
            Left            =   6000
            TabIndex        =   39
            Top             =   1680
            Width           =   1455
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
            Left            =   960
            MaxLength       =   50
            TabIndex        =   38
            Top             =   615
            Width           =   4620
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
            Left            =   2865
            MaxLength       =   50
            TabIndex        =   37
            Tag             =   "Descripción"
            Top             =   270
            Width           =   4590
         End
         Begin VB.TextBox TxtCodigoCli 
            Height          =   300
            Left            =   960
            MaxLength       =   40
            TabIndex        =   3
            Top             =   270
            Width           =   975
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
            Left            =   2445
            TabIndex        =   36
            Top             =   1665
            Width           =   3135
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   1980
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenCompra.frx":1A62
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Buscar Cliente"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Height          =   315
            Left            =   2415
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenCompra.frx":1D6C
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Agregar Cliente"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   405
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
            Left            =   960
            TabIndex        =   33
            Top             =   1665
            Width           =   1455
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
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   32
            Top             =   960
            Width           =   5100
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
            Left            =   960
            MaxLength       =   50
            TabIndex        =   29
            Top             =   1320
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
            Left            =   960
            TabIndex        =   27
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos"
            Height          =   195
            Left            =   6000
            TabIndex        =   45
            Top             =   1455
            Width           =   765
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   285
            TabIndex        =   44
            Top             =   1710
            Width           =   600
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   210
            TabIndex        =   43
            Top             =   645
            Width           =   675
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
            Left            =   120
            TabIndex        =   42
            Top             =   315
            Width           =   780
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   150
            TabIndex        =   41
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   180
            TabIndex        =   40
            Top             =   1350
            Width           =   705
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   1710
         Left            =   -74625
         TabIndex        =   11
         Top             =   570
         Width           =   10395
         Begin VB.CheckBox chkCliente 
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   300
            TabIndex        =   16
            Top             =   435
            Width           =   1215
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   300
            TabIndex        =   20
            Top             =   1215
            Width           =   810
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3360
            MaxLength       =   40
            TabIndex        =   22
            Top             =   375
            Width           =   975
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
            Left            =   4845
            MaxLength       =   50
            TabIndex        =   15
            Tag             =   "Descripción"
            Top             =   375
            Width           =   4620
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1380
            Left            =   9705
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenCompra.frx":20F6
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Buscar "
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Empleado"
            Height          =   195
            Left            =   300
            TabIndex        =   18
            Top             =   825
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
            TabIndex        =   14
            Top             =   855
            Width           =   4635
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3360
            TabIndex        =   24
            Top             =   840
            Width           =   990
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   300
            Left            =   4395
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenCompra.frx":4898
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Buscar Cliente"
            Top             =   375
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarVen 
            Height          =   300
            Left            =   4395
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenCompra.frx":4BA2
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Buscar Vendedor"
            Top             =   840
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   5865
            TabIndex        =   28
            Top             =   1320
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   3360
            TabIndex        =   26
            Top             =   1320
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
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
            TabIndex        =   23
            Top             =   420
            Width           =   780
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2265
            TabIndex        =   21
            Top             =   1365
            Width           =   1005
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4815
            TabIndex        =   19
            Top             =   1380
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Empleado:"
            Height          =   195
            Index           =   0
            Left            =   2535
            TabIndex        =   17
            Top             =   885
            Width           =   750
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4560
         Left            =   -74640
         TabIndex        =   31
         Top             =   2430
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8043
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   59
         Top             =   570
         Width           =   1065
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   1560
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Height          =   240
      Left            =   210
      TabIndex        =   61
      Top             =   7680
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Orden de Compra"
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
      Left            =   2160
      TabIndex        =   60
      Top             =   7440
      Width           =   3180
   End
End
Attribute VB_Name = "frmOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer



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
Private Sub chkVendedor_Click()
    If chkVendedor.Value = Checked Then
        txtVendedor.Enabled = True
        cmdBuscarVen.Enabled = True
    Else
        txtVendedor.Enabled = False
        cmdBuscarVen.Enabled = False
    End If
End Sub

Private Sub cmdAgregarProducto_Click()
    ABMProducto.Show vbModal
    grdGrilla.SetFocus
    grdGrilla.row = 1
End Sub

Private Sub CmdBorrar_Click()
    If txtNroNotaPedido.Text <> "" Then
        If MsgBox("Seguro que desea eliminar la Orden de Compra Nro.: " & Trim(txtNroNotaPedido), vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
           On Error GoTo Seclavose
           
           sql = "SELECT P.EST_CODIGO, E.EST_DESCRI "
           sql = sql & " FROM ORDEN_COMPRA P, ESTADO_DOCUMENTO E"
           sql = sql & " WHERE OC_NUMERO=" & XN(txtNroNotaPedido)
           sql = sql & " AND OC_FECHA=" & XDQ(FechaNotaPedido)
           sql = sql & " AND P.EST_CODIGO=E.EST_CODIGO"
           rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
           
           If rec.EOF = False Then
                If rec!EST_CODIGO <> 1 Then
                    MsgBox "La Orden de Compra no puede ser eliminada," & Chr(13) & _
                           " ya que esta en estado: " & Trim(rec!EST_DESCRI), vbExclamation, TIT_MSGBOX
                    rec.Close
                    Exit Sub
                End If
           End If
           rec.Close
            lblEstado.Caption = "Eliminando..."
            Screen.MousePointer = vbHourglass
            
            sql = "DELETE FROM DETALLE_ORDEN_COMPRA"
            sql = sql & " WHERE OC_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND OC_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
            
            sql = "DELETE FROM ORDEN_COMPRA"
            sql = sql & " WHERE OC_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND OC_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
            
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            CmdNuevo_Click
        End If
    End If
    Exit Sub
    
Seclavose:
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT NP.*, C.PROV_RAZSOC,C.PROV_DOMICI,L.LOC_DESCRI"
    sql = sql & " FROM ORDEN_COMPRA NP, PROVEEDOR C, LOCALIDAD L "
    sql = sql & " WHERE"
    sql = sql & " NP.PROV_CODIGO=C.PROV_CODIGO AND"
    sql = sql & " C.LOC_CODIGO=L.LOC_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND NP.PROV_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND NP.OC_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND NP.OC_FECHA<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY NP.OC_NUMERO, NP.OC_FECHA"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!OC_NUMERO & Chr(9) & rec!OC_FECHA _
                            & Chr(9) & rec!PROV_RAZSOC & Chr(9) & rec!PROV_DOMICI _
                            & Chr(9) & rec!LOC_DESCRI
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
    End If
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
    frmBuscar.TipoBusqueda = 5
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

Private Sub cmdBuscarProducto_Click()
    grdGrilla.SetFocus
    frmBuscar.TipoBusqueda = 2
    frmBuscar.CodListaPrecio = 0
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    
    grdGrilla.Col = 0
    EDITAR grdGrilla, txtEdit, 13
    If Trim(frmBuscar.grdBuscar.Text) <> "" Then txtEdit.Text = frmBuscar.grdBuscar.Text
    TxtEdit_KeyDown vbKeyReturn, 0
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

Private Sub cmdBuscarVendedor_Click()
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

Private Sub cmdGrabar_Click()
    If ValidarNotaPedido = False Then Exit Sub
    If MsgBox("¿Confirma Orden de Compra?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    On Error GoTo HayErrorNota
    
    DBConn.BeginTrans
    sql = "SELECT * FROM ORDEN_COMPRA"
    sql = sql & " WHERE OC_NUMERO=" & XN(txtNroNotaPedido)
    sql = sql & " AND OC_FECHA=" & XDQ(FechaNotaPedido)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = False Then
        If MsgBox("Seguro que modificar la Nota de Pedido Nro.: " & Trim(txtNroNotaPedido), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            sql = "UPDATE ORDEN_COMPRA"
            sql = sql & " SET PROV_CODIGO=" & XN(TxtCodigoCli)
            sql = sql & " ,VEN_CODIGO=" & XN(txtNroVendedor)
            
            sql = sql & " WHERE"
            sql = sql & " OC_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND OC_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
            
            sql = "DELETE FROM DETALLE_ORDEN_COMPRA"
            sql = sql & " WHERE OC_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND OC_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
            
            For I = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(I, 0) <> "" Then
                    sql = "INSERT INTO DETALLE_ORDEN_COMPRA"
                    sql = sql & " (OC_NUMERO,OC_FECHA,DOC_NROITEM,PTO_CODIGO,"
                    sql = sql & "DOC_CANTIDAD)"
                    sql = sql & " VALUES ("
                    sql = sql & XN(txtNroNotaPedido) & ","
                    sql = sql & XDQ(FechaNotaPedido) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 5)) & "," 'NRO ITEM
                    sql = sql & XN(grdGrilla.TextMatrix(I, 0)) & "," 'PRODUCTO CODIGO
                    sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ")" 'CANTIDAD
                    'sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ")" 'PRECIO
                    DBConn.Execute sql
                End If
            Next
            DBConn.CommitTrans
        End If
        
    Else 'ORDEN DE COMPRA NUEVA
        sql = "INSERT INTO ORDEN_COMPRA "
        sql = sql & " (OC_NUMERO,OC_FECHA,TPR_CODIGO,"
        sql = sql & "PROV_CODIGO,VEN_CODIGO,EST_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(txtNroNotaPedido) & ","
        sql = sql & XDQ(FechaNotaPedido) & ","
        sql = sql & "1" & ","
        sql = sql & XN(TxtCodigoCli) & ","
        sql = sql & XN(txtNroVendedor) & ","
        'If chkDetalle.Value = Checked Then
        '    sql = sql & cboCondicion.ItemData(cboCondicion.ListIndex) & ","
        'Else
        '    sql = sql & "NULL,"
        'End If
        'sql = sql & XS(Format(txtNroNotaPedido.Text, "00000000")) & ","
        sql = sql & "1)" 'ESTADO PENDIENTE
        DBConn.Execute sql
           
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                sql = "INSERT INTO DETALLE_ORDEN_COMPRA "
                sql = sql & " (OC_NUMERO,OC_FECHA,DOC_NROITEM,PTO_CODIGO,"
                sql = sql & " DOC_CANTIDAD)"
                sql = sql & " VALUES ("
                sql = sql & XN(txtNroNotaPedido) & ","
                sql = sql & XDQ(FechaNotaPedido) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 5)) & "," 'NRO ITEM
                sql = sql & XN(grdGrilla.TextMatrix(I, 0)) & "," 'PRODUCTO CODIGO
                sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ")" 'CANTIDAD
                'sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ")" 'PRECIO
                DBConn.Execute sql
            End If
        Next
        DBConn.CommitTrans
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    CmdNuevo_Click
    Exit Sub
    
HayErrorNota:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Function ValidarNotaPedido() As Boolean
    Dim I As Integer
    If txtNroNotaPedido.Text = "" Then
        MsgBox "El número de Orden de Compra es requerido", vbExclamation, TIT_MSGBOX
        txtNroNotaPedido.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    If FechaNotaPedido.Text = "" Then
        MsgBox "La Fecha de la Orden de Compra es requerida", vbExclamation, TIT_MSGBOX
        FechaNotaPedido.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    'NO ES OBLIGATORIO
    'If txtNroVendedor.Text = "" Then
    '    MsgBox "El Vendedor es requerido", vbExclamation, TIT_MSGBOX
    '    txtNroVendedor.SetFocus
    '    ValidarNotaPedido = False
    '    Exit Function
    'End If
    If TxtCodigoCli.Text = "" Then
        MsgBox "El Proveedor es requerido", vbExclamation, TIT_MSGBOX
        TxtCodigoCli.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    'NO ES OBLIGATORIO
    'If txtCodigoSuc.Text = "" Then
    '    MsgBox "El Cliente es requerido", vbExclamation, TIT_MSGBOX
    '    TxtCodigoCli.SetFocus
    '    ValidarNotaPedido = False
    '    Exit Function
    'End If
    
    ValidarNotaPedido = True
End Function

Private Sub cmdImprimir_Click()
    'Rep.WindowState = crptMaximized 'crptMinimized
    If MsgBox("¿Confirma La Impresión de la Orden de Compra?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    sql = "UPDATE ORDEN_COMPRA"
    sql = sql & " SET EST_CODIGO =" & 3
    sql = sql & " WHERE"
    sql = sql & " OC_NUMERO=" & XN(txtNroNotaPedido)
    sql = sql & " AND OC_FECHA=" & XDQ(FechaNotaPedido)
    DBConn.Execute sql
    
    
    'ORDEN DE COMPRA DETALLADO

    Rep.Formulas(0) = ""
    If txtNroNotaPedido.Text = "" Then
        MsgBox "Debe seleccionar una Orden de Compra", vbExclamation, TIT_MSGBOX
        txtNroNotaPedido.SetFocus
    Exit Sub
    End If
    Rep.SelectionFormula = ""
    Rep.SelectionFormula = "{ORDEN_COMPRA.OC_NUMERO}=" & txtNroNotaPedido.Text _
                            & " AND {ORDEN_COMPRA.OC_FECHA}= DATE (" & Mid(FechaNotaPedido, 7, 4) & "," & Mid(FechaNotaPedido, 4, 2) & "," & Mid(FechaNotaPedido, 1, 2) & ")"
    'DATE (" & Mid(FechaNotaPedido, 7, 4) & "," & Mid(FechaNotaPedido, 4, 2) & "," & Mid(FechaNotaPedido, 1, 2) & ")"
    Rep.WindowTitle = "Orden de Compra"
    Rep.ReportFileName = DRIVE & DirReport & "rptordencompra.rpt"



     Rep.Destination = crptToWindow
     Rep.WindowState = crptMaximized
     Rep.Action = 1


     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     Rep.Formulas(2) = ""
     CmdNuevo_Click
     
End Sub

Private Sub CmdNuevo_Click()
   For I = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(I, 0) = ""
        grdGrilla.TextMatrix(I, 1) = ""
        grdGrilla.TextMatrix(I, 2) = ""
        grdGrilla.TextMatrix(I, 3) = ""
        grdGrilla.TextMatrix(I, 4) = ""
        grdGrilla.TextMatrix(I, 5) = ""
        
   Next
   FramePedido.Enabled = True
   fraDatos.Enabled = True
   TxtCodigoCli.Text = ""
   TxtCodigoCli.Text = ""
   txtRazSocCli.Text = ""
   txtNombreVendedor.Text = ""
   txtNroVendedor.Text = ""
   FechaNotaPedido.Text = ""
   txtNroNotaPedido.Text = ""
   lblEstadoNota.Caption = ""
   lblEstado.Caption = ""
   tabDatos.Tab = 0
   Call BuscoEstado(1, lblEstadoNota)
   cmdGrabar.Enabled = True
   CmdBorrar.Enabled = True
   cmdImprimir.Enabled = False
   txtNroNotaPedido.SetFocus
End Sub

Private Sub cmdNuevoCliente_Click()
    ABMCliente.Show vbModal
    TxtCodigoCli.SetFocus
End Sub
Private Sub cmdNuevoVendedor_Click()
    ABMVendedor.Show vbModal
    txtNroVendedor.SetFocus
End Sub

Private Sub cmdQuitarProducto_Click()
    If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
        If MsgBox("Seguro que desea quitar el Producto: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 1), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = ""
            'grdGrilla.RowSel
            grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = ""
        End If
    Else
        MsgBox "Debe seleccionar un Producto", vbExclamation, TIT_MSGBOX
        grdGrilla.SetFocus
        grdGrilla.Col = 0
        grdGrilla.row = 1
    End If
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmNotaDePedido = Nothing
        Unload Me
    End If
End Sub



Private Sub FechaNotaPedido_Change()
    If (FechaNotaPedido = "") And txtNroNotaPedido.Text <> "" Then
        FechaNotaPedido.Text = Date
    Else
        If txtNroNotaPedido.Text = "" Then txtNroNotaPedido.SetFocus
        grdGrilla.SetFocus
    End If
End Sub

Private Sub FechaNotaPedido_LostFocus()
    If FechaNotaPedido.Text = "" Then
        FechaNotaPedido.Text = Date
        If txtNroNotaPedido.Text = "" Then txtNroNotaPedido.SetFocus
    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
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
    grdGrilla.FormatString = "Código|Descripción|Cantidad|Rubro|Linea|Orden"
    grdGrilla.ColWidth(0) = 1000 'CODIGO
    grdGrilla.ColWidth(1) = 5900 'DESCRIPCION
    grdGrilla.ColWidth(2) = 1000 'CANTIDAD
    grdGrilla.ColWidth(3) = 2100 'RUBRO
    grdGrilla.ColWidth(4) = 2100 'LINEA
    grdGrilla.ColWidth(5) = 0    'ORDEN
    grdGrilla.Cols = 6
    grdGrilla.Rows = 1
    For I = 2 To 14
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & (I - 1)
    Next
    
'    grdGrilla.FormatString = "Código|Descripción|Cantidad|Orden|Rubro|Linea"
'    grdGrilla.ColWidth(0) = 1000 'CODIGO
'    grdGrilla.ColWidth(1) = 5900 'DESCRIPCION
'    grdGrilla.ColWidth(2) = 1000 'CANTIDAD
'    grdGrilla.ColWidth(3) = 0    'ORDEN
'    grdGrilla.ColWidth(4) = 2100 'RUBRO
'    grdGrilla.ColWidth(5) = 2100 'LINEA
'    grdGrilla.Cols = 6
'    grdGrilla.Rows = 1
'    For I = 2 To 14
'        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & (I - 1) & Chr(9) & "" & Chr(9) & ""
'    Next
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
'    GrdModulos.FormatString = ">Número|^Fecha|Cliente|Sucursal"
    GrdModulos.FormatString = ">Número|^Fecha|Proveedor|Domicilio|Localidad"
    GrdModulos.ColWidth(0) = 1200
    GrdModulos.ColWidth(1) = 1000
    GrdModulos.ColWidth(2) = 3200
    GrdModulos.ColWidth(3) = 3200
    GrdModulos.ColWidth(4) = 3200
    GrdModulos.Rows = 1
    '------------------------------------
    
    '-----------------
    'cboListaPrecio.ListIndex = 0
    'cboCondicion.ListIndex = -1
    'cboCondicion.Enabled = False
'    cmdNuevoRubro.Enabled = False
    cmdImprimir.Enabled = False
    lblEstado.Caption = ""
    Call BuscoEstado(1, lblEstadoNota)
    tabDatos.Tab = 0
End Sub
Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1
            LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
            grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = grdGrilla.RowSel
            grdGrilla.Col = 0
        'Case Else
        '    grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, grdGrilla.Col)) = ""
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
        Case 1
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = "" Then
                cmdGrabar.SetFocus
            End If
        Case 2
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "" Then
                grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "1"
            End If
        End Select
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
       (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Then
        If KeyAscii = vbKeyReturn Then
            If grdGrilla.Col = 3 Then
                If grdGrilla.row < grdGrilla.Rows - 1 Then
                    grdGrilla.row = grdGrilla.row + 1
                    grdGrilla.Col = 0
                Else
                    SendKeys "{TAB}"
                End If
            Else
                grdGrilla.Col = grdGrilla.Col + 1
            End If
        Else
            If (grdGrilla.Col <> 1) Then
                'If KeyAscii > 47 And KeyAscii < 58 Then
                    EDITAR grdGrilla, txtEdit, KeyAscii
                'End If
            Else
                EDITAR grdGrilla, txtEdit, KeyAscii
            End If
        End If
    End If
End Sub

Private Sub grdGrilla_LeaveCell()
    If txtEdit.Visible = False Then Exit Sub
    'If Trim(TxtEdit) = "" Then TxtEdit = "0"
    grdGrilla = txtEdit.Text
    txtEdit.Visible = False
End Sub

Private Sub grdGrilla_GotFocus()
    If grdGrilla.Rows > 1 Then
        If txtEdit.Visible = False Then Exit Sub
        grdGrilla = txtEdit.Text
        txtEdit.Visible = False
    End If
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        CmdNuevo_Click
        txtNroNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        FechaNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
        cmdImprimir.Enabled = True
        tabDatos.Tab = 0
        txtNroNotaPedido_LostFocus
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    txtCliente.Enabled = False
    'txtSucursal.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    txtVendedor.Enabled = False
    cmdGrabar.Enabled = False
    CmdBorrar.Enabled = False
    cmdBuscarCli.Enabled = False
    cmdBuscarVen.Enabled = False
    LimpiarBusqueda
    If Me.Visible = True Then chkCliente.SetFocus
  Else
    'If Me.Visible = True Then txtNroNotaPedido.SetFocus
    cmdGrabar.Enabled = True
    CmdBorrar.Enabled = True
  End If
End Sub

Private Sub LimpiarBusqueda()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    txtVendedor.Text = ""
    txtDesVen.Text = ""
    GrdModulos.Rows = 1
    chkCliente.Value = Unchecked
    chkFecha.Value = Unchecked
    chkVendedor.Value = Unchecked
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
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtDesCli.Text = ""
            txtCliente.SetFocus
            Exit Sub
        End If
        rec.Close
    End If
'    If chkSucursal.Value = Unchecked And chkFecha.Value = Unchecked _
'        And chkVendedor.Value = Unchecked And ActiveControl.Name <> "cmdBuscarCli" _
'        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub

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

Private Sub txtCodigoCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Function BuscoCondicionIVA(IVACodigo As String) As String
    Set Rec3 = New ADODB.Recordset
    sql = "SELECT * FROM CONDICION_IVA"
    sql = sql & " WHERE IVA_CODIGO=" & XN(IVACodigo)
    Rec3.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec3.EOF = False Then
        BuscoCondicionIVA = Rec3!IVA_DESCRI
    Else
        BuscoCondicionIVA = ""
    End If
    Rec3.Close
    Set Rec3 = Nothing
End Function
Private Sub TxtCodigoCli_GotFocus()
    SelecTexto TxtCodigoCli
End Sub

Private Sub TxtCodigoCli_LostFocus()
    'If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Or ActiveControl.Name = "CmdSalir" Then Exit Sub
    If TxtCodigoCli.Text <> "" Then
        sql = "SELECT C.prov_RAZSOC,C.prov_DOMICI,C.prov_CUIT,C.IVA_CODIGO,C.prov_INGBRU,"
        sql = sql & "L.LOC_DESCRI,P.PRO_DESCRI,L.LOC_CODPOS"
        sql = sql & " FROM PROVEEDOR C, LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE "
        sql = sql & "C.LOC_CODIGO = L.LOC_CODIGO AND "
        sql = sql & "C.PRO_CODIGO = P.PRO_CODIGO AND "
        sql = sql & "L.PRO_CODIGO = P.PRO_CODIGO AND "
        sql = sql & "C.PROV_CODIGO=" & XN(TxtCodigoCli)
        'sql = sql & " AND CLI_ESTADO=1"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtRazSocCli.Text = rec!PROV_RAZSOC
            txtDomici.Text = rec!PROV_DOMICI
            txtlocalidad.Text = rec!LOC_DESCRI
            txtprovincia.Text = rec!PRO_DESCRI
            txtCondicionIVA.Text = BuscoCondicionIVA(rec!IVA_CODIGO)
            txtCUIT.Text = IIf(IsNull(rec!PROV_CUIT), "NO INFORMADO", Format(rec!PROV_CUIT, "##-########-#"))
            txtIngBrutos.Text = IIf(IsNull(rec!PROV_INGBRU), "NO INFORMADO", Format(rec!PROV_INGBRU, "###-#####-##"))
            txtcodpos.Text = IIf(IsNull(rec!LOC_CODPOS), "", rec!LOC_CODPOS)
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtRazSocCli.Text = ""
            TxtCodigoCli.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    If grdGrilla.Col = 0 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 2 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 3 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    CarTexto KeyAscii
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 Then
        frmBuscar.TipoBusqueda = 2
        'frmBuscar.CodListaPrecio = 0
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        frmBuscar.Show vbModal
    End If

    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
        Case 0, 1
            If Trim(txtEdit) = "" Then
                txtEdit = ""
                LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                grdGrilla.Col = 0
                grdGrilla.SetFocus
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE"
            If grdGrilla.Col = 0 Then
                sql = sql & " P.PTO_CODIGO LIKE '" & txtEdit & "'"
            Else
                sql = sql & " P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
            End If
                
                sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
                sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
                sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                sql = sql & " AND PTO_ESTADO=1"
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                If rec.RecordCount > 1 Then
                    grdGrilla.SetFocus
                    frmBuscar.TipoBusqueda = 2
                    'LE DIGO EN QUE LISTA DE PRECIO BUSCAR LOS PRECIOS
                    'frmBuscar.CodListaPrecio = 0
                    frmBuscar.TxtDescriB.Text = txtEdit.Text
                    frmBuscar.Show vbModal
                    grdGrilla.Col = 0
                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
                    grdGrilla.Col = 1
                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
                    grdGrilla.Col = 3
                    'grdGrilla.Text = Valido_Importe(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2))
                    'grdGrilla.Col = 4
                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
                    grdGrilla.Col = 4
                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = grdGrilla.RowSel
                    grdGrilla.Col = 2
                Else
                    grdGrilla.Col = 0
                    grdGrilla.Text = Trim(rec!PTO_CODIGO)
                    grdGrilla.Col = 1
                    grdGrilla.Text = Trim(rec!PTO_DESCRI)
                    'grdGrilla.Col = 3
                    'grdGrilla.Text = Valido_Importe(Trim(rec!LIS_PRECIO))
                    grdGrilla.Col = 3
                    grdGrilla.Text = Trim(rec!RUB_DESCRI)
                    grdGrilla.Col = 4
                    grdGrilla.Text = Trim(rec!LNA_DESCRI)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = grdGrilla.RowSel
                    grdGrilla.Col = 2
                End If
                    If BuscoRepetetidos(CLng(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), grdGrilla.RowSel) = False Then
                     grdGrilla.Col = 0
                     grdGrilla_KeyDown vbKeyDelete, 0
                    End If
            Else
                    MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
                    txtEdit.Text = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
            End If
            rec.Close
            Screen.MousePointer = vbNormal
        Case 2
            If Trim(txtEdit) = "" Then grdGrilla.Text = "1"
        Case 3
            If Trim(txtEdit) <> "" Then
                txtEdit.Text = Valido_Importe(txtEdit)
            Else
                MsgBox "Debe ingresar el Importe", vbExclamation, TIT_MSGBOX
                grdGrilla.Col = 3
            End If
        End Select
        grdGrilla.SetFocus
    End If
    If KeyCode = vbKeyEscape Then
       txtEdit.Visible = False
       grdGrilla.SetFocus
    End If
'    If KeyCode = vbKeyF1 Then
'        frmBuscar.TipoBusqueda = 2
'        frmBuscar.CodListaPrecio = 0
'        grdGrilla.Col = 0
'        EDITAR grdGrilla, txtEdit, 13
'        frmBuscar.Show vbModal
'    End If
'
'    If KeyCode = vbKeyReturn Then
'        Select Case grdGrilla.Col
'        Case 0, 1
'            If Trim(txtEdit) = "" Then
'                txtEdit = ""
'                LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
'                grdGrilla.Col = 0
'                grdGrilla.SetFocus
'                Exit Sub
'            End If
'            Screen.MousePointer = vbHourglass
'            sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
'            sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L"
'            sql = sql & " WHERE"
'            If grdGrilla.Col = 0 Then
'                sql = sql & " PTO_CODIGO=" & XN(txtEdit)
'            Else
'                sql = sql & " PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
'            End If
'                sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
'                sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
'                sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
'
'            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'            If rec.EOF = False Then
'                If rec.RecordCount > 1 Then
'                    grdGrilla.SetFocus
'                    frmBuscar.TipoBusqueda = 2
'                    frmBuscar.CodListaPrecio = 0
'                    frmBuscar.TxtDescriB.Text = txtEdit.Text
'                    frmBuscar.Show vbModal
'                    grdGrilla.Col = 0
'                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
'                    grdGrilla.Col = 1
'                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
'                    grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = grdGrilla.RowSel
'                    grdGrilla.Col = 4
'                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
'                    grdGrilla.Col = 5
'                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)
'                    grdGrilla.Col = 2
'                Else
'                    grdGrilla.Col = 0
'                    grdGrilla.Text = Trim(rec!PTO_CODIGO)
'                    grdGrilla.Col = 1
'                    grdGrilla.Text = Trim(rec!PTO_DESCRI)
'                    grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = grdGrilla.RowSel
'                    grdGrilla.Col = 4
'                    grdGrilla.Text = Trim(rec!RUB_DESCRI)
'                    grdGrilla.Col = 5
'                    grdGrilla.Text = Trim(rec!LNA_DESCRI)
'                    grdGrilla.Col = 2
'                End If
'                    If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
'                        If BuscoRepetetidos(CLng(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), grdGrilla.RowSel) = False Then
'                         grdGrilla.Col = 0
'                         grdGrilla_KeyDown vbKeyDelete, 0
'                        End If
'                    End If
'            Else
'                    MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
'                    txtEdit.Text = ""
'                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
'                    grdGrilla.Col = 0
'            End If
'            rec.Close
'            Screen.MousePointer = vbNormal
'        Case 2
'            If Trim(txtEdit) = "" Then grdGrilla.Text = "1"
'        End Select
'        grdGrilla.SetFocus
'    End If
'    If KeyCode = vbKeyEscape Then
'       txtEdit.Visible = False
'       grdGrilla.SetFocus
'    End If
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

Private Sub txtNroNotaPedido_Change()
    If txtNroNotaPedido.Text = "" Then
        FechaNotaPedido.Text = ""
        txtNroNotaPedido.SetFocus
    End If
End Sub

Private Sub txtNroNotaPedido_GotFocus()
     FechaNotaPedido.Text = ""
End Sub

Private Sub txtNroNotaPedido_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaPedido_LostFocus()
    If ActiveControl.Name = "CmdSalir" Or ActiveControl.Name = "chkCliente" _
       Or ActiveControl.Name = "cmdNuevo" Or ActiveControl.Name = "cmdBuscarNotaPedido" Then Exit Sub
    
    If txtNroNotaPedido.Text <> "" Then
        sql = "SELECT O.*, E.EST_DESCRI"
        sql = sql & " FROM ORDEN_COMPRA O, ESTADO_DOCUMENTO E"
        sql = sql & " WHERE O.OC_NUMERO=" & XN(txtNroNotaPedido)
        If FechaNotaPedido.Text <> "" Then
            sql = sql & " AND O.OC_FECHA=" & XDQ(FechaNotaPedido)
        End If
        sql = sql & " AND O.EST_CODIGO=E.EST_CODIGO"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec2.EOF = False Then
            If Rec2.RecordCount > 1 Then
                MsgBox "Hay mas de una Orden de Compra con el Número: " & txtNroNotaPedido.Text, vbInformation, TIT_MSGBOX
                tabDatos.Tab = 1
                Rec2.Close
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Buscando..."
            
            'CARGO CABECERA DE LA NOTA DE PEDIDO
            FechaNotaPedido.Text = Rec2!OC_FECHA
            txtNroVendedor.Text = Rec2!VEN_CODIGO
            'BUSCA FORMA DE PAGO
            'If Not IsNull(Rec2!FPG_CODIGO) Then
            '    chkDetalle.Value = Checked
            '    Call BuscaCodigoProxItemData(Rec2!FPG_CODIGO, cboCondicion)
            'Else
            '    chkDetalle.Value = Unchecked
            'End If


            
            txtNroVendedor_LostFocus
            TxtCodigoCli.Text = Rec2!PROV_CODIGO
            TxtCodigoCli_LostFocus
            
            Call BuscoEstado(Rec2!EST_CODIGO, lblEstadoNota)
            If Rec2!EST_CODIGO <> 1 Then
                cmdGrabar.Enabled = False
                CmdBorrar.Enabled = False
                FramePedido.Enabled = False
                fraDatos.Enabled = False
                grdGrilla.SetFocus
                cmdImprimir.Enabled = False
            Else
                cmdGrabar.Enabled = True
                CmdBorrar.Enabled = True
                FramePedido.Enabled = True
                fraDatos.Enabled = True
            End If
            
            'BUSCO LOS DATOS DEL DETALLE DE LA ORDEN DE COMPRA
            sql = "SELECT DO.*,P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM DETALLE_ORDEN_COMPRA DO, PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE DO.OC_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND DO.OC_FECHA=" & XDQ(FechaNotaPedido)
            sql = sql & " AND DO.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " ORDER BY DO.DOC_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(I, 1) = Rec1!PTO_DESCRI
                    grdGrilla.TextMatrix(I, 2) = IIf(IsNull(Rec1!DOC_CANTIDAD), "", Rec1!DOC_CANTIDAD)
                    'If IsNull(Rec1!DNP_PRECIO) Then
                    '    grdGrilla.TextMatrix(I, 3) = ""
                    'Else
                    '    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DNP_PRECIO)
                    'End If
                    grdGrilla.TextMatrix(I, 3) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 4) = Rec1!LNA_DESCRI
                    grdGrilla.TextMatrix(I, 5) = Rec1!DOC_NROITEM
                    I = I + 1
                    Rec1.MoveNext
                Loop
            End If
            Rec1.Close
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
        Else
            Call BuscoEstado(1, lblEstadoNota)
        End If
        Rec2.Close
    Else
        MsgBox "Debe ingresar el Número de la Orden de Compra", vbExclamation, TIT_MSGBOX
        txtNroNotaPedido.SetFocus
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
        If TxtCodigoCli.Enabled = True Then TxtCodigoCli.SetFocus
    Else
        MsgBox "El Vendedor no existe", vbExclamation, TIT_MSGBOX
        txtNombreVendedor.Text = ""
        txtNroVendedor.SetFocus
    End If
    rec.Close
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
            Exit Sub
        End If
        rec.Close
    End If
'    If chkFecha.Value = Unchecked And ActiveControl.Name <> "cmdNuevo" _
'        And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub


