VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmRemitoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remito de Clientes..."
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7845
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   8080
      TabIndex        =   8
      Top             =   7380
      Width           =   990
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10080
      TabIndex        =   10
      Top             =   7380
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7380
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   9080
      TabIndex        =   9
      Top             =   7380
      Width           =   990
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7300
      Left            =   45
      TabIndex        =   23
      Top             =   45
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   12885
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
      TabPicture(0)   =   "frmRemitoCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "freRemito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "freNotaPedido"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "freCliente"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkNotaPedido"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmRemitoCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox chkNotaPedido 
         Caption         =   "Recupera datos de Nota de Pedido"
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.Frame freCliente 
         Height          =   1815
         Left            =   4050
         TabIndex        =   50
         Top             =   1000
         Width           =   7000
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
            TabIndex        =   68
            Top             =   780
            Width           =   1215
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
            TabIndex        =   67
            Top             =   780
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
            Left            =   5610
            TabIndex        =   58
            Top             =   1395
            Width           =   1215
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
            TabIndex        =   57
            Top             =   465
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
            Left            =   2835
            MaxLength       =   50
            TabIndex        =   56
            Tag             =   "Descripción"
            Top             =   120
            Width           =   3990
         End
         Begin VB.TextBox TxtCodigoCli 
            Height          =   300
            Left            =   930
            MaxLength       =   40
            TabIndex        =   5
            Top             =   120
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
            Left            =   2415
            TabIndex        =   55
            Top             =   1395
            Width           =   3135
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   1920
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Buscar Cliente"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Height          =   315
            Left            =   2385
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Agregar Cliente"
            Top             =   120
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
            Left            =   930
            TabIndex        =   52
            Top             =   1395
            Width           =   1455
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
            TabIndex        =   51
            Top             =   1080
            Width           =   4620
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos"
            Height          =   195
            Left            =   5730
            TabIndex        =   64
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   255
            TabIndex        =   63
            Top             =   1440
            Width           =   600
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   180
            TabIndex        =   62
            Top             =   480
            Width           =   675
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
            TabIndex        =   61
            Top             =   165
            Width           =   525
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   795
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   150
            TabIndex        =   59
            Top             =   1125
            Width           =   705
         End
      End
      Begin VB.Frame freNotaPedido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   4050
         TabIndex        =   42
         Top             =   360
         Width           =   6990
         Begin VB.CommandButton cmdBuscarNotaPedido 
            Height          =   315
            Left            =   2685
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":06CC
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Buscar Nota de Pedido"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtNroNotaPedido 
            Height          =   300
            Left            =   1470
            TabIndex        =   11
            Top             =   270
            Width           =   1155
         End
         Begin FechaCtl.Fecha FechaNotaPedido 
            Height          =   285
            Left            =   4470
            TabIndex        =   12
            Top             =   255
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   3855
            TabIndex        =   45
            Top             =   255
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   750
            TabIndex        =   43
            Top             =   285
            Width           =   600
         End
      End
      Begin VB.Frame freRemito 
         Caption         =   "Remito..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   3920
         Begin TabDlg.SSTab tabLista 
            Height          =   1215
            Left            =   120
            TabIndex        =   74
            Top             =   1200
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   2143
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "Maquinarias"
            TabPicture(0)   =   "frmRemitoCliente.frx":09D6
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame2"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Repuestos"
            TabPicture(1)   =   "frmRemitoCliente.frx":09F2
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame4"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame4 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   -74880
               TabIndex        =   77
               Top             =   360
               Width           =   3495
               Begin VB.ComboBox cboLPrecioRep 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   78
                  Top             =   240
                  Width           =   3225
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   120
               TabIndex        =   75
               Top             =   360
               Width           =   3495
               Begin VB.ComboBox cboListaPrecio 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   76
                  Top             =   240
                  Width           =   3225
               End
            End
         End
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
            Height          =   330
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   1
            Top             =   240
            Width           =   1005
         End
         Begin VB.ComboBox cboStock 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   2505
         End
         Begin VB.TextBox txtNroSucursal 
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
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   0
            Top             =   240
            Width           =   555
         End
         Begin FechaCtl.Fecha FechaRemito 
            Height          =   285
            Left            =   1320
            TabIndex        =   2
            Top             =   585
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label lblEstadoRemito 
            AutoSize        =   -1  'True
            Caption         =   "EST. REMITO"
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
            Left            =   1320
            TabIndex        =   65
            Top             =   930
            Width           =   1215
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Stock:"
            Height          =   210
            Left            =   675
            TabIndex        =   48
            Top             =   120
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   765
            TabIndex        =   46
            Top             =   585
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   660
            TabIndex        =   41
            Top             =   285
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   720
            TabIndex        =   40
            Top             =   915
            Width           =   540
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
         Height          =   1950
         Left            =   -74600
         TabIndex        =   31
         Top             =   540
         Width           =   10410
         Begin VB.Frame Frame1 
            Caption         =   "Estado Remito"
            Height          =   495
            Left            =   840
            TabIndex        =   69
            Top             =   1320
            Width           =   8535
            Begin VB.OptionButton optTod 
               Caption         =   "Todos"
               Height          =   195
               Left            =   6600
               TabIndex        =   73
               Top             =   200
               Width           =   1455
            End
            Begin VB.OptionButton optAnu 
               Caption         =   "Anulados"
               Height          =   195
               Left            =   4845
               TabIndex        =   72
               Top             =   200
               Width           =   1455
            End
            Begin VB.OptionButton optDef 
               Caption         =   "Definitivos"
               Height          =   195
               Left            =   3075
               TabIndex        =   71
               Top             =   200
               Width           =   1455
            End
            Begin VB.OptionButton optPen 
               Caption         =   "Pendientes"
               Height          =   195
               Left            =   1200
               TabIndex        =   70
               Top             =   200
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.CommandButton cmdBuscarVendedor 
            Height          =   315
            Left            =   4290
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":0A0E
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Buscar Vendedor"
            Top             =   660
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4290
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":0D18
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Buscar Cliente"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3240
            TabIndex        =   18
            Top             =   667
            Width           =   990
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
            Left            =   4725
            TabIndex        =   36
            Top             =   675
            Width           =   4620
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   900
            TabIndex        =   15
            Top             =   645
            Width           =   1035
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1395
            Left            =   9660
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":1022
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Buscar "
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   5745
            TabIndex        =   20
            Top             =   1080
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   3240
            TabIndex        =   19
            Top             =   1080
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
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
            Left            =   4725
            MaxLength       =   50
            TabIndex        =   32
            Tag             =   "Descripción"
            Top             =   255
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3240
            MaxLength       =   40
            TabIndex        =   17
            Top             =   255
            Width           =   975
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   900
            TabIndex        =   16
            Top             =   975
            Width           =   810
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   878
            TabIndex        =   14
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Index           =   0
            Left            =   2415
            TabIndex        =   37
            Top             =   712
            Width           =   735
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4695
            TabIndex        =   35
            Top             =   1140
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2145
            TabIndex        =   34
            Top             =   1125
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
            Left            =   2625
            TabIndex        =   33
            Top             =   300
            Width           =   525
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4170
         Left            =   -74640
         TabIndex        =   22
         Top             =   2715
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7355
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame3 
         Height          =   4500
         Left            =   120
         TabIndex        =   26
         Top             =   2730
         Width           =   10935
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   240
            TabIndex        =   27
            Top             =   525
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   1455
            MaxLength       =   60
            TabIndex        =   13
            Top             =   4110
            Width           =   8850
         End
         Begin VB.CommandButton cmdBuscarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoCliente.frx":37C4
            Style           =   1  'Graphical
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Buscar Producto"
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdAgregarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoCliente.frx":3ACE
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Agregar Producto"
            Top             =   540
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdQuitarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoCliente.frx":3DD8
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Producto"
            Top             =   885
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3855
            Left            =   90
            TabIndex        =   6
            Top             =   165
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   6800
            _Version        =   393216
            Rows            =   3
            Cols            =   8
            FixedCols       =   0
            BackColorSel    =   12648447
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   1
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   210
            TabIndex        =   47
            Top             =   4155
            Width           =   1110
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   24
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Remito"
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
      Left            =   3600
      TabIndex        =   66
      Top             =   7440
      Width           =   2085
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
      Height          =   255
      Left            =   150
      TabIndex        =   39
      Top             =   7455
      Width           =   750
   End
End
Attribute VB_Name = "frmRemitoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim w As Integer
Dim TipoBusquedaDoc As Integer
Dim VEstadoRemito As Integer
Dim VCantidadBultos As Integer
Dim Rec1 As ADODB.Recordset
Public nlista As Integer
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

Private Sub chkRemitoSinFactura_Click()
   ' If chkRemitoSinFactura.Value = Checked Then
   '     txtConcepto.Enabled = True
   ' Else
   '     txtConcepto.Enabled = False
   ' End If
End Sub

Private Sub chkNotaPedido_Click()
    If chkNotaPedido.Value = 1 Then
        freNotaPedido.Enabled = True
        txtNroNotaPedido.Enabled = True
        txtNroNotaPedido.SetFocus
        freCliente.Enabled = False
    Else
        freNotaPedido.Enabled = False
        freCliente.Enabled = True
        TxtCodigoCli.SetFocus
    End If
End Sub

Private Sub chkVendedor_Click()
    If chkVendedor.Value = Checked Then
        txtVendedor.Enabled = True
        cmdBuscarVendedor.Enabled = True
    Else
        txtVendedor.Enabled = False
        cmdBuscarVendedor.Enabled = False
    End If
End Sub

Private Sub cmdAgregarProducto_Click()
    Consulta = 3
    'ABMProducto.CODIGOLISTA = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
    ABMProducto.Show vbModal
    If Consulta <> 4 Then
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        If Trim(ABMProducto.TxtCodigo) <> "" Then txtEdit.Text = ABMProducto.TxtCodigo
        TxtEdit_KeyDown vbKeyReturn, 0
    End If
    'grdGrilla.SetFocus
    'grdGrilla.row = 1
End Sub
Function BuscoImporte(nRemito As Integer, nSucursal As Integer) As Double
    Dim nsubtotal As Double
    Dim ntotal As Double
    Dim nIVA As Double
    ntotal = 0
    sql = "SELECT DR.DRC_PRECIO,DR.DRC_CANTIDAD,P.LNA_CODIGO "
    sql = sql & " FROM DETALLE_REMITO_CLIENTE DR,PRODUCTO P "
    sql = sql & " WHERE "
    sql = sql & " DR.PTO_CODIGO = P.PTO_CODIGO"
    sql = sql & " AND DR.RCL_NUMERO =" & nRemito
    sql = sql & " AND DR.RCL_SUCURSAL =" & nSucursal
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        ' me fijo si es maquinaria o repuesto
        If Rec1!LNA_CODIGO = 6 Then
            nIVA = "0,105"
        Else
            nIVA = "1,21"
        End If
        Do While Rec1.EOF = False
            
            nsubtotal = IIf(IsNull(Rec1!DRC_CANTIDAD), 1, Rec1!DRC_CANTIDAD) * Rec1!drc_precio
            ntotal = ntotal + nsubtotal
            Rec1.MoveNext
        Loop
    End If
    ntotal = ntotal * nIVA
    Rec1.Close
    BuscoImporte = ntotal
End Function
Private Sub CmdBuscAprox_Click()
    Dim TotalRemito As Double
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
'    If (chkPend.Value = Unchecked) And (chkDef.Value = Unchecked) And (chkAnu.Value = Unchecked) Then
'        MsgBox "Debe Seleccionar un Estado del Remito", vbInformation
'        chkPend.SetFocus
'    End If
    Select Case TipoBusquedaDoc
    
    
    Case 1 'BUSCA REMITOS
        
        sql = "SELECT RC.*, C.CLI_RAZSOC,C.CLI_DOMICI,L.LOC_DESCRI,P.PRO_DESCRI"
        sql = sql & " FROM REMITO_CLIENTE RC,CLIENTE C, LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE"
        sql = sql & "  RC.CLI_CODIGO=C.CLI_CODIGO"
        sql = sql & "  AND C.LOC_CODIGO=L.LOC_CODIGO"
        sql = sql & "  AND C.PRO_CODIGO=P.PRO_CODIGO"
        sql = sql & "  AND L.PRO_CODIGO=P.PRO_CODIGO"
        If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
        If FechaDesde <> "" Then sql = sql & " AND RC.RCL_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND RC.RCL_FECHA<=" & XDQ(FechaHasta)
        If optPen.Value = True Then
            sql = sql & " AND RC.EST_CODIGO = 1 "
        End If
        If optDef.Value = True Then
            sql = sql & " AND RC.EST_CODIGO = 3 "
        End If
        If optAnu.Value = True Then
            sql = sql & " AND RC.EST_CODIGO = 2 "
        End If
        sql = sql & " ORDER BY RC.RCL_NUMERO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                TotalRemito = BuscoImporte(rec!RCL_NUMERO, rec!RCL_SUCURSAL)
                GrdModulos.AddItem Format(rec!RCL_SUCURSAL, "0000") & "-" & Format(rec!RCL_NUMERO, "00000000") _
                                & Chr(9) & rec!RCL_FECHA _
                                & Chr(9) & Format(TotalRemito, "##0.00") _
                                & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!CLI_DOMICI _
                                & Chr(9) & rec!LOC_DESCRI & Chr(9) & rec!PRO_DESCRI _
                                & Chr(9) & rec!EST_CODIGO _
                                & Chr(9) & rec!NPE_NUMERO & Chr(9) & rec!NPE_FECHA _
                                & Chr(9) & rec!RCL_OBSERVACION & Chr(9) & rec!STK_CODIGO _
                                & Chr(9) & rec!RCL_SINFAC & Chr(9) & rec!CLI_CODIGO
                rec.MoveNext
            Loop
            GrdModulos.SetFocus
        Else
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        End If
        
    Case 2 'BUSCA NOTA DE PEDIDO
        
        sql = "SELECT NP.NPE_NUMERO, NP.NPE_FECHA, C.CLI_RAZSOC, "
        sql = sql & " C.CLI_CODIGO,C.CLI_DOMICI,L.LOC_DESCRI,P.PRO_DESCRI"
        sql = sql & " FROM NOTA_PEDIDO NP, CLIENTE C, LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE"
        sql = sql & " NP.CLI_CODIGO=C.CLI_CODIGO"
        sql = sql & " AND L.LOC_CODIGO=C.LOC_CODIGO"
        sql = sql & " AND P.PRO_CODIGO=C.PRO_CODIGO"
        sql = sql & " AND P.PRO_CODIGO=L.PRO_CODIGO"
        sql = sql & " AND NP.EST_CODIGO = 1"
        If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente)
        If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
        If FechaDesde <> "" Then sql = sql & " AND NP.NPE_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND NP.NPE_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY NPE_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem rec!NPE_NUMERO & Chr(9) & rec!NPE_FECHA _
                                & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!CLI_DOMICI _
                                & Chr(9) & rec!LOC_DESCRI & Chr(9) & rec!PRO_DESCRI _
                                & Chr(9) & rec!CLI_CODIGO
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

Private Sub cmdBuscarNotaPedido_Click()
    TipoBusquedaDoc = 2
    tabDatos.Tab = 1
End Sub

Private Sub cmdBuscarProducto_Click()
'    grdGrilla.SetFocus
'    frmBuscar.TipoBusqueda = 2
'    frmBuscar.CodListaPrecio = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
'    frmBuscar.TxtDescriB.Text = ""
'    frmBuscar.Show vbModal
'    grdGrilla.Col = 0
'    EDITAR grdGrilla, txtEdit, 13
'    If Trim(frmBuscar.grdBuscar.Text) <> "" Then txtEdit.Text = frmBuscar.grdBuscar.Text
'    TxtEdit_KeyDown vbKeyReturn, 0
    Consulta = 3
    'FrmListadePrecios.cbodescri.ListIndex = cboListaPrecio.ListIndex
    
    If tabLista.Tab = 0 Then
        FrmListadePrecios.tabLista.Tab = 0
        FrmListadePrecios.cboListaPrecio.ListIndex = cboListaPrecio.ListIndex
        
    Else
        nlista = 1
        FrmListadePrecios.tabLista.Tab = 1
        FrmListadePrecios.cboLPrecioRep.ListIndex = cboLPrecioRep.ListIndex
    End If
    
    FrmListadePrecios.Show vbModal
    If Consulta <> 4 Then
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        If Trim(FrmListadePrecios.GrdModulos.Text) <> "" Then txtEdit.Text = FrmListadePrecios.GrdModulos.Text
        TxtEdit_KeyDown vbKeyReturn, 0
    End If
    Consulta = 2
End Sub

Private Sub cmdBuscarVendedor_Click()
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

Private Sub CmdGrabar_Click()
    
    On Error GoTo HayErrorRemito
    If ValidarRemito = False Then Exit Sub
    If MsgBox("¿Confirma Remito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    DBConn.BeginTrans
    sql = "SELECT * FROM REMITO_CLIENTE"
    sql = sql & " WHERE RCL_NUMERO=" & XN(txtNroRemito)
    sql = sql & " AND RCL_SUCURSAL=" & XN(txtNroSucursal)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = True Then 'NUEVO REMITO
        sql = "INSERT INTO REMITO_CLIENTE"
        sql = sql & " (RCL_NUMERO,RCL_SUCURSAL,RCL_FECHA,NPE_NUMERO,"
        sql = sql & "NPE_FECHA,RCL_OBSERVACION,"
        sql = sql & "EST_CODIGO,RCL_NUMEROTXT,STK_CODIGO, CLI_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(txtNroRemito) & ","
        sql = sql & XN(txtNroSucursal) & ","
        sql = sql & XDQ(FechaRemito) & ","
        sql = sql & XN(txtNroNotaPedido) & ","
        sql = sql & XDQ(FechaNotaPedido) & ","
        sql = sql & XS(txtObservaciones) & ","
        sql = sql & "1,"    'ESTADO PENDIENTE
        sql = sql & XS(Format(txtNroRemito.Text, "00000000")) & ","
        'sql = sql & cboStock.ItemData(cboStock.ListIndex) & ","
        sql = sql & 1 & ","
        sql = sql & XN(TxtCodigoCli.Text) & ")"
        DBConn.Execute sql
           
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                sql = "INSERT INTO DETALLE_REMITO_CLIENTE"
                sql = sql & " (RCL_NUMERO,RCL_SUCURSAL,RCL_FECHA,DRC_NROITEM,"
                sql = sql & "PTO_CODIGO,DRC_CANTIDAD,DRC_PRECIO,DRC_DETALLE)"
                sql = sql & " VALUES ("
                sql = sql & XN(txtNroRemito) & ","
                sql = sql & XN(txtNroSucursal) & ","
                sql = sql & XDQ(FechaRemito) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 6)) & ","
                If grdGrilla.TextMatrix(I, 0) <> "----------" Then
                    sql = sql & XS(grdGrilla.TextMatrix(I, 0), True) & ","
                Else
                    sql = sql & "99999999" & ","
                End If
                sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ","
                grdGrilla.TextMatrix(I, 1) = Replace(grdGrilla.TextMatrix(I, 1), "'", "´")
                sql = sql & XS(grdGrilla.TextMatrix(I, 1)) & ")"
                DBConn.Execute sql
            End If
        Next
        'ACTUALIZO EL STOCK CUANDO EL REMITO ES DEFINITIVO (STOCK PENDIENTE)
        'Y ES REMITO SIN FACTURAS
       ' If chkRemitoSinFactura.Value = Checked Then
            For I = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(I, 0) <> "" Then
                        sql = "UPDATE DETALLE_STOCK"
                        sql = sql & " SET"
                        'sql = sql & " DST_STKPEN = DST_STKPEN + " & XN(grdGrilla.TextMatrix(I, 2))
                        sql = sql & " DST_STKFIS = DST_STKFIS - " & XN(grdGrilla.TextMatrix(I, 2))
                        sql = sql & " WHERE STK_CODIGO = 1 "
                        '& cboStock.ItemData(cboStock.ListIndex)
                        sql = sql & " AND PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "' "
                        DBConn.Execute sql
                End If
            Next
       ' End If
        'CAMBIO ESTADO DE LA NOTA DE PEDIDO (LE PONGO DEFINITIVO)
        If chkNotaPedido.Value = 1 Then
            sql = "UPDATE NOTA_PEDIDO SET EST_CODIGO=3"
            sql = sql & " WHERE"
            sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
        End If
        
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO AL REMITO
        sql = "UPDATE PARAMETROS SET NRO_REMITO=" & XN(txtNroRemito)
        DBConn.Execute sql
        
        DBConn.CommitTrans
    Else
        ' modifico el Remito
        'If MsgBox("Confirma modificar el Remito Nro.: " & Trim(txtNroSucursal.Text) & "-" & Trim(txtNroNotaPedido.Text) & " ", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            sql = "UPDATE REMITO_CLIENTE"
            sql = sql & " SET CLI_CODIGO=" & XN(TxtCodigoCli)
            sql = sql & " ,RCL_OBSERVACION=" & XS(txtObservaciones)
            sql = sql & " ,STK_CODIGO= 1"
            sql = sql & " ,RCL_NUMEROTXT=" & XS(Format(txtNroRemito.Text, "00000000"))
            
            sql = sql & " WHERE"
            sql = sql & " RCL_NUMERO=" & XN(txtNroRemito)
            sql = sql & " AND RCL_FECHA=" & XDQ(FechaRemito)
            DBConn.Execute sql
            
            sql = "DELETE FROM DETALLE_REMITO_CLIENTE"
            sql = sql & " WHERE RCL_NUMERO=" & XN(txtNroRemito)
            sql = sql & " AND RCL_SUCURSAL=" & XN(txtNroSucursal)
            sql = sql & " AND RCL_FECHA=" & XDQ(FechaRemito)
            DBConn.Execute sql
            
            For I = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(I, 0) <> "" Then
                    sql = "INSERT INTO DETALLE_REMITO_CLIENTE"
                    sql = sql & " (RCL_NUMERO,RCL_SUCURSAL,RCL_FECHA,DRC_NROITEM,PTO_CODIGO,"
                    sql = sql & "DRC_CANTIDAD,DRC_PRECIO,DRC_DETALLE)"
                    sql = sql & " VALUES ("
                    sql = sql & XN(txtNroRemito) & ","
                    sql = sql & XN(txtNroSucursal) & ","
                    sql = sql & XDQ(FechaRemito) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 6)) & "," 'NRO ITEM
                    If grdGrilla.TextMatrix(I, 0) <> "----------" Then
                        sql = sql & XS(grdGrilla.TextMatrix(I, 0), True) & "," 'PRODUCTO CODIGO
                    Else
                        sql = sql & "99999999" & ","
                    End If
                                        
                    sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & "," 'CANTIDAD
                    
                    'MsgBox "CANTIDAD = " & XN(grdGrilla.TextMatrix(I, 2))
                    
                    sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & "," 'PRECIO
                    
                    grdGrilla.TextMatrix(I, 1) = Replace(grdGrilla.TextMatrix(I, 1), "'", "´")
                    sql = sql & XS(grdGrilla.TextMatrix(I, 1)) & ")" 'DESCRIPCION
                    DBConn.Execute sql
                End If
            Next
            
            'ACTUALIZO EL STOCK CUANDO EL REMITO ES DEFINITIVO (STOCK PENDIENTE)
            'Y ES REMITO SIN FACTURAS
            ' If chkRemitoSinFactura.Value = Checked Then
            For I = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(I, 0) <> "" Then
                        sql = "UPDATE DETALLE_STOCK"
                        sql = sql & " SET"
                        sql = sql & " DST_STKPEN = DST_STKPEN - " & XN(IIf(grdGrilla.TextMatrix(I, 7) = "", "0", grdGrilla.TextMatrix(I, 7))) + XN(grdGrilla.TextMatrix(I, 2))
                        sql = sql & " WHERE STK_CODIGO = 1"
                        '& cboStock.ItemData(cboStock.ListIndex)
                        sql = sql & " AND PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "' "
                        DBConn.Execute sql
                End If
            Next
            ' End If
            
            DBConn.CommitTrans
            End If
        'End If
        
        
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    
    If MsgBox("¿Desea Facturar el Remito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        
        frmFacturaCliente.txtRemSuc = Format(txtNroSucursal.Text, "0000")
        frmFacturaCliente.txtNroRemito = Format(txtNroRemito.Text, "00000000")
        
        frmFacturaCliente.Show vbModal
        
        
    End If
    
    CmdNuevo_Click
    Exit Sub
    
HayErrorRemito:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarRemito() As Boolean
    Dim BAND As Integer
    If FechaRemito.Text = "" Then
        MsgBox "La Fecha del Remito es requerida", vbExclamation, TIT_MSGBOX
        FechaRemito.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    If chkNotaPedido.Value = 1 Then
        If txtNroNotaPedido.Text = "" Then
            MsgBox "El número de Nota de Pedido es requerido", vbExclamation, TIT_MSGBOX
            txtNroNotaPedido.SetFocus
            ValidarRemito = False
            Exit Function
        End If
        If FechaNotaPedido.Text = "" Then
            MsgBox "La Fecha de la Nota de pedido es requerida", vbExclamation, TIT_MSGBOX
            FechaNotaPedido.SetFocus
            ValidarRemito = False
            Exit Function
        End If
    End If
'    If chkRemitoSinFactura.Value = Checked Then
'        If txtConcepto.Text = "" Then
'            MsgBox "Debe ingresar un concepto", vbExclamation, TIT_MSGBOX
'            txtConcepto.SetFocus
'            ValidarRemito = False
'            Exit Function
'        End If
'    End If
    If TxtCodigoCli.Text = "" Then
        MsgBox "Debe ingresar un Cliente", vbExclamation, TIT_MSGBOX
        TxtCodigoCli.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    BAND = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            BAND = 1
        End If
    Next
    If BAND = 0 Then
        MsgBox "Debe ingresar un item en detalle", vbExclamation, TIT_MSGBOX
        grdGrilla.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    ValidarRemito = True
End Function

Private Sub cmdImprimir_Click()
    If MsgBox("¿Confirma Impresión Remito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    VCantidadBultos = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            VCantidadBultos = CInt(grdGrilla.TextMatrix(I, 2)) + VCantidadBultos
        End If
    Next I
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
    ImprimirRemito
End Sub


Public Sub ImprimirEncabezado()
 '---------------IMPRIME EL ENCABEZADO DEL REMITO-------------------
    'Imprimir 15.8, 2.7, False, Format(FechaRemito, "dd/mm/yyyy")
    Imprimir 15.7, 4.4, False, Format(Day(FechaRemito), "00")
    Imprimir 16.75, 4.4, False, Format(Month(FechaRemito), "00")
    Imprimir 17.8, 4.4, False, Format(Year(FechaRemito), "yy")
    
    'PROBAR IMPRESIÓN
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CUIT,C.CLI_INGBRU, L.LOC_DESCRI,CI.IVA_CODIGO"
    sql = sql & ", P.PRO_DESCRI,CI.IVA_DESCRI"
    sql = sql & " FROM CLIENTE C, LOCALIDAD L, REMITO_CLIENTE R, PROVINCIA P, CONDICION_IVA CI"
    sql = sql & " WHERE  R.RCL_NUMERO=" & XN(txtNroRemito)
    sql = sql & " AND R.RCL_FECHA=" & XDQ(FechaRemito)
    sql = sql & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Imprimir 2.5, 7, True, Trim(Rec1!CLI_RAZSOC)
        Imprimir 10.8, 6.8, False, Trim(Rec1!CLI_DOMICI)
        'nota de pedido
        'Imprimir 13, 5.3, True, "Nro.Pedido: " & Format(txtNroNotaPedido.Text, "00000000")
        Imprimir 10.8, 7.2, False, Trim(Rec1!LOC_DESCRI) & " - " & Trim(Rec1!PRO_DESCRI)
        'fecha nota pedido
        'If Rec1!IVA_CODIGO = 1 Then
        '    Imprimir 3.7, 7.8, False, "X"
        'Else
        '    Imprimir 9.5, 7.8, False, "X"
        'End If
        'Imprimir 13, 5.7, True, "Fecha: " & Format(FechaNotaPedido.Text, "dd/mm/yyyy")
        'Imprimir 1, 6.4, False, Trim(Rec1!IVA_DESCRI)
        Imprimir 13.7, 7.8, False, IIf(IsNull(Rec1!CLI_CUIT), "", Format(Rec1!CLI_CUIT, "##-########-#"))
        Imprimir 14.2, 8.5, False, IIf(IsNull(Rec1!CLI_INGBRU), "", Format(Rec1!CLI_INGBRU, "###-#####-##"))
    End If
    Rec1.Close
    'Imprimir 18.4, 7.9, False, CStr(VCantidadBultos)
    'Imprimir 0, 9.1, False, "Código"
    'Imprimir 3.1, 9.1, False, "Descripción"
    'Imprimir 12, 9.1, False, "Cantidad"
    'Imprimir 15, 9.1, False, "Rubro"
End Sub

Public Sub ImprimirRemito()
    Dim Renglon As Double
    Dim canttxt As Integer
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Imprimiendo..."
    
    For w = 1 To 2 'SE IMPRIME POR DUPLICADO
      '-----IMPRESION DEL ENCABEZADO------------------
        ImprimirEncabezado
        
      '---- IMPRESION DEL REINTEGRO ------------------
        Renglon = 9.8
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                Imprimir 1, Renglon, False, Format(grdGrilla.TextMatrix(I, 0), "000000000")  'codigo
                If Len(grdGrilla.TextMatrix(I, 1)) < 35 Then
                    Imprimir 3, Renglon, False, grdGrilla.TextMatrix(I, 1)  'descripcion
                Else
                    CortarCadena Renglon, grdGrilla.TextMatrix(I, 1)
'                    Imprimir 3, Renglon, False, Left(grdGrilla.TextMatrix(I, 1), 36) 'descripcion
'                    Imprimir 3, Renglon + 0.5, False, Mid(grdGrilla.TextMatrix(I, 1), 37, 38) 'descripcion
'                    Imprimir 3, Renglon + 1, False, Mid(grdGrilla.TextMatrix(I, 1), 75, 36) 'descripcion
'                    Imprimir 3, Renglon + 1.5, False, Mid(grdGrilla.TextMatrix(I, 1), 111, 36) 'descripcion
'                    Imprimir 3, Renglon + 2, False, Mid(grdGrilla.TextMatrix(I, 1), 147, 36) 'descripcion
'                    Imprimir 3, Renglon + 2.5, False, Mid(grdGrilla.TextMatrix(I, 1), 183, 36) 'descripcion
'                    Imprimir 3, Renglon + 3, False, Mid(grdGrilla.TextMatrix(I, 1), 219, 36) 'descripcion
                    canttxt = Len(grdGrilla.TextMatrix(I, 1))
                    canttxt = canttxt / 36 'es para sacar la cantidad de renglones
                    canttxt = Int(canttxt)
                End If
                Imprimir 13.2, Renglon, False, grdGrilla.TextMatrix(I, 2) 'cantidad
                'Imprimir 15, Renglon, False, Left(grdGrilla.TextMatrix(I, 4), 20) 'rubro
                Renglon = Renglon + (canttxt * 0.5) + 0.5
            End If
        Next I
        '-----OBSERVACIONES---------------------
        If txtObservaciones.Text <> "" Then
            Imprimir 0.5, Renglon + 1, False, "Observaciones: " & Trim(txtObservaciones.Text)
        End If
        'txtObservaciones
        '------------DATOS REPRESENTADA----------------------
'           Set Rec1 = New ADODB.Recordset
'           If chkFacturaTerceros.Value = Checked Then
'                sql = "SELECT REP_RAZSOC, REP_CUIT"
'                sql = sql & " FROM REPRESENTADA"
'                sql = sql & " WHERE REP_CODIGO=" & cboRep.ItemData(cboRep.ListIndex)
'                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'                If Rec1.EOF = False Then
'                    Imprimir 0, 16.2, True, "Corresponde a Factura Nro.: " & Format(txtNroFactura.Text, "00000000")
'                    Imprimir 0, 16.6, True, "Por Cuenta y Orden de " & Trim(Rec1!REP_RAZSOC)
'                    Imprimir 0, 17, True, "CUIT: " & IIf(IsNull(Rec1!REP_CUIT), "NO INFORMADO", Format(Rec1!REP_CUIT, "##-########-#"))
'                End If
'                Rec1.Close
'           End If
'          '------------DATOS SUCURSAL-------------------------
'
'           sql = "SELECT S.SUC_DESCRI,S.SUC_DOMICI, L.LOC_DESCRI"
'           sql = sql & " FROM SUCURSAL S, NOTA_PEDIDO NP, LOCALIDAD L"
'           sql = sql & " WHERE NP.NPE_NUMERO=" & XN(txtNroNotaPedido)
'           sql = sql & " AND NP.NPE_FECHA=" & XDQ(FechaNotaPedido)
'           sql = sql & " AND NP.SUC_CODIGO=S.SUC_CODIGO"
'           sql = sql & " AND S.LOC_CODIGO=L.LOC_CODIGO"
'           sql = sql & " AND S.PRO_CODIGO=L.PRO_CODIGO"
'           sql = sql & " AND S.PAI_CODIGO=L.PAI_CODIGO"
'           Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'           If Rec1.EOF = False Then
'                Imprimir 0, 17.4, True, "Entregar: " & Left(Trim(Rec1!SUC_DESCRI), 25) & " -- " & Left(Trim(Rec1!SUC_DOMICI), 30) & " (" & Left(Trim(Rec1!LOC_DESCRI), 20) & ")"
'           End If
'           Rec1.Close
          '----------------------------------------------------
        Printer.EndDoc
    Next w
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
   For I = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(I, 0) = ""
        grdGrilla.TextMatrix(I, 1) = ""
        grdGrilla.TextMatrix(I, 2) = ""
        grdGrilla.TextMatrix(I, 3) = ""
        grdGrilla.TextMatrix(I, 4) = ""
        grdGrilla.TextMatrix(I, 5) = ""
        grdGrilla.TextMatrix(I, 6) = I
   Next
   'grillaNotaPedido.TextMatrix(0, 1) = ""
   'grillaNotaPedido.TextMatrix(1, 1) = ""
   'grillaNotaPedido.TextMatrix(2, 1) = ""
   FechaNotaPedido.Text = ""
   txtNroNotaPedido.Text = ""
 '  chkRemitoSinFactura.Value = Unchecked
  ' txtConcepto.Text = ""
   lblEstadoRemito.Caption = ""
   txtObservaciones.Text = ""
   lblEstado.Caption = ""
   cmdGrabar.Enabled = True
   freRemito.Enabled = True
   freCliente.Enabled = True
   
    'BUSCO EL NUMERO DE REMITO QUE CORRESPONDE
    txtNroRemito.Text = BuscoUltimoRenito
   'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRemito) 'ESTADO PENDIENTE
    VEstadoRemito = 1
    '--------------
    FechaRemito.Enabled = True
    txtNroNotaPedido.Enabled = True
    FechaNotaPedido.Enabled = True
    cmdBuscarNotaPedido.Enabled = True
    '--------------
    tabDatos.Tab = 0
    TipoBusquedaDoc = 1
    FechaRemito.Text = Date
    cboListaPrecio.ListIndex = 0
    cboListaPrecio.Enabled = True
    cboListaPrecio.SetFocus
    
    TxtCodigoCli.Text = ""
    TxtCodigoCli_Change
    
    
    chkNotaPedido.Enabled = True
    chkNotaPedido.Value = 0
    freNotaPedido.Enabled = False
    freCliente.Enabled = True
    
End Sub

Private Sub cmdNuevoCliente_Click()
    ABMCliente.Show vbModal
    TxtCodigoCli.SetFocus
End Sub
Private Sub cmdQuitarProducto_Click()
    If MsgBox("Seguro que desea quitar el Detalle: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 1), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
    End If
End Sub

Private Sub cmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmRemitoCliente = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        TipoBusquedaDoc = 1
        tabDatos.Tab = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl.Name <> "grdGrilla" And _
        Me.ActiveControl.Name <> "txtEdit" And _
        KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    grdGrilla.FormatString = "Código|Descripción|Cantidad|Precio|Rubro|Linea|Orden|CANT"
    grdGrilla.ColWidth(0) = 1500 'CODIGO
    grdGrilla.ColWidth(1) = 5900 'DESCRIPCION
    grdGrilla.ColWidth(2) = 1000 'CANTIDAD
    grdGrilla.ColWidth(3) = 1100 'PRECIO
    grdGrilla.ColWidth(4) = 2100 'RUBRO
    grdGrilla.ColWidth(5) = 2100 'LINEA
    grdGrilla.ColWidth(6) = 0    'ORDEN
    grdGrilla.ColWidth(7) = 0    'CANTIDAD STOCK
    grdGrilla.Cols = 8
    grdGrilla.Rows = 1
    For I = 2 To 14
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & (I - 1)
    Next
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "^Número|^Fecha|Importe|Cliente|Domicilio|Localidad|Provincia|Cod Estado|NP NUMERO|NP FECHA|OBSERVACIONES|" _
                              & "STOCK|REMITO SIN FACTURA|CODIGOCLIENTE"
    GrdModulos.ColWidth(0) = 1500 'NUMERO
    GrdModulos.ColWidth(1) = 1000 'FECHA
    GrdModulos.ColWidth(2) = 1000 'IMPORTE A COBRAR
    GrdModulos.ColWidth(3) = 3200 'CLIENTE
    GrdModulos.ColWidth(4) = 3200 'DOMICILIO
    GrdModulos.ColWidth(5) = 3200 'Localidad
    GrdModulos.ColWidth(6) = 3200 'Provincia
    GrdModulos.ColWidth(7) = 0    'COD ESTADO
    GrdModulos.ColWidth(8) = 0    'NOTA PEDIDO NUMERO
    GrdModulos.ColWidth(9) = 0    'NOTA PEDIDO FECHA
    GrdModulos.ColWidth(10) = 0   'OBSERVACIONES
    GrdModulos.ColWidth(11) = 0   'STOCK
    GrdModulos.ColWidth(12) = 0   'REMITO SIN FACTURAS
    GrdModulos.ColWidth(13) = 0   'CODIGOCLIENTE
    
    GrdModulos.Rows = 1
    '------------------------------------
    'grillaNotaPedido.ColWidth(0) = 950
    'grillaNotaPedido.ColWidth(1) = 5300
    'grillaNotaPedido.TextMatrix(0, 0) = "    Cliente:"
    'grillaNotaPedido.TextMatrix(1, 0) = " Sucursal:"
    'grillaNotaPedido.TextMatrix(2, 0) = "Vendedor:"
    '------------------------------------
    lblEstado.Caption = ""
    'CARGO EL COMBO DE LISTA DE PRECIOS DE MAQUINARIAS
    CargoCboListaPrecio
    'CARGO EL COMBO DE LISTA DE PRECIOS DE REPUESTOS
    CargoCboLPrecioRep
    'BUSCO EL NUMERO DE REMITO QUE CORRESPONDE
    txtNroRemito.Text = BuscoUltimoRenito
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRemito) 'ESTADO PENDIENTE
    VEstadoRemito = 1
    FechaRemito.Text = Date
    'CARGO COMBO STOCK
    'CargaCboStock
    'PONGO ENABLE LOS DATOS DE LA FACTURA DE TERCEROS

    'txtConcepto.Enabled = False
    TipoBusquedaDoc = 1 'ESTO ES PARA BUSCAR REMITOS(1), (2)PARA BUSCAR NOTA DE PEDIDO
    tabDatos.Tab = 0
    
    freNotaPedido.Enabled = False
    freCliente.Enabled = True
End Sub

Private Sub CargaCboStock()
    sql = "SELECT S.STK_CODIGO,R.REP_RAZSOC"
    sql = sql & " FROM STOCK S, REPRESENTADA R"
    sql = sql & " WHERE S.REP_CODIGO=R.REP_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboStock.AddItem rec!REP_RAZSOC
            cboStock.ItemData(cboStock.NewIndex) = rec!STK_CODIGO
            rec.MoveNext
        Loop
        cboStock.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub CargoCboListaPrecio() '' Lista de Precios de Repuestos
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 6"   '6: Maquinaria
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboListaPrecio.AddItem rec!LIS_DESCRI
            cboListaPrecio.ItemData(cboListaPrecio.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboListaPrecio.ListIndex = 0
    End If
    rec.Close
End Sub
Private Sub CargoCboLPrecioRep() '' Lista de Precios de Repuestos
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 7"   '6: Repuestos
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboLPrecioRep.AddItem rec!LIS_DESCRI
            cboLPrecioRep.ItemData(cboLPrecioRep.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboLPrecioRep.ListIndex = 0
    End If
    rec.Close
End Sub

Private Function BuscoUltimoRenito() As String
    'ACA BUSCA EL NUMERO DE REMITO SIGUIENTE AL ULTIMO CARGADO
    sql = "SELECT (NRO_REMITO) + 1 AS ULTIMO"
    sql = sql & " FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        txtNroSucursal.Text = Sucursal
        BuscoUltimoRenito = Format(rec!Ultimo, "00000000")
    End If
    rec.Close
End Function

Private Sub grdGrilla_Click()
'    If grdGrilla.Col = 1 Then
'        If grdGrilla.TextMatrix(grdGrilla.RowSel, 1) <> "" Then
'            txtEdit.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 1)
'            txtEdit_GotFocus
'        End If
'    End If
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1
            LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
            grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
            grdGrilla.Col = 0
        'Case Else
        '    grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, grdGrilla.Col)) = ""
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
        Case 1
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = "" Then
                txtObservaciones.SetFocus
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
               ' End If
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
        Select Case TipoBusquedaDoc
        Case 1 'BUSCA REMITOS
            lblEstado.Caption = "Buscando..."
            Screen.MousePointer = vbHourglass
            Set Rec1 = New ADODB.Recordset
            
            txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4)
            txtNroRemito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)
            FechaRemito.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
            'CARGO EL ESTADO
            Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 7)), lblEstadoRemito)
            VEstadoRemito = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 7))
            
            If VEstadoRemito <> 1 Then
                cmdGrabar.Enabled = False
                freCliente.Enabled = False
                freNotaPedido.Enabled = False
                freRemito.Enabled = False
                chkNotaPedido.Enabled = False
                grdGrilla.SetFocus
            Else
                cmdGrabar.Enabled = True
                freCliente.Enabled = True
                freNotaPedido.Enabled = True
                freRemito.Enabled = True
                chkNotaPedido.Enabled = True
            End If
            
            
            txtNroNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
            FechaNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 9)
            
            
            'BUSCO STOCK
            'If GrdModulos.TextMatrix(GrdModulos.RowSel, 10) <> "" Then
            '    Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 10)), cboStock)
            'Else
            '    cboStock.ListIndex = 0
            'End If
            txtObservaciones.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 9)
            
            TxtCodigoCli.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 13)
            TxtCodigoCli_LostFocus
            
            
        '----BUSCO DETALLE DEL REMITO------------------
            
            sql = "SELECT DRC.*, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM DETALLE_REMITO_CLIENTE DRC, PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE DRC.RCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4))
            sql = sql & " AND DRC.RCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8))
            sql = sql & " AND DRC.RCL_FECHA=" & XDQ(GrdModulos.TextMatrix(GrdModulos.RowSel, 1))
            sql = sql & " AND DRC.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " ORDER BY DRC.DRC_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = IIf(Rec1!PTO_CODIGO = 99999999, "----------", Rec1!PTO_CODIGO)
                    If (grdGrilla.TextMatrix(I, 0)) = "----------" Then
                        grdGrilla.TextMatrix(I, 1) = Rec1!DRC_DETALLE
                    Else
                        grdGrilla.TextMatrix(I, 1) = Rec1!PTO_DESCRI
                    
                        grdGrilla.TextMatrix(I, 2) = IIf(IsNull(Rec1!DRC_CANTIDAD), "", Rec1!DRC_CANTIDAD)
                        grdGrilla.TextMatrix(I, 3) = IIf(IsNull(Rec1!drc_precio), "", Valido_Importe(Rec1!drc_precio))
                        grdGrilla.TextMatrix(I, 4) = IIf(IsNull(Rec1!RUB_DESCRI), "", Rec1!RUB_DESCRI)
                        grdGrilla.TextMatrix(I, 5) = IIf(IsNull(Rec1!LNA_DESCRI), "", Rec1!LNA_DESCRI)
                        grdGrilla.TextMatrix(I, 6) = IIf(IsNull(Rec1!DRC_NROITEM), "", Rec1!DRC_NROITEM)
                        grdGrilla.TextMatrix(I, 7) = IIf(IsNull(Rec1!DRC_CANTIDAD), "", Rec1!DRC_CANTIDAD)
                    End If
                    I = I + 1
                    Rec1.MoveNext
                Loop
            End If
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            '--------------
            FechaRemito.Enabled = False
            txtNroNotaPedido.Enabled = False
            FechaNotaPedido.Enabled = False
            cmdBuscarNotaPedido.Enabled = False
            '--------------
            tabDatos.Tab = 0
            grdGrilla.SetFocus
            grdGrilla.row = 1
            Rec1.Close
        '----------------------------------------------
        Case 2 'BUSCA NOTA PEDIDO
            txtNroNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
            FechaNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
            tabDatos.Tab = 0
            txtNroNotaPedido_LostFocus
            TxtCodigoCli.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 7)
            TxtCodigoCli_LostFocus
        End Select
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    txtCliente.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    txtVendedor.Enabled = False
    cmdGrabar.Enabled = False
    cmdBuscarCli.Enabled = False
    cmdBuscarVendedor.Enabled = False
    LimpiarBusqueda
    If Me.Visible = True Then chkCliente.SetFocus
    If TipoBusquedaDoc = 1 Then
        frameBuscar.Caption = "Buscar Remito por..."
        chkVendedor.Enabled = False
        txtVendedor.Enabled = False
    Else
        frameBuscar.Caption = "Buscar Notas de Pedido Pendientes por..."
        chkVendedor.Enabled = True
        txtVendedor.Enabled = True
    End If
  Else
    If VEstadoRemito = 1 Then
        cmdGrabar.Enabled = True
    Else
        cmdGrabar.Enabled = False
    End If
  End If
End Sub

Private Sub LimpiarBusqueda()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
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
'        And chkVendedor.Value = Unchecked And ActiveControl.Name <> "cmdBuscarCli" _
'        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub

Private Function BuscoCondicionIVA(IVACodigo As String) As String
    sql = "SELECT * FROM CONDICION_IVA"
    sql = sql & " WHERE IVA_CODIGO=" & XN(IVACodigo)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscoCondicionIVA = Rec1!IVA_DESCRI
    Else
        BuscoCondicionIVA = ""
    End If
    Rec1.Close
End Function

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    CarTexto KeyAscii
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

Private Sub TxtCodigoCli_GotFocus()
    SelecTexto TxtCodigoCli
End Sub

Private Sub txtCodigoCli_KeyPress(KeyAscii As Integer)
        KeyAscii = CarNumeroEntero(KeyAscii)
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
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtRazSocCli.Text = rec!CLI_RAZSOC
            txtDomici.Text = IIf(IsNull(rec!CLI_DOMICI), "", rec!CLI_DOMICI)
            txtlocalidad.Text = rec!LOC_DESCRI
            txtprovincia.Text = rec!PRO_DESCRI
            txtCondicionIVA.Text = BuscoCondicionIVA(rec!IVA_CODIGO)
            txtCUIT.Text = IIf(IsNull(rec!CLI_CUIT), "", Format(rec!CLI_CUIT, "##-########-#"))
            txtIngBrutos.Text = IIf(IsNull(rec!CLI_INGBRU), "", Format(rec!CLI_INGBRU, "###-#####-##"))
            txtcodpos.Text = IIf(IsNull(rec!LOC_CODPOS), "", rec!LOC_CODPOS)
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtRazSocCli.Text = ""
            TxtCodigoCli.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtEdit_GotFocus()
    SelecTexto txtEdit
End Sub

'Private Sub txtEdit_Click()
'    If grdGrilla.Col = 2 Then
'        If txtEdit.Text <> "" Then
'            EDITAR grdGrilla, txtEdit, 1
'        End If
'    End If
'End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    'If grdGrilla.Col = 0 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 0 Then
        'CarTexto KeyAscii
        txtEdit.MaxLength = 10
    End If
    If grdGrilla.Col = 1 Then
        txtEdit.MaxLength = 50
    End If
    If grdGrilla.Col = 2 Then
        KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
        txtEdit.MaxLength = 10
    End If
    If grdGrilla.Col = 3 Then
        KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
        txtEdit.MaxLength = 10
    End If
    CarTexto KeyAscii
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 Then
        frmBuscar.TipoBusqueda = 2
        frmBuscar.CodListaPrecio = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        frmBuscar.Show vbModal
    End If

    If KeyCode = vbKeyReturn Then
       If grdGrilla.Col = 0 And Trim(txtEdit) <> "----------" Then
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
                txtEdit.Text = Replace(txtEdit.Text, "'", "´")
                If cboListaPrecio.ListIndex = 0 Then 'Busca en los Productos
                    sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI, P.PTO_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI"
                    sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, TIPO_PRESENTACION RE"
                    sql = sql & " WHERE"
                    If grdGrilla.Col = 0 Then
                        sql = sql & " P.PTO_CODIGO LIKE '" & txtEdit.Text & "'"
                    Else
                        sql = sql & " P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
                    End If
                    sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
                    sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                    sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
                    'sql = sql & " AND P.PTO_ESTADO=1"
                    '*********
                Else  ' Busca en un Lista de Precios
                    sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, P.PTO_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI"
                    sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L,TIPO_PRESENTACION RE"
                    sql = sql & " WHERE"
                    If grdGrilla.Col = 0 Then
                        sql = sql & " P.PTO_CODIGO LIKE '" & txtEdit.Text & "'"
                    Else
                        sql = sql & " P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
                    End If
                        'sql = sql & " AND P.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex) & ""
                        'sql = sql & " AND P.PTO_CODIGO=D.PTO_CODIGO"
                        sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
                        sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
                        sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                        sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
                       ' sql = sql & " AND P.PTO_ESTADO=1"
                End If
                rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If rec.EOF = False Then
                    If rec.RecordCount > 1 Then
                        grdGrilla.SetFocus
                        frmBuscar.TipoBusqueda = 2
                        'LE DIGO EN QUE LISTA DE PRECIO BUSCAR LOS PRECIOS
                        If cboListaPrecio.ListIndex <> 0 Then '<TODOS>
                            frmBuscar.CodListaPrecio = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
                        Else
                                frmBuscar.CodListaPrecio = 0 ' BUSCA EN LA TABLA PRODUCTOS
                        End If
                        frmBuscar.TxtDescriB.Text = txtEdit.Text
                        frmBuscar.Show vbModal
                        grdGrilla.Col = 0
                        grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
                        grdGrilla.Col = 1
                        grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
                        grdGrilla.Col = 3
                        grdGrilla.Text = Valido_Importe(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2))
                        grdGrilla.Col = 4
                        grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
                        grdGrilla.Col = 5
                        grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
                        grdGrilla.Col = 2
                    Else
                        grdGrilla.Col = 0
                        grdGrilla.Text = Trim(rec!PTO_CODIGO)
                        grdGrilla.Col = 1
                        grdGrilla.Text = Trim(rec!PTO_DESCRI)
                        grdGrilla.Col = 3
                        If cboListaPrecio.ListIndex = 0 Then
                            grdGrilla.Text = Valido_Importe(Trim(rec!PTO_PRECIO))
                        Else
                            grdGrilla.Text = Valido_Importe(Trim(rec!PTO_PRECIO))
                        End If
                        grdGrilla.Col = 4
                        grdGrilla.Text = Trim(rec!RUB_DESCRI)
                        grdGrilla.Col = 5
                        grdGrilla.Text = Trim(rec!LNA_DESCRI)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
                        grdGrilla.Col = 2
                    End If
                        If BuscoRepetetidos(grdGrilla.TextMatrix(grdGrilla.RowSel, 0), grdGrilla.RowSel) = False Then
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
                If Trim(txtEdit) = "" Then
                    grdGrilla.Text = "1"
                    txtEdit.Text = "1"
                End If
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                    Set Rec1 = New ADODB.Recordset
                    sql = "SELECT P.PTO_STKMIN, DS.DST_STKFIS, DS.DST_STKPEN"
                    sql = sql & " FROM PRODUCTO P, DETALLE_STOCK DS"
                    sql = sql & " WHERE P.PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(grdGrilla.RowSel, 0) & "'"
                    sql = sql & " AND P.PTO_CODIGO = DS.PTO_CODIGO"
                    sql = sql & " AND STK_CODIGO=" & cboStock.ItemData(cboStock.ListIndex)
                    
                    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                    If Rec1.EOF = False Then
                        If (CInt(Rec1!DST_STKFIS) - CInt(txtEdit.Text)) < CInt(Rec1!PTO_STKMIN) Then
                            MsgBox "El producto esta por debajo del Stock Minimo" & Chr(13) & Chr(13) & _
                            " Stock Minimo=" & Rec1!PTO_STKMIN & Chr(13) & _
                            " Stock Pendiente=" & Rec1!DST_STKPEN & Chr(13) & _
                            " Stock Fisico=" & Rec1!DST_STKFIS & Chr(13) & _
                            " Stock Fisico - Cantidad=" & (CInt(Rec1!DST_STKFIS) - CInt(txtEdit.Text)), vbExclamation, TIT_MSGBOX
                        End If
                    End If
                    Rec1.Close
                End If
                
            Case 3
                If Trim(txtEdit) <> "" Then
                    txtEdit.Text = Valido_Importe(txtEdit)
                Else
                    MsgBox "Debe ingresar el Importe", vbExclamation, TIT_MSGBOX
                    grdGrilla.Col = 3
                End If
            End Select
        End If
        grdGrilla.SetFocus
    End If
    If KeyCode = vbKeyEscape Then
       txtEdit.Visible = False
       grdGrilla.SetFocus
    End If
End Sub

Private Function BuscoRepetetidos(Codigo As String, Linea As Integer) As Boolean
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" And grdGrilla.TextMatrix(I, 0) <> "----------" Then
            If Codigo = grdGrilla.TextMatrix(I, 0) And (I <> Linea) Then
                MsgBox "El producto ya fue elegido anteriormente", vbExclamation, TIT_MSGBOX
                BuscoRepetetidos = False
                Exit Function
            End If
        End If
    Next
    BuscoRepetetidos = True
End Function

Private Sub txtNroNotaPedido_GotFocus()
    SelecTexto txtNroNotaPedido
End Sub

Private Sub txtNroNotaPedido_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaPedido_LostFocus()
    
    If txtNroNotaPedido.Text <> "" Then
        sql = "SELECT NP.*, E.EST_DESCRI"
        sql = sql & " FROM NOTA_PEDIDO NP, ESTADO_DOCUMENTO E"
        sql = sql & " WHERE NP.NPE_NUMERO=" & XN(txtNroNotaPedido)
        If FechaNotaPedido.Text <> "" Then
            sql = sql & " AND NP.NPE_FECHA=" & XDQ(FechaNotaPedido)
        End If
        sql = sql & " AND NP.EST_CODIGO=E.EST_CODIGO"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec2.EOF = False Then
            If Rec2.RecordCount > 1 Then
                MsgBox "Hay mas de una Nota de Pedido con el Número: " & txtNroNotaPedido.Text, vbInformation, TIT_MSGBOX
                Rec2.Close
                cmdBuscarNotaPedido_Click
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Buscando..."
            
            'CARGO CABECERA DE LA NOTA DE PEDIDO
            FechaNotaPedido.Text = Rec2!NPE_FECHA
            'grillaNotaPedido.TextMatrix(0, 1) = BuscoCliente(Rec2!CLI_CODIGO)
            'grillaNotaPedido.TextMatrix(1, 1) = BuscoSucursal(Rec2!SUC_CODIGO, Rec2!CLI_CODIGO)
            'grillaNotaPedido.TextMatrix(2, 1) = BuscoVendedor(Rec2!VEN_CODIGO)
            'lblEstadoNotaPedido.Caption = "Estado: " & Rec2!EST_DESCRI
            TxtCodigoCli.Text = Rec2!CLI_CODIGO
            TxtCodigoCli_LostFocus
            If Rec2!EST_CODIGO <> 1 Then
                MsgBox "La Nota de Pedido número: " & txtNroNotaPedido.Text & Chr(13) & Chr(13) & _
                       "No puede ser asignada al Remito por su estado (" & Rec2!EST_DESCRI & ")", vbExclamation, TIT_MSGBOX
                LimpiarNotaPedido
                cmdGrabar.Enabled = False
                Screen.MousePointer = vbNormal
                lblEstado.Caption = ""
                Rec2.Close
                Exit Sub
            Else
                cmdGrabar.Enabled = True
            End If
            Rec2.Close
            
        '-----BUSCO LOS DATOS DEL DETALLE DE LA NOTA DE PEDIDO---------
            sql = "SELECT DNP.*,P.PTO_DESCRI, D.LIS_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM DETALLE_NOTA_PEDIDO DNP, PRODUCTO P, RUBROS R, LINEAS L, DETALLE_LISTA_PRECIO D"
            sql = sql & " WHERE DNP.NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND DNP.NPE_FECHA=" & XDQ(FechaNotaPedido)
            sql = sql & " AND DNP.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND DNP.PTO_CODIGO=D.PTO_CODIGO"
            sql = sql & " AND D.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " ORDER BY DNP.DNP_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(I, 1) = Rec1!PTO_DESCRI
                    grdGrilla.TextMatrix(I, 2) = Rec1!DNP_CANTIDAD
                    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!LIS_PRECIO)
                    grdGrilla.TextMatrix(I, 4) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 5) = Rec1!LNA_DESCRI
                    grdGrilla.TextMatrix(I, 6) = Rec1!DNP_NROITEM
                    I = I + 1
                    Rec1.MoveNext
                Loop
            End If
            Rec1.Close
            '--------------------------------------------------
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
            'chkRemitoSinFactura.SetFocus
        Else
            MsgBox "La Nota de Pedido no existe", vbExclamation, TIT_MSGBOX
            If Rec2.State = 1 Then Rec2.Close
            LimpiarNotaPedido
        End If
    End If
End Sub

Private Sub LimpiarNotaPedido()
    txtNroNotaPedido.Text = ""
    FechaNotaPedido.Text = ""
    'grillaNotaPedido.TextMatrix(0, 1) = ""
    'grillaNotaPedido.TextMatrix(1, 1) = ""
    'grillaNotaPedido.TextMatrix(2, 1) = ""
    txtNroNotaPedido.SetFocus
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

Private Sub txtObservaciones_GotFocus()
    SelecTexto txtObservaciones
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
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
Private Sub CortarCadena(Renglon As Double, Cadena As String)
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
    
        Imprimir 3.2, Renglon, False, cadena1
        Imprimir 3.2, Renglon + 0.5, False, cadena2
        Imprimir 3.2, Renglon + 1, False, cadena3
        Imprimir 3.2, Renglon + 1.5, False, cadena4
        Imprimir 3.2, Renglon + 2, False, cadena5
        Imprimir 3.2, Renglon + 2.5, False, cadena6
        Imprimir 3.2, Renglon + 3, False, cadena7
        Imprimir 3.2, Renglon + 3.5, False, cadena8
    'Else
    '    cadena1 = cadena
    '    MsgBox cadena1
    'End If
    
End Sub

