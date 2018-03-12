VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRemitoProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remito de Proveedores..."
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8130
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   570
      Left            =   8020
      TabIndex        =   8
      Top             =   7500
      Width           =   990
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   570
      Left            =   10080
      TabIndex        =   10
      Top             =   7500
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   570
      Left            =   6990
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7500
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   570
      Left            =   9050
      TabIndex        =   9
      Top             =   7500
      Width           =   990
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7380
      Left            =   50
      TabIndex        =   28
      Top             =   50
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   13018
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
      TabPicture(0)   =   "frmRemitoProveedor.frx":0000
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
      TabPicture(1)   =   "frmRemitoProveedor.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CheckBox chkNotaPedido 
         Caption         =   "Recupera datos de Orden de Compra"
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Width           =   3135
      End
      Begin VB.Frame freCliente 
         Height          =   1935
         Left            =   4050
         TabIndex        =   55
         Top             =   1060
         Width           =   6975
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
            TabIndex        =   73
            Top             =   900
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
            TabIndex        =   72
            Top             =   900
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
            TabIndex        =   63
            Top             =   1520
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
            TabIndex        =   62
            Top             =   585
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
            TabIndex        =   61
            Tag             =   "Descripción"
            Top             =   240
            Width           =   3990
         End
         Begin VB.TextBox TxtCodigoCli 
            Height          =   300
            Left            =   930
            MaxLength       =   40
            TabIndex        =   5
            Top             =   240
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
            TabIndex        =   60
            Top             =   1520
            Width           =   3135
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   1920
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoProveedor.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Buscar Cliente"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Height          =   315
            Left            =   2385
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoProveedor.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Agregar Cliente"
            Top             =   240
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
            TabIndex        =   57
            Top             =   1520
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
            TabIndex        =   56
            Top             =   1200
            Width           =   4620
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos"
            Height          =   195
            Left            =   5730
            TabIndex        =   69
            Top             =   1320
            Width           =   765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   255
            TabIndex        =   68
            Top             =   1560
            Width           =   600
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   180
            TabIndex        =   67
            Top             =   603
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
            Left            =   90
            TabIndex        =   66
            Top             =   285
            Width           =   780
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   921
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   150
            TabIndex        =   64
            Top             =   1239
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
         Height          =   705
         Left            =   4050
         TabIndex        =   47
         Top             =   360
         Width           =   6990
         Begin VB.CommandButton cmdBuscarNotaPedido 
            Height          =   315
            Left            =   2685
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoProveedor.frx":06CC
            Style           =   1  'Graphical
            TabIndex        =   49
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
         Begin MSComCtl2.DTPicker FechaNotaPedido 
            Height          =   315
            Left            =   4440
            TabIndex        =   12
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   56950785
            CurrentDate     =   41098
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   3855
            TabIndex        =   50
            Top             =   255
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   750
            TabIndex        =   48
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
         Height          =   2640
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   3920
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
            Height          =   330
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   1
            Top             =   240
            Width           =   1005
         End
         Begin VB.ComboBox cboStock 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   2505
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
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   0
            Top             =   240
            Width           =   555
         End
         Begin TabDlg.SSTab tabLista 
            Height          =   1215
            Left            =   120
            TabIndex        =   74
            Top             =   1320
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   2143
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "Accesorios"
            TabPicture(0)   =   "frmRemitoProveedor.frx":09D6
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame2"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Repuestos"
            TabPicture(1)   =   "frmRemitoProveedor.frx":09F2
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame4"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame2 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   120
               TabIndex        =   77
               Top             =   360
               Width           =   3495
               Begin VB.ComboBox cboListaPrecio 
                  Height          =   315
                  Left            =   600
                  Style           =   2  'Dropdown List
                  TabIndex        =   78
                  Top             =   240
                  Width           =   2505
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   -74880
               TabIndex        =   75
               Top             =   360
               Width           =   3495
               Begin VB.ComboBox cboLPrecioRep 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   76
                  Top             =   240
                  Width           =   3225
               End
            End
         End
         Begin MSComCtl2.DTPicker FechaRemito 
            Height          =   315
            Left            =   1320
            TabIndex        =   2
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   56950785
            CurrentDate     =   41098
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
            TabIndex        =   70
            Top             =   1050
            Width           =   1215
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Stock:"
            Height          =   210
            Left            =   795
            TabIndex        =   53
            Top             =   0
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   765
            TabIndex        =   51
            Top             =   705
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   660
            TabIndex        =   46
            Top             =   285
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   720
            TabIndex        =   45
            Top             =   1035
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
         Height          =   2070
         Left            =   -74600
         TabIndex        =   36
         Top             =   540
         Width           =   10410
         Begin VB.Frame Frame1 
            Caption         =   "Estado Remito"
            Height          =   495
            Left            =   840
            TabIndex        =   21
            Top             =   1440
            Width           =   8535
            Begin VB.OptionButton optPen 
               Caption         =   "Pendientes"
               Height          =   195
               Left            =   1200
               TabIndex        =   22
               Top             =   200
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.OptionButton optDef 
               Caption         =   "Definitivos"
               Height          =   195
               Left            =   3075
               TabIndex        =   23
               Top             =   200
               Width           =   1455
            End
            Begin VB.OptionButton optAnu 
               Caption         =   "Anulados"
               Height          =   195
               Left            =   4845
               TabIndex        =   24
               Top             =   200
               Width           =   1455
            End
            Begin VB.OptionButton optTod 
               Caption         =   "Todos"
               Height          =   195
               Left            =   6600
               TabIndex        =   25
               Top             =   200
               Width           =   1455
            End
         End
         Begin VB.CommandButton cmdBuscarVendedor 
            Height          =   315
            Left            =   4290
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoProveedor.frx":0A0E
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Buscar Vendedor"
            Top             =   660
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4290
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoProveedor.frx":0D18
            Style           =   1  'Graphical
            TabIndex        =   43
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
            TabIndex        =   41
            Top             =   675
            Width           =   4620
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Empleado"
            Height          =   195
            Left            =   900
            TabIndex        =   15
            Top             =   765
            Width           =   1035
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1395
            Left            =   9660
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoProveedor.frx":1022
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Buscar "
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   555
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
            TabIndex        =   37
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
            Top             =   1095
            Width           =   810
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   878
            TabIndex        =   14
            Top             =   435
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   3240
            TabIndex        =   19
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   56950785
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   6240
            TabIndex        =   20
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   56950785
            CurrentDate     =   41098
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Empleado:"
            Height          =   195
            Index           =   0
            Left            =   2415
            TabIndex        =   42
            Top             =   705
            Width           =   750
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   5175
            TabIndex        =   40
            Top             =   1140
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2145
            TabIndex        =   39
            Top             =   1125
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
            Left            =   2385
            TabIndex        =   38
            Top             =   300
            Width           =   780
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4170
         Left            =   -74640
         TabIndex        =   27
         Top             =   2715
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7355
         _Version        =   393216
         Cols            =   13
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame3 
         Height          =   4275
         Left            =   120
         TabIndex        =   31
         Top             =   2950
         Width           =   10935
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   1455
            MaxLength       =   60
            TabIndex        =   13
            Top             =   3870
            Width           =   8850
         End
         Begin VB.CommandButton cmdBuscarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoProveedor.frx":37C4
            Style           =   1  'Graphical
            TabIndex        =   35
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
            Picture         =   "frmRemitoProveedor.frx":3ACE
            Style           =   1  'Graphical
            TabIndex        =   34
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
            Picture         =   "frmRemitoProveedor.frx":3DD8
            Style           =   1  'Graphical
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Producto"
            Top             =   885
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   270
            TabIndex        =   32
            Top             =   525
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3615
            Left            =   90
            TabIndex        =   6
            Top             =   165
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   6376
            _Version        =   393216
            Rows            =   3
            Cols            =   6
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   210
            TabIndex        =   52
            Top             =   3915
            Width           =   1110
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   29
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
      TabIndex        =   71
      Top             =   7680
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
      TabIndex        =   44
      Top             =   7695
      Width           =   750
   End
End
Attribute VB_Name = "frmRemitoProveedor"
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

Private Sub Check1_Click()

End Sub

Private Sub chkCliente_Click()
    If chkCliente.Value = Checked Then
        txtCliente.Enabled = True
        cmdBuscarCli.Enabled = True
    Else
        txtCliente.Enabled = False
        cmdBuscarCli.Enabled = False
        txtCliente.Text = ""
    End If
End Sub

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        FechaDesde.Enabled = True
        FechaHasta.Enabled = True
    Else
        FechaDesde.Enabled = False
        FechaHasta.Enabled = False
        FechaDesde.Value = Null
        FechaHasta.Value = Null
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
        txtVendedor.Text = ""
    End If
End Sub

Private Sub cmdAgregarProducto_Click()
'    ABMProducto.Show vbModal
'    grdGrilla.SetFocus
'    grdGrilla.row = 1
    
    Consulta = 3
    'ABMProducto.CODIGOLISTA = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
    ABMProducto.Show vbModal
    If Consulta <> 4 Then
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        If Trim(ABMProducto.txtcodigo) <> "" Then txtEdit.Text = ABMProducto.txtcodigo
        TxtEdit_KeyDown vbKeyReturn, 0
    End If
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Select Case TipoBusquedaDoc
    
    
    Case 1 'BUSCA REMITOS
        
        sql = "SELECT RC.*, C.PROV_RAZSOC,C.PROV_DOMICI,L.LOC_DESCRI,P.PRO_DESCRI"
        sql = sql & " FROM REMITO_PROVEEDOR RC,Proveedor C, LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE"
        sql = sql & "  RC.PROV_CODIGO=C.PROV_CODIGO"
        sql = sql & "  AND C.LOC_CODIGO=L.LOC_CODIGO"
        sql = sql & "  AND C.PRO_CODIGO=P.PRO_CODIGO"
        sql = sql & "  AND L.PRO_CODIGO=P.PRO_CODIGO"
        If txtCliente.Text <> "" Then sql = sql & " AND RC.PROV_CODIGO=" & XN(txtCliente)
        If Not IsNull(FechaDesde) Then sql = sql & " AND RC.RPR_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND RC.RPR_FECHA<=" & XDQ(FechaHasta)
        If optPen.Value = True Then
            sql = sql & " AND RC.EST_CODIGO = 1 "
        End If
        If optDef.Value = True Then
            sql = sql & " AND RC.EST_CODIGO = 3 "
        End If
        If optAnu.Value = True Then
            sql = sql & " AND RC.EST_CODIGO = 2 "
        End If
        sql = sql & " ORDER BY RC.RPR_NUMERO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem Format(rec!RPR_SUCURSAL, "0000") & "-" & Format(rec!RPR_NUMERO, "00000000") _
                                & Chr(9) & rec!RPR_FECHA _
                                & Chr(9) & rec!PROV_RAZSOC & Chr(9) & rec!PROV_DOMICI _
                                & Chr(9) & rec!LOC_DESCRI & Chr(9) & rec!PRO_DESCRI _
                                & Chr(9) & rec!EST_CODIGO _
                                & Chr(9) & rec!OC_NUMERO & Chr(9) & rec!OC_FECHA _
                                & Chr(9) & rec!RPR_OBSERVACION & Chr(9) & rec!STK_CODIGO _
                                & Chr(9) & rec!RPR_SINFAC & Chr(9) & rec!PROV_CODIGO
                rec.MoveNext
            Loop
            GrdModulos.SetFocus
        Else
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        End If
        
    Case 2 'BUSCA ORDEN DE COMPRA
        
        sql = "SELECT NP.OC_NUMERO, NP.OC_FECHA, C.PROV_RAZSOC, "
        sql = sql & " C.PROV_CODIGO,C.PROV_DOMICI,L.LOC_DESCRI,P.PRO_DESCRI"
        sql = sql & " FROM ORDEN_COMPRA NP, Proveedor C, LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE"
        sql = sql & " NP.PROV_CODIGO=C.PROV_CODIGO"
        sql = sql & " AND L.LOC_CODIGO=C.LOC_CODIGO"
        sql = sql & " AND P.PRO_CODIGO=C.PRO_CODIGO"
        sql = sql & " AND P.PRO_CODIGO=L.PRO_CODIGO"
        sql = sql & " AND NP.EST_CODIGO = 3"
        If txtCliente.Text <> "" Then sql = sql & " AND NP.PROV_CODIGO=" & XN(txtCliente)
        If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
        If Not IsNull(FechaDesde) Then sql = sql & " AND NP.OC_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND NP.OC_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY OC_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem rec!OC_NUMERO & Chr(9) & rec!OC_FECHA _
                                & Chr(9) & rec!PROV_RAZSOC & Chr(9) & rec!PROV_DOMICI _
                                & Chr(9) & rec!LOC_DESCRI & Chr(9) & rec!PRO_DESCRI _
                                & Chr(9) & rec!PROV_CODIGO
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
    frmBuscar.TipoBusqueda = 5
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 1
        TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 2
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
    
    Dim nlista As Integer
    
    Consulta = 3
    
    If tabLista.Tab = 0 Then
        FrmListadePrecios.tabLista.Tab = 0
        FrmListadePrecios.cboListaPrecio.ListIndex = cboListaPrecio.ListIndex
    Else
        FrmListadePrecios.tabLista.Tab = 1
        FrmListadePrecios.cboLPrecioRep.ListIndex = cboLPrecioRep.ListIndex
    End If
    'Call BuscaCodigoProxItemData(nlista, FrmListadePrecios.cbodescri)
    FrmListadePrecios.Show vbModal
    If Consulta <> 4 Then
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        If Trim(FrmListadePrecios.GrdModulos.Text) <> "" Then txtEdit.Text = FrmListadePrecios.GrdModulos.Text
        TxtEdit_KeyDown vbKeyReturn, 0
    End If
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

Private Sub cmdGrabar_Click()
    If ValidarRemito = False Then Exit Sub
    
    On Error GoTo HayErrorRemito
    
    If MsgBox("¿Confirma Remito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    DBConn.BeginTrans
    sql = "SELECT * FROM REMITO_Proveedor"
    sql = sql & " WHERE RPR_NUMERO=" & XN(txtNroRemito)
    sql = sql & " AND RPR_SUCURSAL=" & XN(txtNroSucursal)
    sql = sql & " AND PROV_CODIGO =" & XN(TxtCodigoCli)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = True Then 'NUEVO REMITO
        sql = "INSERT INTO REMITO_PROVEEDOR"
        sql = sql & " (RPR_NUMERO,RPR_SUCURSAL,RPR_FECHA,OC_NUMERO,"
        sql = sql & "OC_FECHA,RPR_OBSERVACION,"
        sql = sql & "EST_CODIGO,RPR_NUMEROTXT,STK_CODIGO, PROV_CODIGO)"
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
                sql = "INSERT INTO DETALLE_REMITO_PROVEEDOR"
                sql = sql & " (RPR_NUMERO,RPR_SUCURSAL,PROV_CODIGO,RPR_FECHA,DRPR_NROITEM,"
                sql = sql & "PTO_CODIGO,DRPR_CANTIDAD,DRPR_PRECIO,DRPR_DETALLE)"
                sql = sql & " VALUES ("
                sql = sql & XN(txtNroRemito) & ","
                sql = sql & XN(txtNroSucursal) & ","
                sql = sql & XN(TxtCodigoCli) & ","
                sql = sql & XDQ(FechaRemito) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 6)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(I, 0), True) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(I, 1)) & ")"
                DBConn.Execute sql
            End If
        Next
         'ACTUALIZO EL STOCK CUANDO EL REMITO ES DEFINITIVO (STOCK PENDIENTE)
         For I = 1 To grdGrilla.Rows - 1
             If grdGrilla.TextMatrix(I, 0) <> "" Then
                     sql = "UPDATE DETALLE_STOCK"
                     sql = sql & " SET"
                     sql = sql & " DST_STKFIS = DST_STKFIS + " & XN(grdGrilla.TextMatrix(I, 2))
                     sql = sql & " WHERE STK_CODIGO= 1"
                     sql = sql & " AND PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "'"
                     DBConn.Execute sql
             End If
         Next
         
        'CAMBIO ESTADO DE LA ORDEN DE COMPRA (LE PONGO DEFINITIVO)
        If chkNotaPedido.Value = 1 Then
            sql = "UPDATE ORDEN_COMPRA SET EST_CODIGO=4"
            sql = sql & " WHERE"
            sql = sql & " OC_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND OC_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
        End If
        
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO AL REMITO
        'sql = "UPDATE PARAMETROS SET NRO_REMITO=" & XN(txtNroRemito)
        'DBConn.Execute sql
        
        DBConn.CommitTrans
    Else
        ' modifico el Remito
        'If MsgBox("Confirma modificar el Remito Nro.: " & Trim(txtNroSucursal.Text) & "-" & Trim(txtNroNotaPedido.Text) & " ", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            sql = "UPDATE REMITO_PROVEEDOR"
            sql = sql & "  SET RPR_OBSERVACION=" & XS(txtObservaciones)
            sql = sql & " ,STK_CODIGO=" & cboStock.ItemData(cboStock.ListIndex)
            sql = sql & " ,RPR_NUMEROTXT=" & XS(Format(txtNroRemito.Text, "00000000"))
            
            sql = sql & " WHERE"
            sql = sql & " RPR_NUMERO=" & XN(txtNroRemito)
            sql = sql & " AND RPR_FECHA=" & XDQ(FechaRemito)
            sql = sql & " AND PROV_CODIGO=" & XN(TxtCodigoCli)
            DBConn.Execute sql
            
            sql = "DELETE FROM DETALLE_REMITO_PROVEEDOR"
            sql = sql & " WHERE RPR_NUMERO=" & XN(txtNroRemito)
            sql = sql & " AND RPR_SUCURSAL=" & XN(txtNroSucursal)
            sql = sql & " AND RPR_FECHA=" & XDQ(FechaRemito)
            sql = sql & " AND PROV_CODIGO=" & XN(TxtCodigoCli)
            DBConn.Execute sql
            
            For I = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(I, 0) <> "" Then
                    sql = "INSERT INTO DETALLE_REMITO_PROVEEDOR"
                    sql = sql & " (RPR_NUMERO,RPR_SUCURSAL,PROV_CODIGO,RPR_FECHA,DRPR_NROITEM,PTO_CODIGO,"
                    sql = sql & "DRPR_CANTIDAD,DRPR_PRECIO,DRPR_DETALLE)"
                    sql = sql & " VALUES ("
                    sql = sql & XN(txtNroRemito) & ","
                    sql = sql & XN(txtNroSucursal) & ","
                    sql = sql & XN(TxtCodigoCli) & ","
                    sql = sql & XDQ(FechaRemito) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 6)) & "," 'NRO ITEM
                    sql = sql & XS(grdGrilla.TextMatrix(I, 0), True) & "," 'PRODUCTO CODIGO
                    sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & "," 'CANTIDAD
                    sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & "," 'PRECIO
                    sql = sql & XS(grdGrilla.TextMatrix(I, 1)) & ")"  ' DETALLE
                    DBConn.Execute sql
                End If
            Next
            DBConn.CommitTrans
            End If
        'End If
        
           'ACTUALIZO EL STOCK CUANDO EL REMITO ES DEFINITIVO (STOCK PENDIENTE)
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                    sql = "UPDATE DETALLE_STOCK"
                    sql = sql & " SET"
                    sql = sql & " DST_STKFIS = DST_STKFIS + " & XN(grdGrilla.TextMatrix(I, 2))
                    sql = sql & " WHERE STK_CODIGO= 1"
                    sql = sql & " AND PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "'"
                    DBConn.Execute sql
            End If
        Next
        
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    
    If MsgBox("¿Desea Facturar el Remito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        'frmFacturaCliente.TipoBusquedaDoc = 2 'BUSCA REMITOS
        frmfacturaproveedor.txtRemSuc = Format(txtNroSucursal.Text, "0000")
        frmfacturaproveedor.txtNroRemito = Format(txtNroRemito.Text, "00000000")
        frmfacturaproveedor.TxtCodigoCli = TxtCodigoCli.Text
        
        frmfacturaproveedor.Show vbModal
        
        'frmFacturaCliente.GrdModulos.ColWidth(0) = 0 'TIPO FACTURA
        'frmFacturaCliente.tabDatos.Tab = 1
        'frmFacturaCliente.Show vbModal
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
    
    If txtNroSucursal.Text = "" Then
        MsgBox "El Número de Sucursal es requerido", vbExclamation, TIT_MSGBOX
        txtNroSucursal.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    If txtNroRemito.Text = "" Then
        MsgBox "El Número del Remito es requerido", vbExclamation, TIT_MSGBOX
        txtNroRemito.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    If IsNull(FechaRemito.Value) Then
        MsgBox "La Fecha del Remito es requerida", vbExclamation, TIT_MSGBOX
        FechaRemito.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    If chkNotaPedido.Value = 1 Then
        If txtNroNotaPedido.Text = "" Then
            MsgBox "El número de Orden de Compra es requerido", vbExclamation, TIT_MSGBOX
            txtNroNotaPedido.SetFocus
            ValidarRemito = False
            Exit Function
        End If
        If FechaNotaPedido.Value = "" Then
            MsgBox "La Fecha de la Orden de Compra es requerida", vbExclamation, TIT_MSGBOX
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
        MsgBox "Debe ingresar un Proveedor", vbExclamation, TIT_MSGBOX
        TxtCodigoCli.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    
    
    ValidarRemito = True
End Function
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
   FechaNotaPedido.Value = ""
   txtNroNotaPedido.Text = ""
 '  chkRemitoSinFactura.Value = Unchecked
  ' txtConcepto.Text = ""
   txtNroRemito.Text = ""
   lblEstadoRemito.Caption = ""
   txtObservaciones.Text = ""
   lblEstado.Caption = ""
   cmdGrabar.Enabled = True
   freRemito.Enabled = True
   freCliente.Enabled = True
   
    'BUSCO EL NUMERO DE REMITO QUE CORRESPONDE
    'txtNroRemito.Text = BuscoUltimoRenito
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
    FechaRemito.Value = Date
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
    ABMProveedor.Show vbModal
    TxtCodigoCli.SetFocus
End Sub
Private Sub cmdQuitarProducto_Click()
    If MsgBox("Seguro que desea quitar el Producto: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 1), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
    End If
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmRemitoProveedor = Nothing
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
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    grdGrilla.FormatString = "Código|Descripción|Cantidad|Precio|Rubro|Linea|Orden"
    grdGrilla.ColWidth(0) = 1500 'CODIGO
    grdGrilla.ColWidth(1) = 5900 'DESCRIPCION
    grdGrilla.ColWidth(2) = 1000 'CANTIDAD
    grdGrilla.ColWidth(3) = 1100 'PRECIO
    grdGrilla.ColWidth(4) = 2100 'RUBRO
    grdGrilla.ColWidth(5) = 2100 'LINEA
    grdGrilla.ColWidth(6) = 0    'ORDEN
    grdGrilla.Cols = 7
    grdGrilla.Rows = 1
    For I = 2 To 25
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & (I - 1)
    Next
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "^Número|^Fecha|Proveedor|Domicilio|Localidad|Provincia|Cod Estado|NP NUMERO|NP FECHA|OBSERVACIONES|" _
                              & "STOCK|REMITO SIN FACTURA|CODIGOPROVEEDOR"
    GrdModulos.ColWidth(0) = 1500 'NUMERO
    GrdModulos.ColWidth(1) = 1000 'FECHA
    GrdModulos.ColWidth(2) = 3200 'PROVEEDOR
    GrdModulos.ColWidth(3) = 3200 'DOMICILIO
    GrdModulos.ColWidth(4) = 3200 'Localidad
    GrdModulos.ColWidth(5) = 3200 'Provincia
    GrdModulos.ColWidth(6) = 0    'COD ESTADO
    GrdModulos.ColWidth(7) = 0    'NOTA PEDIDO NUMERO
    GrdModulos.ColWidth(8) = 0    'NOTA PEDIDO FECHA
    GrdModulos.ColWidth(9) = 0   'OBSERVACIONES
    GrdModulos.ColWidth(10) = 0   'STOCK
    GrdModulos.ColWidth(11) = 0   'REMITO SIN FACTURAS
    GrdModulos.ColWidth(12) = 0   'CODIGOPROVEEDOR
    
    GrdModulos.Rows = 1
    '------------------------------------
    'grillaNotaPedido.ColWidth(0) = 950
    'grillaNotaPedido.ColWidth(1) = 5300
    'grillaNotaPedido.TextMatrix(0, 0) = "    Cliente:"
    'grillaNotaPedido.TextMatrix(1, 0) = " Sucursal:"
    'grillaNotaPedido.TextMatrix(2, 0) = "Vendedor:"
    '------------------------------------
    lblEstado.Caption = ""
    'CARGO EL COMBO DE LISTA DE PRECIOS
    CargoCboListaPrecio
    
    'CARGO EL COMBO DE LISTA DE PRECIOS DE REPUESTOS
    CargoCboLPrecioRep
    
    
    'BUSCO EL NUMERO DE REMITO QUE CORRESPONDE
    'txtNroRemito.Text = BuscoUltimoRenito
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRemito) 'ESTADO PENDIENTE
    VEstadoRemito = 1
    FechaRemito.Value = Date
    'CARGO COMBO STOCK
    CargaCboStock
    'PONGO ENABLE LOS DATOS DE LA FACTURA DE TERCEROS

    'txtConcepto.Enabled = False
    TipoBusquedaDoc = 1 'ESTO ES PARA BUSCAR REMITOS(1), (2)PARA BUSCAR orden de compra
    tabDatos.Tab = 0
    
    tabLista.Tab = 0
    
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

Private Sub CargoCboListaPrecio() '' Lista de Precios de ACCESORIOS
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 2"   '6: ACCESORIOS
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
    sql = sql & " AND P.LNA_CODIGO = 1"   '1: Repuestos
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
        Select Case TipoBusquedaDoc
        Case 1 'BUSCA REMITOS
            lblEstado.Caption = "Buscando..."
            Screen.MousePointer = vbHourglass
            Set Rec1 = New ADODB.Recordset
            
            txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4)
            txtNroRemito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)
            FechaRemito.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
            'CARGO EL ESTADO
            Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6)), lblEstadoRemito)
            VEstadoRemito = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6))
            
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
            
            
            txtNroNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 7)
            FechaNotaPedido.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
            
            
            'BUSCO STOCK
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 10) <> "" Then
                Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 10)), cboStock)
            Else
                cboStock.ListIndex = 0
            End If
            txtObservaciones.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
            
            TxtCodigoCli.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 12)
            TxtCodigoCli_LostFocus
            
            
        '----BUSCO DETALLE DEL REMITO------------------
            
            sql = "SELECT DRPR.*, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM DETALLE_REMITO_PROVEEDOR DRPR, PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE DRPR.RPR_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4))
            sql = sql & " AND DRPR.RPR_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8))
            sql = sql & " AND DRPR.RPR_FECHA=" & XDQ(GrdModulos.TextMatrix(GrdModulos.RowSel, 1))
            sql = sql & " AND PROV_CODIGO = " & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 12))
            
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
                    grdGrilla.TextMatrix(I, 4) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 5) = Rec1!LNA_DESCRI
                    grdGrilla.TextMatrix(I, 6) = Rec1!DRPR_NROITEM
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
            FechaNotaPedido.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
            tabDatos.Tab = 0
            txtNroNotaPedido_LostFocus
            TxtCodigoCli.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
            TxtCodigoCli_LostFocus
        End Select
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub MaskEdBox1_Change()

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
    'LimpiarBusqueda
    If Me.Visible = True Then chkCliente.SetFocus
    If TipoBusquedaDoc = 1 Then
        frameBuscar.Caption = "Buscar Remito por..."
        chkVendedor.Enabled = False
        txtVendedor.Enabled = False
    Else
        frameBuscar.Caption = "Buscar Ordenes de Compra a facturar por..."
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
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscoCondicionIVA = rec!IVA_DESCRI
    Else
        BuscoCondicionIVA = ""
    End If
    rec.Close
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
        sql = "SELECT C.PROV_RAZSOC,C.PROV_DOMICI,C.PROV_CUIT,C.IVA_CODIGO,C.PROV_INGBRU,"
        sql = sql & "L.LOC_DESCRI,P.PRO_DESCRI,L.LOC_CODPOS"
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
        
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtRazSocCli.Text = ""
            TxtCodigoCli.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    'If grdGrilla.Col = 0 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 2 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    If grdGrilla.Col = 3 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
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
               If lblEstadoRemito.Caption = "PENDIENTE" Then
                    Screen.MousePointer = vbHourglass
                    sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI, P.PTO_PRECIOC, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI,P.PTO_STKMIN"
                    sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, TIPO_PRESENTACION RE"
                    sql = sql & " WHERE"
                    If grdGrilla.Col = 0 Then
                        sql = sql & " P.PTO_CODIGO LIKE '" & txtEdit & "'"
                    Else
                        sql = sql & " P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
                    End If
                    sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
                    sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                    sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
                    'sql = sql & " AND P.PTO_ESTADO=1"
                    '*********
                    
                Else
                    sql = "SELECT DRC.PTO_CODIGO,DRC.DRC_DETALLE, DRC.DRC_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI"
                    sql = sql & " FROM DETALLE_REMITO_PROVEEDOR DRC, PRODUCTO P,LINEAS L,RUBROS R,TIPO_PRESENTACION RE"
                    sql = sql & " WHERE"
                    sql = sql & " DRC.PTO_CODIGO = P.PTO_CODIGO AND "
                    If grdGrilla.Col = 0 Then
                        sql = sql & " DRC.PTO_CODIGO LIKE '" & txtEdit.Text & "'"
                    Else
                        sql = sql & " DRC.DRC_DETALLE LIKE '" & Trim(txtEdit) & "%'"
                    End If
                    sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
                    sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                    sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
                End If
            Else  ' Busca en un Lista de Precios
                sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, P.PTO_PRECIOC, R.RUB_DESCRI, L.LNA_DESCRI"
                sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, TIPO_PRESENTACION T "
                sql = sql & " WHERE"
                If grdGrilla.Col = 0 Then
                    sql = sql & " P.PTO_CODIGO LIKE '" & txtEdit & "'"
                Else
                    sql = sql & " P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
                End If
                    sql = sql & " AND P.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
                    'sql = sql & " AND P.PTO_CODIGO=D.PTO_CODIGO"
                    sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
                    sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
                    sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                    sql = sql & " AND T.TPRE_CODIGO=P.TPRE_CODIGO"
                    'sql = sql & " AND P.PTO_ESTADO=1"
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
                     If lblEstadoRemito.Caption = "PENDIENTE" Then
                        grdGrilla.Col = 1
                        grdGrilla.Text = Trim(rec!PTO_DESCRI)
                        grdGrilla.Col = 3
                        If cboListaPrecio.ListIndex = 0 Then
                            grdGrilla.Text = Valido_Importe(Trim(rec!PTO_PRECIOC))
                        Else
                            grdGrilla.Text = Valido_Importe(Trim(rec!PTO_PRECIOC))
                        End If
                    Else
                        grdGrilla.Col = 1
                        grdGrilla.Text = Trim(rec!DRPR_DETALLE)
                        grdGrilla.Col = 3
                        grdGrilla.Text = Valido_Importe(Trim(rec!DRPR_PRECIO))
                    End If
                    
                    
                    grdGrilla.Col = 4
                    grdGrilla.Text = Trim(rec!RUB_DESCRI)
                    grdGrilla.Col = 5
                    grdGrilla.Text = Trim(rec!LNA_DESCRI)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
                    grdGrilla.Col = 2
                    'CONTROLAR SI EL ARTICULO NO TIENE STOCK MINIMO
                    If rec!LNA_DESCRI = "REPUESTOS" Then
                        If rec!PTO_STKMIN = 0 Or rec!PTO_STKMIN = Null Then
                            MsgBox "El repuesto no tiene cargado el Stock Minimo!", vbExclamation, TIT_MSGBOX
                        End If
                    End If
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
'            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
'                Set Rec1 = New ADODB.Recordset
'                sql = "SELECT P.PTO_STKMIN, DS.DST_STKFIS, DS.DST_STKPEN"
'                sql = sql & " FROM PRODUCTO P, DETALLE_STOCK DS"
'                sql = sql & " WHERE P.PTO_CODIGO=" & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 0))
'                sql = sql & " AND P.PTO_CODIGO=DS.PTO_CODIGO"
'                sql = sql & " AND STK_CODIGO=" & cboStock.ItemData(cboStock.ListIndex)
'
'                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'                If Rec1.EOF = False Then
'                    If (CInt(Rec1!DST_STKFIS) - CInt(txtEdit.Text)) < CInt(Rec1!PTO_STKMIN) Then
'                        MsgBox "El producto esta por debajo del Stock Minimo" & Chr(13) & Chr(13) & _
'                        " Stock Minimo=" & Rec1!PTO_STKMIN & Chr(13) & _
'                        " Stock Pendiente=" & Rec1!DST_STKPEN & Chr(13) & _
'                        " Stock Fisico=" & Rec1!DST_STKFIS & Chr(13) & _
'                        " Stock Fisico - Cantidad=" & (CInt(Rec1!DST_STKFIS) - CInt(txtEdit.Text)), vbExclamation, TIT_MSGBOX
'                    End If
'                End If
'                Rec1.Close
'            End If
            
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
End Sub

Private Function BuscoRepetetidos(Codigo As String, Linea As Integer) As Boolean
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
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
        sql = sql & " FROM ORDEN_COMPRA NP, ESTADO_DOCUMENTO E"
        sql = sql & " WHERE NP.OC_NUMERO=" & XN(txtNroNotaPedido)
        If FechaNotaPedido.Value <> "" Then
            sql = sql & " AND NP.OC_FECHA=" & XDQ(FechaNotaPedido)
        End If
        sql = sql & " AND NP.EST_CODIGO=E.EST_CODIGO"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec2.EOF = False Then
            If Rec2.RecordCount > 1 Then
                MsgBox "Hay mas de una Orden de Compra con el Número: " & txtNroNotaPedido.Text, vbInformation, TIT_MSGBOX
                Rec2.Close
                cmdBuscarNotaPedido_Click
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Buscando..."
            
            'CARGO CABECERA DE LA Orden de Compra
            FechaNotaPedido.Value = Rec2!OC_FECHA
            'grillaNotaPedido.TextMatrix(0, 1) = BuscoCliente(Rec2!PROV_CODIGO)
            'grillaNotaPedido.TextMatrix(1, 1) = BuscoSucursal(Rec2!SUC_CODIGO, Rec2!PROV_CODIGO)
            'grillaNotaPedido.TextMatrix(2, 1) = BuscoVendedor(Rec2!VEN_CODIGO)
            'lblEstadoNotaPedido.Caption = "Estado: " & Rec2!EST_DESCRI
            TxtCodigoCli.Text = Rec2!PROV_CODIGO
            TxtCodigoCli_LostFocus
            If Rec2!EST_CODIGO <> 1 And Rec2!EST_CODIGO <> 3 Then
                MsgBox "La Orden de Compra número: " & txtNroNotaPedido.Text & Chr(13) & Chr(13) & _
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
            
        '-----BUSCO LOS DATOS DEL DETALLE DE LA Orden de Compra---------
            sql = "SELECT DNP.*,P.PTO_DESCRI, D.LIS_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM DETALLE_ORDEN_COMPRA DNP, PRODUCTO P, RUBROS R, LINEAS L, DETALLE_LISTA_PRECIO D"
            sql = sql & " WHERE DNP.OC_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND DNP.OC_FECHA=" & XDQ(FechaNotaPedido)
            sql = sql & " AND DNP.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND DNP.PTO_CODIGO=D.PTO_CODIGO"
            sql = sql & " AND D.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " ORDER BY DNP.DOC_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(I, 1) = Rec1!PTO_DESCRI
                    grdGrilla.TextMatrix(I, 2) = IIf(IsNull(Rec1!DOC_CANTIDAD), 1, Rec1!DOC_CANTIDAD)
                    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!LIS_PRECIO)
                    grdGrilla.TextMatrix(I, 4) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 5) = Rec1!LNA_DESCRI
                    grdGrilla.TextMatrix(I, 6) = Rec1!DOC_NROITEM
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
            MsgBox "La Orden de Compra no existe", vbExclamation, TIT_MSGBOX
            If Rec2.State = 1 Then Rec2.Close
            LimpiarNotaPedido
        End If
    End If
End Sub

Private Sub LimpiarNotaPedido()
    txtNroNotaPedido.Text = ""
    FechaNotaPedido.Value = ""
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

Private Function BuscoProveedor(Codigo As String) As String
        sql = "SELECT PROV_RAZSOC FROM PROVEEDOR"
        sql = sql & " WHERE PROV_CODIGO=" & XN(Codigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            BuscoProveedor = rec!PROV_RAZSOC
        Else
            BuscoProveedor = "No se encontro el Proveedor"
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

Private Sub txtNroRemito_Change()
    'SelecTexto txtNroRemito
End Sub
Private Sub txtNroRemito_GotFocus()
    SelecTexto txtNroRemito
End Sub

'Private Sub txtNroRemito_KeyPress(KeyAscii As Integer)
'    If txtNroRemito.Text = "" Then
'        txtNroRemito.Text = Sucursal
'    Else
'        txtNroRemito.Text = Format(txtNroRemito.Text, "00000000")
'    End If
'End Sub

Private Sub txtNroRemito_LostFocus()
    If txtNroRemito.Text = "" Then
        'txtNroRemito.Text = Sucursal
    Else
        txtNroRemito.Text = Format(txtNroRemito.Text, "00000000")
    End If
End Sub

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
