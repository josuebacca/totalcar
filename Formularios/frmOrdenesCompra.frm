VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmOrdenesCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Compra"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7935
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   555
      Left            =   7125
      TabIndex        =   78
      Top             =   7335
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   555
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7335
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   555
      Left            =   6120
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
      TabIndex        =   11
      Top             =   7335
      Width           =   990
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      Height          =   555
      Left            =   9135
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7335
      Width           =   990
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7155
      Left            =   0
      TabIndex        =   25
      Top             =   120
      Width           =   11145
      _ExtentX        =   19659
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
      TabPicture(0)   =   "frmOrdenesCompra.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FramePedido"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraDatos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmOrdenesCompra.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame8 
         Caption         =   "Transporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8280
         TabIndex        =   79
         Top             =   2425
         Width           =   2745
         Begin VB.TextBox txttransporte 
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
            MaxLength       =   50
            TabIndex        =   5
            Top             =   220
            Width           =   2445
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Empleado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   64
         Top             =   2425
         Width           =   4425
         Begin VB.CommandButton cmdNuevoVendedor 
            Height          =   315
            Left            =   1185
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenesCompra.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Agregar Vendedor"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarVendedor 
            Height          =   315
            Left            =   765
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenesCompra.frx":03C2
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Buscar Vendedor"
            Top             =   240
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
            Left            =   1680
            TabIndex        =   65
            Top             =   240
            Width           =   2685
         End
         Begin VB.TextBox txtNroVendedor 
            Height          =   300
            Left            =   180
            TabIndex        =   3
            Top             =   240
            Width           =   540
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
         Height          =   2070
         Left            =   -74625
         TabIndex        =   44
         Top             =   330
         Width           =   10395
         Begin VB.OptionButton optDef 
            Caption         =   "Definitivos"
            Height          =   195
            Left            =   7920
            TabIndex        =   21
            Top             =   1800
            Width           =   1575
         End
         Begin VB.OptionButton optPend 
            Caption         =   "Pendientes"
            Height          =   195
            Left            =   5280
            TabIndex        =   20
            Top             =   1800
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   3360
            TabIndex        =   19
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CommandButton cmdBuscarVen 
            Height          =   300
            Left            =   4395
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenesCompra.frx":06CC
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Buscar Vendedor"
            Top             =   840
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   300
            Left            =   4395
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenesCompra.frx":09D6
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Buscar Cliente"
            Top             =   375
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3360
            TabIndex        =   16
            Top             =   840
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
            Left            =   4845
            TabIndex        =   49
            Top             =   855
            Width           =   4635
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Empleado"
            Height          =   195
            Left            =   300
            TabIndex        =   13
            Top             =   825
            Width           =   1035
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1380
            Left            =   9705
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenesCompra.frx":0CE0
            Style           =   1  'Graphical
            TabIndex        =   22
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
            Left            =   4845
            MaxLength       =   50
            TabIndex        =   45
            Tag             =   "Descripción"
            Top             =   375
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3360
            MaxLength       =   40
            TabIndex        =   15
            Top             =   375
            Width           =   975
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   300
            TabIndex        =   14
            Top             =   1215
            Width           =   810
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   300
            TabIndex        =   12
            Top             =   435
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   3360
            TabIndex        =   17
            Top             =   1320
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
            Left            =   5970
            TabIndex        =   18
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   56950785
            CurrentDate     =   41098
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   2760
            TabIndex        =   80
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Index           =   0
            Left            =   2535
            TabIndex        =   50
            Top             =   885
            Width           =   735
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4935
            TabIndex        =   48
            Top             =   1380
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2265
            TabIndex        =   47
            Top             =   1365
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
            TabIndex        =   46
            Top             =   420
            Width           =   780
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
         Left            =   3840
         TabIndex        =   30
         Top             =   360
         Width           =   7185
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
            TabIndex        =   62
            Top             =   960
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
            Left            =   960
            MaxLength       =   50
            TabIndex        =   59
            Top             =   1320
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
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   57
            Top             =   960
            Width           =   4860
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
            TabIndex        =   55
            Top             =   1665
            Width           =   1455
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Height          =   315
            Left            =   2415
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenesCompra.frx":3482
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Agregar Cliente"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   1980
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenesCompra.frx":380C
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Buscar Cliente"
            Top             =   270
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
            Left            =   2445
            TabIndex        =   37
            Top             =   1665
            Width           =   3135
         End
         Begin VB.TextBox TxtCodigoCli 
            Height          =   300
            Left            =   960
            MaxLength       =   40
            TabIndex        =   4
            Top             =   270
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
            Left            =   2865
            MaxLength       =   50
            TabIndex        =   24
            Tag             =   "Descripción"
            Top             =   270
            Width           =   4230
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
            TabIndex        =   32
            Top             =   615
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
            Left            =   5640
            TabIndex        =   31
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   180
            TabIndex        =   60
            Top             =   1350
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   150
            TabIndex        =   58
            Top             =   990
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
            Left            =   120
            TabIndex        =   36
            Top             =   315
            Width           =   780
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   210
            TabIndex        =   35
            Top             =   645
            Width           =   675
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   285
            TabIndex        =   34
            Top             =   1710
            Width           =   600
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos"
            Height          =   195
            Left            =   5760
            TabIndex        =   33
            Top             =   1455
            Width           =   765
         End
      End
      Begin VB.Frame FramePedido 
         Caption         =   "Pedido..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2805
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   3720
         Begin VB.TextBox txtNroNotaPedido 
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
            Left            =   1020
            MaxLength       =   8
            TabIndex        =   0
            Top             =   315
            Width           =   1275
         End
         Begin TabDlg.SSTab tabLista 
            Height          =   1215
            Left            =   120
            TabIndex        =   2
            Top             =   1440
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   2143
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "Maquinarias"
            TabPicture(0)   =   "frmOrdenesCompra.frx":3B16
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame6"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Repuestos"
            TabPicture(1)   =   "frmOrdenesCompra.frx":3B32
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame7"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame7 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   -74880
               TabIndex        =   76
               Top             =   360
               Width           =   3255
               Begin VB.ComboBox cboLPrecioC 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   77
                  Top             =   240
                  Width           =   2985
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   120
               TabIndex        =   74
               Top             =   360
               Width           =   3255
               Begin VB.ComboBox cboListaPrecio 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   75
                  Top             =   240
                  Width           =   2985
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   -74880
               TabIndex        =   72
               Top             =   360
               Width           =   3495
               Begin VB.ComboBox cboLPrecioRep 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   73
                  Top             =   240
                  Width           =   3225
               End
            End
         End
         Begin MSComCtl2.DTPicker FechaNotaPedido 
            Height          =   315
            Left            =   1020
            TabIndex        =   1
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   56950785
            CurrentDate     =   41098
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   420
            TabIndex        =   54
            Top             =   1065
            Width           =   540
         End
         Begin VB.Label lblEstadoNota 
            AutoSize        =   -1  'True
            Caption         =   "EST. Orden de Compra"
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
            TabIndex        =   53
            Top             =   1080
            Width           =   1965
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   465
            TabIndex        =   29
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   360
            TabIndex        =   28
            Top             =   345
            Width           =   600
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4440
         Left            =   -74640
         TabIndex        =   23
         Top             =   2520
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7832
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame3 
         Height          =   3870
         Left            =   120
         TabIndex        =   40
         Top             =   3120
         Width           =   10920
         Begin VB.CommandButton cmdBuscarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmOrdenesCompra.frx":3B4E
            Style           =   1  'Graphical
            TabIndex        =   43
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
            Picture         =   "frmOrdenesCompra.frx":3E58
            Style           =   1  'Graphical
            TabIndex        =   42
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
            Picture         =   "frmOrdenesCompra.frx":4162
            Style           =   1  'Graphical
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Producto"
            Top             =   885
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   270
            TabIndex        =   7
            Top             =   495
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3645
            Left            =   120
            TabIndex        =   6
            Top             =   165
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   6429
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
         End
      End
      Begin VB.Frame Frame1 
         Height          =   780
         Left            =   105
         TabIndex        =   63
         Top             =   1920
         Visible         =   0   'False
         Width           =   3720
         Begin VB.CommandButton cmdNuevoRubro 
            Height          =   315
            Left            =   2955
            MaskColor       =   &H000000FF&
            Picture         =   "frmOrdenesCompra.frx":4EE4
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Agregar Condición de Venta"
            Top             =   0
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.ComboBox cboCondicion 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   0
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.CheckBox chkDetalle 
            Alignment       =   1  'Right Justify
            Caption         =   "NP Detallada"
            Height          =   195
            Left            =   0
            TabIndex        =   68
            Top             =   60
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Condición:"
            Height          =   195
            Left            =   1320
            TabIndex        =   71
            Top             =   45
            Visible         =   0   'False
            Width           =   810
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   26
         Top             =   570
         Width           =   1065
      End
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
      Left            =   2040
      TabIndex        =   61
      Top             =   7440
      Width           =   3195
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
      TabIndex        =   52
      Top             =   7440
      Width           =   750
   End
End
Attribute VB_Name = "frmOrdenesCompra"
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
        txtCliente.Text = ""
        txtDesCli.Text = ""
        txtCliente.Enabled = False
        cmdBuscarCli.Enabled = False
    End If
End Sub

Private Sub chkDetalle_Click()
    If chkDetalle.Value = Checked Then
        cboListaPrecio.ListIndex = 0
        cboCondicion.ListIndex = 0
        cboCondicion.Enabled = True
        cmdNuevoRubro.Enabled = True
    Else
        cboListaPrecio.ListIndex = 0
        cboCondicion.ListIndex = -1
        cboCondicion.Enabled = False
        cmdNuevoRubro.Enabled = False
    End If
End Sub

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        FechaDesde.Enabled = True
        FechaHasta.Enabled = True
    Else
        FechaDesde.Value = Null
        FechaHasta.Value = Null
        FechaDesde.Enabled = False
        FechaHasta.Enabled = False
    End If
End Sub
Private Sub chkVendedor_Click()
    If chkVendedor.Value = Checked Then
        txtVendedor.Enabled = True
        cmdBuscarVen.Enabled = True
    Else
        txtVendedor.Text = ""
        txtDesVen.Text = ""
        txtVendedor.Enabled = False
        cmdBuscarVen.Enabled = False
    End If
End Sub

Private Sub cmdAgregarProducto_Click()
    Consulta = 3
    'ABMProducto.CODIGOLISTA = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
    ABMProducto.Show vbModal
    If Consulta <> 4 Then
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        If Trim(ABMProducto.txtcodigo) <> "" Then txtEdit.Text = ABMProducto.txtcodigo
        TxtEdit_KeyDown vbKeyReturn, 0
    End If
    'grdGrilla.SetFocus
    'grdGrilla.row = 1
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
    sql = sql & " FROM ORDEN_COMPRA NP, PROVEEDOR C,VENDEDOR V, LOCALIDAD L, PROVINCIA P "
    sql = sql & " WHERE"
    sql = sql & " NP.PROV_CODIGO=C.PROV_CODIGO AND"
    sql = sql & " C.LOC_CODIGO=L.LOC_CODIGO AND"
    sql = sql & " NP.VEN_CODIGO=V.VEN_CODIGO AND"
    sql = sql & " C.PRO_CODIGO=P.PRO_CODIGO AND"
    sql = sql & " P.PRO_CODIGO=L.PRO_CODIGO "
    If txtCliente.Text <> "" Then sql = sql & " AND NP.PROV_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND NP.OC_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND NP.OC_FECHA<=" & XDQ(FechaHasta)
    If optPend.Value = True Then sql = sql & " AND NP.EST_CODIGO = " & XN("1")
    If optDef.Value = True Then sql = sql & " AND NP.EST_CODIGO = " & XN("3")
    
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
        frmBuscar.grdBuscar.Col = 1
        TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 2
        txtRazSocCli.Text = frmBuscar.grdBuscar.Text
        TxtCodigoCli_LostFocus
    Else
        TxtCodigoCli.SetFocus
    End If
End Sub

Private Sub cmdBuscarProducto_Click()
    Consulta = 3
    If tabLista.Tab = 0 Then
        FrmListadePrecios.tabLista.Tab = 0
        FrmListadePrecios.cboListaPrecio.ListIndex = cboListaPrecio.ListIndex
    Else
        FrmListadePrecios.tabLista.Tab = 1
        FrmListadePrecios.cboLPrecioRep.ListIndex = cboLPrecioC.ListIndex
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
            sql = sql & " ,OC_TRANSP=" & XS(txttransporte)
            
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
                    sql = sql & I & "," 'NRO ITEM
                    sql = sql & XS(grdGrilla.TextMatrix(I, 0)) & "," 'PRODUCTO CODIGO
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
        sql = sql & "PROV_CODIGO,VEN_CODIGO,OC_TRANSP,EST_CODIGO)"
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
        sql = sql & XS(txttransporte) & ","
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
                sql = sql & XN(grdGrilla.TextMatrix(I, 6)) & "," 'NRO ITEM
                sql = sql & XS(grdGrilla.TextMatrix(I, 0)) & "," 'PRODUCTO CODIGO
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
    cmdImprimir_Click
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
    If IsNull(FechaNotaPedido.Value) Then
        MsgBox "La Fecha de la Orden de Compra es requerida", vbExclamation, TIT_MSGBOX
        FechaNotaPedido.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    'NO ES OBLIGATORIO
    If txtNroVendedor.Text = "" Then
        MsgBox "El Vendedor es requerido", vbExclamation, TIT_MSGBOX
        txtNroVendedor.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
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
    If MsgBox("¿Confirma Impresión de la Orden de Compra?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    
    sql = "UPDATE ORDEN_COMPRA"
    sql = sql & " SET EST_CODIGO =" & 3
    sql = sql & " WHERE"
    sql = sql & " OC_NUMERO=" & XN(txtNroNotaPedido)
    sql = sql & " AND OC_FECHA=" & XDQ(FechaNotaPedido)
    DBConn.Execute sql
    
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
    ImprimirOC
End Sub

Public Sub ImprimirOC()
    Dim Renglon As Double
    Dim canttxt As Integer
    Dim w As Integer
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Imprimiendo..."
    
    For w = 1 To 1 'SE IMPRIME POR una sola
      '-----IMPRESION DEL ENCABEZADO------------------
        ImprimirEncabezado
        
      '---- IMPRESION DE LA OC ------------------
        Renglon = 13
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                If grdGrilla.TextMatrix(I, 5) = "MAQUINARIA" Then 'MAQUINARIA
                    Imprimir 2, Renglon, False, "* " & grdGrilla.TextMatrix(I, 2) & " Maq." 'cantidad
                Else
                    If grdGrilla.TextMatrix(I, 5) = "REPUESTOS" Then 'REPUESTOS
                        Imprimir 2, Renglon, False, "* " & grdGrilla.TextMatrix(I, 2) & " Rep." 'cantidad
                    End If
                End If
                Imprimir 4, Renglon, False, grdGrilla.TextMatrix(I, 0) 'codigo
                Imprimir 6.5, Renglon, False, grdGrilla.TextMatrix(I, 1) 'descri
                Renglon = Renglon + 0.5
                    
                    
            End If
        Next I
        '-----OBSERVACIONES---------------------
        If txtNombreVendedor.Text <> "" Then
            Imprimir 12, Renglon + 3, False, "Por TOTALCAR"
            Imprimir 14, Renglon + 3, False, UCase(Trim(txtNombreVendedor.Text))
        End If
            
        Printer.EndDoc
    Next w
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub
Private Function FechaLarga(Fecha As String) As String
Dim DIA As Integer
Dim MES As Integer
Dim año As Integer
Dim nombredelmes As String
    
    'VerificarFecha = True
    
    If Val(Fecha) = 0 Then Exit Function
    
    If Len(Fecha) < 8 Then
        Beep
        MsgBox "Fecha Incompleta !", vbExclamation, TIT_MSGBOX
        'VerificarFecha = False
        Exit Function
    End If
    
    DIA = Val(Mid(Trim(Fecha), 1, 2))
    MES = Val(Mid(Trim(Fecha), 4, 2))
    año = Val(Mid(Trim(Fecha), 7, 4))
    
    If MES = 1 Then nombredelmes = "Enero"
    If MES = 2 Then nombredelmes = "Febrero"
    If MES = 3 Then nombredelmes = "Marzo"
    If MES = 4 Then nombredelmes = "Abril"
    If MES = 5 Then nombredelmes = "Mayo"
    If MES = 6 Then nombredelmes = "Junio"
    If MES = 7 Then nombredelmes = "Julio"
    If MES = 8 Then nombredelmes = "Agosto"
    If MES = 9 Then nombredelmes = "Septiembre"
    If MES = 10 Then nombredelmes = "Octubre"
    If MES = 11 Then nombredelmes = "Noviembre"
    If MES = 12 Then nombredelmes = "Diciembre"
    
    FechaLarga = DIA & " de " & nombredelmes & " de " & año
End Function

Public Sub ImprimirEncabezado()
 '-----------IMPRIME EL ENCABEZADO DE LA FACTURA-------------------
    Dim Fecha As String
    Dim año As String
    'año = String(4, Year(FechaFactura))
    año = Year(FechaNotaPedido)
    Fecha = FechaLarga(FechaNotaPedido)
    Imprimir 12, 6, False, "Pilar " & Fecha
'    Imprimir 15.7, 12, False, Format(Day(FechaNotaPedido), "00") & "/"
'    Imprimir 16.75, 12, False, Format(Month(FechaNotaPedido), "00") & "/"
'    Imprimir 17.8, 12, False, Mid(año, 3, 2)
    
    
    
    Imprimir 8.5, 8, False, txtRazSocCli.Text
    Imprimir 2, 10, False, "Atención: SECCIÓN REPUESTO"
    Imprimir 2, 11, False, "Despachar por Transporte " & UCase(Trim(txttransporte.Text)) & " lo siguiente:"
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
   FramePedido.Enabled = True
   fraDatos.Enabled = True
   TxtCodigoCli.Text = ""
   chkDetalle.Value = Unchecked
   TxtCodigoCli.Text = ""
   txtRazSocCli.Text = ""
   txtNombreVendedor.Text = ""
   txtNroVendedor.Text = ""
   FechaNotaPedido.Value = Date
   txtNroNotaPedido.Text = ""
   lblEstadoNota.Caption = ""
   lblEstado.Caption = ""
   tabDatos.Tab = 0
   Call BuscoEstado(1, lblEstadoNota)
   cmdGrabar.Enabled = True
   CmdBorrar.Enabled = True
   txtNroNotaPedido.Text = BuscoUltimoRenito
   If txtNroNotaPedido.Text = "" Then
        txtNroNotaPedido.Text = "00000001"
   End If
   FechaNotaPedido = Date
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
        Set frmOrdenesCompra = Nothing
        Unload Me
    End If
End Sub

Private Sub FechaNotaPedido_LostFocus()
    If IsNull(FechaNotaPedido.Value) Then
        FechaNotaPedido.Value = Date
        'If txtNroNotaPedido.Text = "" Then txtNroNotaPedido.SetFocus
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
Private Function BuscoUltimoRenito() As String
    'ACA BUSCA EL NUMERO DE ORDEN DE COMPRA SIGUIENTE AL ULTIMO CARGADO
    sql = "SELECT MAX(OC_NUMERO) + 1 AS ULTIMO"
    sql = sql & " FROM ORDEN_COMPRA"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        'txtNroSucursal.Text = Sucursal
        BuscoUltimoRenito = Format(rec!Ultimo, "00000000")
    End If
    rec.Close
End Function

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)
    grdGrilla.FormatString = "Código|Descripción|Cantidad|Precio|Rubro|Linea|Orden"
    grdGrilla.ColWidth(0) = 1000 'CODIGO
    grdGrilla.ColWidth(1) = 5900 'DESCRIPCION
    grdGrilla.ColWidth(2) = 1000 'CANTIDAD
    grdGrilla.ColWidth(3) = 0 'PRECIO
    grdGrilla.ColWidth(4) = 2100 'RUBRO
    grdGrilla.ColWidth(5) = 2100 'LINEA
    grdGrilla.ColWidth(6) = 0    'ORDEN
    grdGrilla.Cols = 7
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
    GrdModulos.FormatString = ">Número|^Fecha|Cliente|Domicilio|Localidad"
    GrdModulos.ColWidth(0) = 1200
    GrdModulos.ColWidth(1) = 1000
    GrdModulos.ColWidth(2) = 3200
    GrdModulos.ColWidth(3) = 3200
    GrdModulos.ColWidth(4) = 3200
    GrdModulos.Rows = 1
    '------------------------------------
    
    txtNroNotaPedido.Text = BuscoUltimoRenito
    If txtNroNotaPedido.Text = "" Then
        txtNroNotaPedido.Text = "00000001"
    End If
    'CARGO COMBO LISTA DE PRECIOS
    CargoCboListaPrecio
    'CARGO EL COMBO DE LISTA DE PRECIOS DE REPUESTOS
    CargoCboLPrecioRep
    
    'CARGO CONDICIONES DE PAGO
    
    LlenarComboFormaPago
    'CARGO COMBO REPRESENTADA
    'CargoComboRepresentada
    '-----------------
    cboListaPrecio.ListIndex = 0
    cboCondicion.ListIndex = -1
    cboCondicion.Enabled = False
    cmdNuevoRubro.Enabled = False
    lblEstado.Caption = ""
    Call BuscoEstado(1, lblEstadoNota)
    tabDatos.Tab = 0
    FechaNotaPedido = Date
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
            cboLPrecioC.AddItem rec!LIS_DESCRI
            cboLPrecioC.ItemData(cboLPrecioC.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboLPrecioC.ListIndex = 0
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
        txtNroNotaPedido.Text = Format(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), "00000000")
        FechaNotaPedido.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
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
        txtProvincia.Text = ""
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
    If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Or ActiveControl.Name = "CmdSalir" Then Exit Sub
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
            txtProvincia.Text = rec!PRO_DESCRI
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
    'If grdGrilla.Col = 0 Then KeyAscii = CarTexto(KeyAscii)
    If grdGrilla.Col = 2 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 3 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
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
                    grdGrilla.Col = 4
                    'grdGrilla.Text = Valido_Importe(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2))
                    'grdGrilla.Col = 4
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
                    'grdGrilla.Col = 3
                    'grdGrilla.Text = Valido_Importe(Trim(rec!LIS_PRECIO))
                    grdGrilla.Col = 4
                    grdGrilla.Text = Trim(rec!RUB_DESCRI)
                    grdGrilla.Col = 5
                    grdGrilla.Text = Trim(rec!LNA_DESCRI)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
                    grdGrilla.Col = 2
                End If
'                    If BuscoRepetetidos(CLng(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), grdGrilla.RowSel) = False Then
'                     grdGrilla.Col = 0
'                     grdGrilla_KeyDown vbKeyDelete, 0
'                    End If
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
        FechaNotaPedido.Value = Date
    End If
End Sub

Private Sub txtNroNotaPedido_GotFocus()
     FechaNotaPedido.Value = Null
End Sub

Private Sub txtNroNotaPedido_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaPedido_LostFocus()
Set Rec1 = New ADODB.Recordset

    If ActiveControl.Name = "CmdSalir" Or ActiveControl.Name = "chkCliente" _
       Or ActiveControl.Name = "cmdNuevo" Or ActiveControl.Name = "cmdBuscarNotaPedido" Then Exit Sub
    
    If txtNroNotaPedido.Text <> "" Then
       sql = "SELECT O.*, E.EST_DESCRI"
        sql = sql & " FROM ORDEN_COMPRA O, ESTADO_DOCUMENTO E"
        sql = sql & " WHERE O.OC_NUMERO=" & XN(txtNroNotaPedido)
        If FechaNotaPedido.Value <> "" Then
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
            FechaNotaPedido.Value = Rec2!OC_FECHA
            txtNroVendedor.Text = Rec2!VEN_CODIGO
            txttransporte.Text = IIf(IsNull(Rec2!OC_TRANSP), "", Rec2!OC_TRANSP)
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
                'cmdGrabar.Enabled = False
                'cmdBorrar.Enabled = False
                FramePedido.Enabled = False
                fraDatos.Enabled = False
                grdGrilla.SetFocus
                'cmdImprimir.Enabled = False
            Else
                cmdGrabar.Enabled = True
                CmdBorrar.Enabled = True
                FramePedido.Enabled = True
                fraDatos.Enabled = True
            End If
            'CARGO ESTADO
            
            If lblEstadoNota.Caption = "PENDIENTE" Then
                cmdImprimir.Enabled = False
                cmdGrabar.Enabled = True
            End If
            If lblEstadoNota.Caption = "ANULADO" Then
                cmdImprimir.Enabled = False
                cmdGrabar.Enabled = False
            End If
            If lblEstadoNota.Caption = "DEFINITIVO" Then
               cmdImprimir.Enabled = True
               'cmdGrabar.Enabled = False
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
                    grdGrilla.TextMatrix(I, 4) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 5) = Rec1!LNA_DESCRI
                    grdGrilla.TextMatrix(I, 6) = Rec1!DOC_NROITEM
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
