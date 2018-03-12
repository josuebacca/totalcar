VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNotaDePedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuesto"
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
      Height          =   555
      Left            =   7125
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   7335
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   555
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7335
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   555
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7335
      Width           =   990
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   555
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7335
      Width           =   990
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      Height          =   555
      Left            =   9135
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7335
      Width           =   990
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7275
      Left            =   0
      TabIndex        =   21
      Top             =   30
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   12832
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
      TabPicture(0)   =   "frmNotaDePedido.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FramePedido"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraDatos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmNotaDePedido.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
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
         TabIndex        =   41
         Top             =   570
         Width           =   10395
         Begin VB.CommandButton cmdBuscarVen 
            Height          =   300
            Left            =   4395
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Buscar Vendedor"
            Top             =   840
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   300
            Left            =   4395
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Buscar Cliente"
            Top             =   375
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3360
            TabIndex        =   15
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
            TabIndex        =   46
            Top             =   855
            Width           =   4635
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   300
            TabIndex        =   12
            Top             =   825
            Width           =   1035
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1380
            Left            =   9705
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   18
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
            TabIndex        =   42
            Tag             =   "Descripción"
            Top             =   375
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3360
            MaxLength       =   40
            TabIndex        =   14
            Top             =   375
            Width           =   975
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   300
            TabIndex        =   13
            Top             =   1215
            Width           =   810
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   300
            TabIndex        =   11
            Top             =   435
            Width           =   855
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   3360
            TabIndex        =   16
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61931521
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   5970
            TabIndex        =   17
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61931521
            CurrentDate     =   41098
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Index           =   0
            Left            =   2535
            TabIndex        =   47
            Top             =   885
            Width           =   735
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4935
            TabIndex        =   45
            Top             =   1380
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2265
            TabIndex        =   44
            Top             =   1365
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
            Left            =   2745
            TabIndex        =   43
            Top             =   420
            Width           =   525
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   " Datos del Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2450
         Left            =   4080
         TabIndex        =   26
         Top             =   320
         Width           =   7065
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
            TabIndex        =   63
            Top             =   1080
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
            TabIndex        =   60
            Top             =   1440
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
            TabIndex        =   58
            Top             =   1080
            Width           =   4740
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
            TabIndex        =   52
            Top             =   1785
            Width           =   1455
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Height          =   315
            Left            =   2415
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":2DEE
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Agregar Cliente"
            Top             =   390
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   1980
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":3178
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Buscar Cliente"
            Top             =   390
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
            TabIndex        =   33
            Top             =   1785
            Width           =   3135
         End
         Begin VB.TextBox TxtCodigoCli 
            Height          =   300
            Left            =   960
            MaxLength       =   40
            TabIndex        =   2
            Top             =   390
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
            TabIndex        =   20
            Tag             =   "Descripción"
            Top             =   390
            Width           =   4110
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
            TabIndex        =   28
            Top             =   735
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
            TabIndex        =   27
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   180
            TabIndex        =   61
            Top             =   1470
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   150
            TabIndex        =   59
            Top             =   1110
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
            Left            =   360
            TabIndex        =   32
            Top             =   435
            Width           =   525
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   210
            TabIndex        =   31
            Top             =   765
            Width           =   675
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   285
            TabIndex        =   30
            Top             =   1830
            Width           =   600
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos"
            Height          =   195
            Left            =   5640
            TabIndex        =   29
            Top             =   1575
            Width           =   765
         End
      End
      Begin VB.Frame FramePedido 
         Caption         =   "Prespuesto..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2450
         Left            =   120
         TabIndex        =   23
         Top             =   320
         Width           =   3960
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
            Left            =   960
            MaxLength       =   8
            TabIndex        =   0
            Top             =   260
            Width           =   1275
         End
         Begin TabDlg.SSTab tabLista 
            Height          =   1215
            Left            =   120
            TabIndex        =   64
            Top             =   1150
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   2143
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "Maquinarias"
            TabPicture(0)   =   "frmNotaDePedido.frx":3482
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame5"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Repuestos"
            TabPicture(1)   =   "frmNotaDePedido.frx":349E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame6"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame6 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   -74880
               TabIndex        =   74
               Top             =   360
               Width           =   3495
               Begin VB.ComboBox cboLPrecioRep 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   75
                  Top             =   240
                  Width           =   3225
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   120
               TabIndex        =   67
               Top             =   360
               Width           =   3495
               Begin VB.ComboBox cboListaPrecio 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   68
                  Top             =   240
                  Width           =   3225
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   -74880
               TabIndex        =   65
               Top             =   360
               Width           =   3495
               Begin VB.ComboBox cboLPrecioRepviejo 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   66
                  Top             =   240
                  Width           =   3225
               End
            End
         End
         Begin MSComCtl2.DTPicker FechaNotaPedido 
            Height          =   315
            Left            =   960
            TabIndex        =   1
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   61931521
            CurrentDate     =   41098
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   420
            TabIndex        =   51
            Top             =   900
            Width           =   540
         End
         Begin VB.Label lblEstadoNota 
            AutoSize        =   -1  'True
            Caption         =   "EST. NOTA PEDIDO"
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
            TabIndex        =   50
            Top             =   900
            Width           =   1770
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   465
            TabIndex        =   25
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   360
            TabIndex        =   24
            Top             =   300
            Width           =   600
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4560
         Left            =   -74640
         TabIndex        =   19
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
      Begin VB.Frame Frame1 
         Height          =   780
         Left            =   105
         TabIndex        =   53
         Top             =   1575
         Visible         =   0   'False
         Width           =   10920
         Begin VB.TextBox txtNroVendedor 
            Height          =   300
            Left            =   900
            TabIndex        =   72
            Top             =   0
            Visible         =   0   'False
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
            Left            =   0
            TabIndex        =   71
            Top             =   345
            Visible         =   0   'False
            Width           =   3165
         End
         Begin VB.CommandButton cmdBuscarVendedor 
            Height          =   315
            Left            =   1725
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":34BA
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Buscar Vendedor"
            Top             =   0
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoVendedor 
            Height          =   315
            Left            =   2145
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":37C4
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Agregar Vendedor"
            Top             =   0
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.CheckBox chkDetalle 
            Alignment       =   1  'Right Justify
            Caption         =   "NP Detallada"
            Height          =   195
            Left            =   105
            TabIndex        =   3
            Top             =   345
            Width           =   1260
         End
         Begin VB.ComboBox cboListaPrecioviejo 
            Height          =   315
            Left            =   8340
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   285
            Width           =   2505
         End
         Begin VB.ComboBox cboCondicion 
            Height          =   315
            Left            =   2445
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   285
            Width           =   4185
         End
         Begin VB.CommandButton cmdNuevoRubro 
            Height          =   315
            Left            =   6660
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":3B4E
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Agregar Condición de Venta"
            Top             =   285
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   105
            TabIndex        =   73
            Top             =   45
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Lista de Precios:"
            Height          =   195
            Left            =   7110
            TabIndex        =   56
            Top             =   330
            Width           =   1170
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Condición:"
            Height          =   195
            Left            =   1665
            TabIndex        =   55
            Top             =   330
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4500
         Left            =   105
         TabIndex        =   36
         Top             =   2700
         Width           =   11040
         Begin VB.CommandButton cmdPrecio 
            Caption         =   "$"
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
            Left            =   10515
            Style           =   1  'Graphical
            TabIndex        =   88
            ToolTipText     =   "Actualizar Precios"
            Top             =   1200
            Width           =   390
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   1605
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   81
            Top             =   4155
            Width           =   8865
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
            Left            =   1605
            TabIndex        =   80
            Top             =   3800
            Width           =   1350
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
            Left            =   9120
            TabIndex        =   79
            Top             =   3800
            Width           =   1350
         End
         Begin VB.TextBox txtPorcentajeIva 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4290
            TabIndex        =   78
            Top             =   3800
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
            Left            =   6810
            TabIndex        =   77
            Top             =   3800
            Width           =   1155
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
            Left            =   10590
            TabIndex        =   76
            Top             =   3780
            Width           =   255
         End
         Begin VB.CommandButton cmdBuscarProducto 
            Height          =   330
            Left            =   10515
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaDePedido.frx":3ED8
            Style           =   1  'Graphical
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "Buscar Producto"
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdAgregarProducto 
            Height          =   330
            Left            =   10515
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaDePedido.frx":41E2
            Style           =   1  'Graphical
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   "Agregar Producto"
            Top             =   530
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdQuitarProducto 
            Height          =   330
            Left            =   10515
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaDePedido.frx":44EC
            Style           =   1  'Graphical
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Producto"
            Top             =   865
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   270
            TabIndex        =   37
            Top             =   495
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3495
            Left            =   120
            TabIndex        =   6
            Top             =   165
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   6165
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
            AllowUserResizing=   3
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   360
            TabIndex        =   86
            Top             =   4200
            Width           =   1110
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   765
            TabIndex        =   85
            Top             =   3855
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   8655
            TabIndex        =   84
            Top             =   3850
            Width           =   405
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "% I.V.A.:"
            Height          =   195
            Left            =   3630
            TabIndex        =   83
            Top             =   3855
            Width           =   600
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   6180
            TabIndex        =   82
            Top             =   3855
            Width           =   570
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   22
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Presupuesto"
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
      Left            =   2880
      TabIndex        =   62
      Top             =   7560
      Width           =   2655
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
      TabIndex        =   49
      Top             =   7680
      Width           =   750
   End
End
Attribute VB_Name = "frmNotaDePedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
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
        If MsgBox("Seguro que desea eliminar el Presupuesto Nro.: " & Trim(txtNroNotaPedido), vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
           On Error GoTo Seclavose
           
           sql = "SELECT P.EST_CODIGO, E.EST_DESCRI "
           sql = sql & " FROM NOTA_PEDIDO P, ESTADO_DOCUMENTO E"
           sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido)
           sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
           sql = sql & " AND P.EST_CODIGO=E.EST_CODIGO"
           rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
           
           If rec.EOF = False Then
                If rec!EST_CODIGO <> 1 Then
                    MsgBox "El Presupuesto no puede ser eliminado," & Chr(13) & _
                           " ya que esta en estado: " & Trim(rec!EST_DESCRI), vbExclamation, TIT_MSGBOX
                    rec.Close
                    Exit Sub
                End If
           End If
           rec.Close
            lblEstado.Caption = "Eliminando..."
            Screen.MousePointer = vbHourglass
            
            sql = "DELETE FROM DETALLE_NOTA_PEDIDO"
            sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
            
            sql = "DELETE FROM NOTA_PEDIDO"
            sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
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
    
    sql = "SELECT NP.*, C.CLI_RAZSOC,C.CLI_DOMICI,L.LOC_DESCRI"
    sql = sql & " FROM NOTA_PEDIDO NP, CLIENTE C, LOCALIDAD L "
    sql = sql & " WHERE"
    sql = sql & " NP.CLI_CODIGO=C.CLI_CODIGO AND"
    sql = sql & " C.LOC_CODIGO=L.LOC_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND NP.NPE_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND NP.NPE_FECHA<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY NP.NPE_NUMERO, NP.NPE_FECHA"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem Format(rec!NPE_NUMERO, "00000000") & Chr(9) & rec!NPE_FECHA _
                            & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!CLI_DOMICI _
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

Private Sub cmdBuscarProducto_Click()
'    grdGrilla.SetFocus
'    frmBuscar.TipoBusqueda = 2
'    frmBuscar.CodListaPrecio = 0
'    frmBuscar.TxtDescriB.Text = ""
'    frmBuscar.Show vbModal
'
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
    If MsgBox("¿Confirma el Presupuesto?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    On Error GoTo HayErrorNota
    
    DBConn.BeginTrans
    sql = "SELECT * FROM NOTA_PEDIDO"
    sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido)
    sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = False Then
        If MsgBox("Seguro que modificar el Presupuesto Nro.: " & Trim(txtNroNotaPedido), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            sql = "UPDATE NOTA_PEDIDO"
            sql = sql & " SET CLI_CODIGO=" & XN(TxtCodigoCli)
            sql = sql & " ,VEN_CODIGO=" & XN(txtNroVendedor)
            sql = sql & " ,NPE_IVA=" & XN(txtPorcentajeIva)
            sql = sql & " ,NPE_SUBTOTAL=" & XN(txtSubtotal)
            sql = sql & " ,NPE_TOTAL=" & XN(txtTotal)
            sql = sql & ", NPE_OBSERV =" & XS(txtObservaciones)
            If chkDetalle.Value = Checked Then
                sql = sql & " ,FPG_CODIGO=" & cboCondicion.ItemData(cboCondicion.ListIndex)
            Else
                sql = sql & " ,FPG_CODIGO=NULL"
            End If
            
            sql = sql & " WHERE"
            sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
            
            sql = "DELETE FROM DETALLE_NOTA_PEDIDO"
            sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
            
            For I = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(I, 0) <> "" Then
                    sql = "INSERT INTO DETALLE_NOTA_PEDIDO"
                    sql = sql & " (NPE_NUMERO,NPE_FECHA,DNP_NROITEM,PTO_CODIGO,"
                    sql = sql & "DNP_CANTIDAD,DNP_PRECIO,DNP_IMPORTE)"
                    sql = sql & " VALUES ("
                    sql = sql & XN(txtNroNotaPedido) & ","
                    sql = sql & XDQ(FechaNotaPedido) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 7)) & "," 'NRO ITEM
                    sql = sql & XS(grdGrilla.TextMatrix(I, 0)) & "," 'PRODUCTO CODIGO
                    sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & "," 'CANTIDAD
                    sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & "," 'PRECIO
                    sql = sql & XN(grdGrilla.TextMatrix(I, 4)) & ")" 'IMPORTE
                    DBConn.Execute sql
                End If
            Next
            DBConn.CommitTrans
        End If
        
    Else 'PEDIDO- PRESUPUESTO NUEVO
        sql = "INSERT INTO NOTA_PEDIDO"
        sql = sql & " (NPE_NUMERO,NPE_FECHA,CLI_CODIGO,"
        sql = sql & "VEN_CODIGO,FPG_CODIGO,NPE_NUMEROTXT,NPE_IVA,NPE_SUBTOTAL,NPE_TOTAL,NPE_OBSERV,EST_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(txtNroNotaPedido) & ","
        sql = sql & XDQ(FechaNotaPedido) & ","
        sql = sql & XN(TxtCodigoCli) & ","
        sql = sql & XN(txtNroVendedor) & ","
        If chkDetalle.Value = Checked Then
            sql = sql & cboCondicion.ItemData(cboCondicion.ListIndex) & ","
        Else
            sql = sql & "NULL,"
        End If
        sql = sql & XS(Format(txtNroNotaPedido.Text, "00000000")) & ","
        'sql = sql & XS(txtNroNotaPedido) & ","
        sql = sql & XN(txtPorcentajeIva) & ","
        sql = sql & XN(txtSubtotal) & ","
        sql = sql & XN(txtTotal) & ","
        sql = sql & XS(txtObservaciones) & ","
        sql = sql & "1)" 'ESTADO PENDIENTE
        DBConn.Execute sql
           
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                sql = "INSERT INTO DETALLE_NOTA_PEDIDO"
                sql = sql & " (NPE_NUMERO,NPE_FECHA,DNP_NROITEM,PTO_CODIGO,"
                sql = sql & " DNP_CANTIDAD,DNP_PRECIO,DNP_IMPORTE)"
                sql = sql & " VALUES ("
                sql = sql & XN(txtNroNotaPedido) & ","
                sql = sql & XDQ(FechaNotaPedido) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 7)) & "," 'NRO ITEM
                sql = sql & XS(grdGrilla.TextMatrix(I, 0)) & "," 'PRODUCTO CODIGO
                sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & "," 'CANTIDAD
                sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & "," 'PRECIO
                sql = sql & XN(grdGrilla.TextMatrix(I, 4)) & ")" 'IMPORTE
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
        MsgBox "El número de Presupuesto es requerido", vbExclamation, TIT_MSGBOX
        txtNroNotaPedido.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    If IsNull(FechaNotaPedido.Value) Then
        MsgBox "La Fecha del Presupuesto es requerida", vbExclamation, TIT_MSGBOX
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
        MsgBox "El Cliente es requerido", vbExclamation, TIT_MSGBOX
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
 If MsgBox("¿Confirma Impresión del Presupuesto?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
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
    Dim w As Integer
    For w = 1 To 2 'SE IMPRIME UNA SOLA COPIA
      '-----IMPRESION DEL ENCABEZADO------------------
        ImprimirEncabezado
        
      '---- IMPRESION DEL PRESUPUESTO ------------------
        Renglon = 10
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                Imprimir 1, Renglon, False, grdGrilla.TextMatrix(I, 0)  'codigo
                canttxt = 0
                If Len(grdGrilla.TextMatrix(I, 1)) < 36 Then
                    Imprimir 3.2, Renglon, False, grdGrilla.TextMatrix(I, 1) 'descripcion
                Else
                     CortarCadena Renglon, grdGrilla.TextMatrix(I, 1)
                    
                    
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
                    Imprimir 10.6, Renglon, False, grdGrilla.TextMatrix(I, 2) 'cantidad
                    Imprimir 11.6, Renglon, False, grdGrilla.TextMatrix(I, 3) 'precio
                    'Imprimir 15, Renglon, False, grdGrilla.TextMatrix(I, 4) 'bonificacion
                'End If
                Imprimir 13.4, Renglon, False, Valido_Importe(grdGrilla.TextMatrix(I, 4)) 'importe
                Renglon = Renglon + (canttxt * 0.5) + 0.5
            
                'IMPRIMO DATOS DE MAQUINARIA
'                If grdGrilla.TextMatrix(I, 10) = 6 Then  'SI LINEA MAQUINARIA (hay que ver si es tractor o sembradora)
'                    sql = "SELECT * FROM PRODUCTO "
'                    sql = sql & "WHERE PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "'"
'                    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'                    If Rec1.EOF = False Then
'                        If Not IsNull(Rec1!PTO_TIPO) Then
'                            Imprimir 2, Renglon + 3, False, "Tipo.............: " & IIf(IsNull(Rec1!PTO_TIPO), "", Rec1!PTO_TIPO)
'                        End If
'                        If Not IsNull(Rec1!PTO_TIPMOD) Then
'                            Imprimir 11, Renglon + 3, False, "Modelo.........: " & IIf(IsNull(Rec1!PTO_TIPMOD), "", Rec1!PTO_TIPMOD)
'                        End If
'                        If Not IsNull(Rec1!PTO_TRACCI) Then
'                            Imprimir 2, Renglon + 3.4, False, "Tracción......: " & IIf(IsNull(Rec1!PTO_TRACCI), "", Rec1!PTO_TRACCI)
'                        End If
'                        If Not IsNull(Rec1!PTO_TIPO) Then
'                            Imprimir 11, Renglon + 3.4, False, IIf(Rec1!PTO_CABINA = 1, "Con Cabina", "Sin Cabina")
'                        End If
'                        If Not IsNull(Rec1!PTO_MOTMAR) Then
'                            Imprimir 2, Renglon + 3.8, False, "Motor Marca: " & IIf(IsNull(Rec1!PTO_MOTMAR), "", Rec1!PTO_MOTMAR)
'                        End If
'                        If Not IsNull(Rec1!PTO_MOTMOD) Then
'                            Imprimir 11, Renglon + 3.8, False, "Modelo.........: " & IIf(IsNull(Rec1!PTO_MOTMOD), "", Rec1!PTO_MOTMOD)
'                        End If
'                        If Not IsNull(Rec1!PTO_ASPIRA) Then
'                            Imprimir 2, Renglon + 4.2, False, "Aspiración...: " & IIf(IsNull(Rec1!PTO_ASPIRA), "", Rec1!PTO_ASPIRA)
'                        End If
'                        If Not IsNull(Rec1!PTO_MOTNRO) Then
'                            Imprimir 11, Renglon + 4.2, False, "Motor Nro.....: " & IIf(IsNull(Rec1!PTO_MOTNRO), "", Rec1!PTO_MOTNRO)
'                        End If
'                        If Not IsNull(Rec1!PTO_CHASIS) Then
'                            Imprimir 2, Renglon + 4.6, False, "Chasis Nro..: " & IIf(IsNull(Rec1!PTO_CHASIS), "", Rec1!PTO_CHASIS)
'                        End If
'                        If Not IsNull(Rec1!PTO_SERIE) Then
'                            Imprimir 11, Renglon + 4.6, False, "Serie.............: " & IIf(IsNull(Rec1!PTO_SERIE), "", Rec1!PTO_SERIE)
'                        End If
'                        If Not IsNull(Rec1!PTO_NEUMDE) Then
'                            Imprimir 2, Renglon + 5, False, "Neum. Del...: " & IIf(IsNull(Rec1!PTO_NEUMDE), "", Rec1!PTO_NEUMDE)
'                        End If
'                        If Not IsNull(Rec1!PTO_NEDECA) Then
'                            Imprimir 11, Renglon + 5, False, "Cantidad.......: " & IIf(IsNull(Rec1!PTO_NEDECA), "", Rec1!PTO_NEDECA)
'                        End If
'                        If Not IsNull(Rec1!PTO_NEUMTR) Then
'                            Imprimir 2, Renglon + 5.4, False, "Neum. Tra...: " & IIf(IsNull(Rec1!PTO_NEUMTR), "", Rec1!PTO_NEUMTR)
'                        End If
'                        If Not IsNull(Rec1!PTO_NETRCA) Then
'                            Imprimir 11, Renglon + 5.4, False, "Cantidad.......: " & IIf(IsNull(Rec1!PTO_NETRCA), "", Rec1!PTO_NETRCA)
'                        End If
'                        If Not IsNull(Rec1!PTO_TIPO) Then
'                            Imprimir 2, Renglon + 5.8, False, IIf(Rec1!PTO_KITCON = 1, "Con Kit Confort", "Sin Kit Confort")
'                        End If
'                        If Not IsNull(Rec1!PTO_SALHID) Then
'                            Imprimir 11, Renglon + 5.8, False, "Salida Hidr....: " & IIf(IsNull(Rec1!PTO_SALHID), "", Rec1!PTO_SALHID)
'                        End If
'                        If Not IsNull(Rec1!PTO_POSARA) Then
'                            Imprimir 2, Renglon + 6.2, False, "Posic. Aran..: " & IIf(IsNull(Rec1!PTO_POSARA), "", Rec1!PTO_POSARA)
'                        End If
'                        If Not IsNull(Rec1!PTO_CERFAB) Then
'                            Imprimir 2, Renglon + 6.2, False, "Cert. Fabrica: " & IIf(IsNull(Rec1!PTO_CERFAB), "", Rec1!PTO_CERFAB)
'                        End If
'                        If Not IsNull(Rec1!PTO_OPCION1) Then
'                            Imprimir 2, Renglon + 6.6, False, IIf(IsNull(Rec1!PTO_OPCION1), "", Rec1!PTO_OPCION1)
'                        End If
'                        If Not IsNull(Rec1!PTO_OPCION2) Then
'                            Imprimir 11, Renglon + 6.6, False, IIf(IsNull(Rec1!PTO_OPCION2), "", Rec1!PTO_OPCION2)
'                        End If
'                    End If
                   ' Rec1.Close
                'End If
            End If
        Next I
            '-----OBSERVACIONES---------------------
            'If txtObservaciones.Text <> "" Then
                Imprimir 1, 25, False, "PRESUPUESTO VALIDO POR EL TERMINO DE 15 DIAS"
                'CortarCadena Renglon + 9.5, Trim(txtObservaciones)
                'Imprimir 5, & Trim(txtObservaciones.Text)
            'End If
            'Imprimir 0, 16.5, True, "texto de bajo del detalle"
            '-------------IMPRIMO TOTALES--------------------
            Imprimir 11.8, 21.5, True, "Subtotal: $ " & txtSubtotal.Text
            
'            If txtPorcentajeBoni.Text <> "" Then
'                If chkBonificaEnPesos.Value = Checked Then
'                    Imprimir 3.5, 15, True, "$" & txtPorcentajeBoni.Text
'                    Imprimir 4.4, 15.5, True, txtImporteBoni.Text
'                Else
'                    Imprimir 3.5, 15, True, "%" & txtPorcentajeBoni.Text
'                    Imprimir 4.4, 15.5, True, txtImporteBoni.Text
'                End If
'                Imprimir 7.1, 15.5, True, txtSubTotalBoni.Text
'            End If
            'Imprimir 16, 19.9, True, txtSubtotal.Text
             'If txtPorcentajeBoni.Text <> "" Then
                 'Imprimir 16.8, 22.5, True, txtSubTotalBoni.Text
            ' Else
            '     Imprimir 16.8, 22.5, True, txtSubtotal.Text
            ' End If
            
            Imprimir 11.6, 23.5, True, "IVA: " & txtPorcentajeIva.Text & "  $ " & txtImporteIva.Text
            'Imprimir 12.2, 23.5, True, " $" & txtImporteIva.Text
            Imprimir 12, 25.5, True, "TOTAL: $ " & txtTotal.Text
            
        Printer.EndDoc
    Next w
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Public Sub ImprimirEncabezado()
 '-----------IMPRIME EL ENCABEZADO DE LA FACTURA-------------------
    Dim año As String
    'año = String(4, Year(FechaFactura))
    año = Year(FechaNotaPedido)
    Imprimir 12.5, 5.3, False, "Fecha: " & Format(Day(FechaNotaPedido), "00") & "/" & Format(Month(FechaNotaPedido), "00") & "/" & Mid(año, 3, 2)

    'Imprimir 16.55, 4.3, False, Format(Month(FechaNotaPedido), "00")
    'Imprimir 17.6, 4.3, False, Mid(año, 3, 2)
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CUIT,C.CLI_INGBRU, L.LOC_DESCRI"
    sql = sql & ", P.PRO_DESCRI,CI.IVA_DESCRI,C.IVA_CODIGO"
    sql = sql & " FROM CLIENTE C, LOCALIDAD L, NOTA_PEDIDO RC,"
    sql = sql & " PROVINCIA P, CONDICION_IVA CI"
    sql = sql & " WHERE RC.NPE_NUMERO=" & XN(txtNroNotaPedido)
    'sql = sql & " AND RC.RCL_SUCURSAL=" & XN(txtRemSuc)
    sql = sql & " AND RC.NPE_FECHA=" & XDQ(FechaNotaPedido)
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        
        Imprimir 1, 7, True, "Cliente: " & Trim(Rec1!CLI_RAZSOC)
        Imprimir 8.6, 6.8, False, "Domicilio: " & Trim(IIf(IsNull(Rec1!CLI_DOMICI), "", Rec1!CLI_DOMICI))
        'REMITO
        'Imprimir 13.8, 7.2, True, Format(txtNroRemito.Text, "00000000") & " del " & Format(FechaRemito.Text, "dd/mm/yyyy")
        Imprimir 8.6, 7.2, False, "Localidad: " & Trim(Rec1!LOC_DESCRI) & " - " & Trim(Rec1!PRO_DESCRI)
        'Imprimir 1, 6.3, False, Trim(Rec1!IVA_DESCRI)
'        If Rec1!IVA_CODIGO = 1 Then
'            Imprimir 3.4, 7.9, False, "X"
'        Else
'            If Rec1!IVA_CODIGO = 4 Then 'Exento
'                Imprimir 10.1, 7.9, False, "X"
'            Else
'                Imprimir 7.1, 7.9, False, "X"
'            End If
'        End If
        'Imprimir 1, 6.3, False, Trim(Rec1!IVA_DESCRI)
        Imprimir 8.6, 7.9, False, "CUIT: " & IIf(IsNull(Rec1!CLI_CUIT), "", Format(Rec1!CLI_CUIT, "##-########-#"))
        Imprimir 8.6, 8.5, False, IIf(IsNull(Rec1!CLI_INGBRU), "", Format(Rec1!CLI_INGBRU, "###-#####-##"))
        Imprimir 1, 9, True, "CODIGO      DESCRIPCION                                                 CANT    PRECIO    IMPORTE"
        
    End If
    Rec1.Close
     
    
'    If cboCondicion.ItemData(cboCondicion.ListIndex) = 1 Then
'        Imprimir 7.1, 8.5, False, "X"
'    Else
'        Imprimir 10.25, 8.5, False, "X"
'    End If
    
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
   txtNroNotaPedido.Text = BuscoUltimoPedido
   txtNroNotaPedido.SetFocus
   BuscoIva
   txtImporteIva.Text = ""
   txtSubtotal.Text = ""
   txtTotal.Text = ""
   txtObservaciones.Text = ""
End Sub
Private Function SumaTotal() As Double
    Dim VTotal As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 4) <> "" Then
            VTotal = VTotal + (CDbl(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3)))
        End If
    Next
    SumaTotal = Valido_Importe(CStr(VTotal))
End Function
Private Sub BuscoIva()
    sql = "SELECT IVA FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtPorcentajeIva.Text = IIf(IsNull(rec!IVA), "", Format(rec!IVA, "0.00"))
    End If
    rec.Close
End Sub
Private Function BuscoUltimoPedido() As String
    'ACA BUSCA EL NUMERO DE PEDIDO SIGUIENTE AL ULTIMO CARGADO
    sql = "SELECT MAX(NPE_NUMERO) + 1 AS ULTIMO"
    sql = sql & " FROM NOTA_PEDIDO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        'txtNroSucursal.Text = Sucursal
        BuscoUltimoPedido = Format(rec!Ultimo, "00000000")
    End If
    rec.Close
End Function

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

Private Sub cmdPrecio_Click()
Dim BAND As Integer
    BAND = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            BAND = 1
        End If
    Next
    If BAND = 0 Then
        MsgBox "Para actualizar los precios debe haber al menos un Item el detalle", vbExclamation, TIT_MSGBOX
        grdGrilla.SetFocus
        Exit Sub
    End If
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            sql = "SELECT PTO_CODIGO,PTO_PRECIO FROM PRODUCTO"
            sql = sql & " WHERE PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "'"
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                grdGrilla.TextMatrix(I, 3) = Valido_Importe(rec!PTO_PRECIO)
            End If
            rec.Close
        End If
    Next
    TotalPresupuesto
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

Private Sub cmdredondeo_Click()
    txtTotal.Enabled = True
    If txtTotal.Text <> "" Then
        txtTotal.Text = Round(txtTotal.Text, 0)
        txtTotal.Text = Valido_Importe(txtTotal.Text)
        txtImporteIva.Text = txtTotal.Text - txtSubtotal
        txtImporteIva.Text = Valido_Importe(txtImporteIva)
    End If

End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmNotaDePedido = Nothing
        Unload Me
    End If
End Sub

Private Sub FechaNotaPedido_LostFocus()
    If IsNull(FechaNotaPedido.Value) Then
        FechaNotaPedido.Value = Date
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
    grdGrilla.FormatString = "Código|Descripción|Cantidad|Precio|Importe|Rubro|Linea|Orden"
    grdGrilla.ColWidth(0) = 1000 'CODIGO
    grdGrilla.ColWidth(1) = 5900 'DESCRIPCION
    grdGrilla.ColWidth(2) = 1000 'CANTIDAD
    grdGrilla.ColWidth(3) = 1100 'PRECIO
    grdGrilla.ColWidth(4) = 1100 'RUBRO
    grdGrilla.ColWidth(5) = 2100 'RUBRO
    grdGrilla.ColWidth(6) = 2100 'LINEA
    grdGrilla.ColWidth(7) = 0    'ORDEN
    grdGrilla.Cols = 8
    grdGrilla.Rows = 1
    For I = 2 To 14
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & (I - 1)
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
    'CARGO EL COMBO DE LISTA DE PRECIOS DE MAQUINARIAS
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
    'ULTIMO PRESUPEUSTO O PEDIDO
    txtNroNotaPedido.Text = BuscoUltimoPedido
    FechaNotaPedido.Value = Date
    BuscoIva
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
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(txtCliente)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
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
        sql = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CUIT,C.IVA_CODIGO,C.CLI_INGBRU,"
        sql = sql & "L.LOC_DESCRI,P.PRO_DESCRI,L.LOC_CODPOS"
        sql = sql & " FROM CLIENTE C, LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE "
        sql = sql & "C.LOC_CODIGO = L.LOC_CODIGO AND "
        sql = sql & "C.PRO_CODIGO = P.PRO_CODIGO AND "
        sql = sql & "L.PRO_CODIGO = P.PRO_CODIGO AND "
        sql = sql & "C.CLI_CODIGO=" & XN(TxtCodigoCli)
        'sql = sql & " AND CLI_ESTADO=1"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtRazSocCli.Text = rec!CLI_RAZSOC
            txtDomici.Text = IIf(IsNull(rec!CLI_DOMICI), "", rec!CLI_DOMICI)
            txtlocalidad.Text = rec!LOC_DESCRI
            txtProvincia.Text = rec!PRO_DESCRI
            txtCondicionIVA.Text = BuscoCondicionIVA(rec!IVA_CODIGO)
            txtCUIT.Text = IIf(IsNull(rec!CLI_CUIT), "NO INFORMADO", Format(rec!CLI_CUIT, "##-########-#"))
            txtIngBrutos.Text = IIf(IsNull(rec!CLI_INGBRU), "NO INFORMADO", Format(rec!CLI_INGBRU, "###-#####-##"))
            txtcodpos.Text = IIf(IsNull(rec!LOC_CODPOS), "", rec!LOC_CODPOS)
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtRazSocCli.Text = ""
            TxtCodigoCli.SetFocus
        End If
        rec.Close
    End If
End Sub

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
    
    'Del Pedido
    'If grdGrilla.Col = 0 Then KeyAscii = CarTexto(KeyAscii)
    'If grdGrilla.Col = 2 Then KeyAscii = CarNumeroEntero(KeyAscii)
    'If grdGrilla.Col = 3 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    'If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    'CarTexto KeyAscii
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
                    If lblEstadoNota.Caption = "PENDIENTE" Then
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
                    Else
                        sql = "SELECT DRC.PTO_CODIGO,DRC.DRC_DETALLE, DRC.DRC_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI"
                        sql = sql & " FROM DETALLE_REMITO_CLIENTE DRC, PRODUCTO P,LINEAS L,RUBROS R,TIPO_PRESENTACION RE"
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
                    grdGrilla.Col = 5
                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
                    grdGrilla.Col = 6
                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)
                    If grdGrilla.Text = "MAQUINARIA" Then
                        txtPorcentajeIva = "10,50"
                    Else
                        txtPorcentajeIva = "21,00"
                    End If
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = grdGrilla.RowSel
                    grdGrilla.Col = 2
                Else
'                    grdGrilla.Col = 0
'                    grdGrilla.Text = Trim(rec!PTO_CODIGO)
'                    grdGrilla.Col = 1
'                    grdGrilla.Text = Trim(rec!PTO_DESCRI)
'                    If cboListaPrecio.ListIndex = 0 Then
'                        grdGrilla.Text = Valido_Importe(Trim(rec!PTO_PRECIO))
'                    Else
'                        grdGrilla.Text = Valido_Importe(Trim(rec!PTO_PRECIO))
'                    End If
'                    grdGrilla.Col = 4
'                    grdGrilla.Text = Trim(rec!RUB_DESCRI)
'                    grdGrilla.Col = 5
'                    grdGrilla.Text = Trim(rec!LNA_DESCRI)
'                    grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
'                    grdGrilla.Col = 2
                     grdGrilla.Col = 0
                        grdGrilla.Text = Trim(rec!PTO_CODIGO)
                        If lblEstadoNota.Caption = "PENDIENTE" Then
                            grdGrilla.Col = 1
                            grdGrilla.Text = Trim(rec!PTO_DESCRI)
                            grdGrilla.Col = 3
                            If cboListaPrecio.ListIndex = 0 Then
                                grdGrilla.Text = Valido_Importe(Trim(rec!PTO_PRECIO))
                            Else
                                grdGrilla.Text = Valido_Importe(Trim(rec!PTO_PRECIO))
                            End If
                        Else
                            grdGrilla.Col = 1
                            grdGrilla.Text = Trim(rec!DRC_DETALLE)
                            grdGrilla.Col = 3
                            grdGrilla.Text = Valido_Importe(Trim(rec!DRC_PRECIO))
                        End If
                        
                        grdGrilla.Col = 5
                        grdGrilla.Text = Trim(rec!RUB_DESCRI)
                        grdGrilla.Col = 6
                        grdGrilla.Text = Trim(rec!LNA_DESCRI)
                        If grdGrilla.Text = "MAQUINARIA" Then
                            txtPorcentajeIva = "10,50"
                        Else
                            txtPorcentajeIva = "21,00"
                        End If
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = grdGrilla.RowSel
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
            If Trim(txtEdit) = "" Then grdGrilla.Text = "1"
            grdGrilla_LeaveCell
            TotalPresupuesto
            grdGrilla.SetFocus
        Case 3
            If Trim(txtEdit) <> "" Then
                txtEdit.Text = Valido_Importe(txtEdit)
                grdGrilla_LeaveCell
                TotalPresupuesto
                grdGrilla.SetFocus
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
Private Sub TotalPresupuesto()
    grdGrilla_LeaveCell
    'Me.txtTotal.Text = 0
    Dim TOTAL As Double
    Dim subtotal As Double
    Dim impIva As Double
    
    Me.grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Format(CDbl(Me.grdGrilla.TextMatrix(Me.grdGrilla.RowSel, 2)) * CDbl(Me.grdGrilla.TextMatrix(Me.grdGrilla.RowSel, 3)), "#,##0.00")
    
    txtSubtotal.Text = Valido_Importe(SumaTotal)
    txtImporteIva.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
    txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
    subtotal = txtSubtotal
    impIva = txtImporteIva
    TOTAL = subtotal + impIva
    txtTotal.Text = Format(TOTAL, "0.00")

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

Private Sub txtNroNotaPedido_Change()
    If txtNroNotaPedido.Text = "" Then
        'fechaNotaPedido.value = null
    End If
End Sub

Private Sub txtNroNotaPedido_GotFocus()
     'fechaNotaPedido.value = null
End Sub

Private Sub txtNroNotaPedido_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaPedido_LostFocus()
    If ActiveControl.Name = "CmdSalir" Or ActiveControl.Name = "chkCliente" _
       Or ActiveControl.Name = "cmdNuevo" Or ActiveControl.Name = "cmdBuscarNotaPedido" Then Exit Sub
    
    If txtNroNotaPedido.Text <> "" Then
        sql = "SELECT NP.*, E.EST_DESCRI"
        sql = sql & " FROM NOTA_PEDIDO NP, ESTADO_DOCUMENTO E"
        sql = sql & " WHERE NP.NPE_NUMERO=" & XN(txtNroNotaPedido)
        If FechaNotaPedido.Value <> "" Then
            sql = sql & " AND NP.NPE_FECHA=" & XDQ(FechaNotaPedido)
        End If
        sql = sql & " AND NP.EST_CODIGO=E.EST_CODIGO"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec2.EOF = False Then
            If Rec2.RecordCount > 1 Then
                MsgBox "Hay mas de un Presupuesto con el Número: " & txtNroNotaPedido.Text, vbInformation, TIT_MSGBOX
                tabDatos.Tab = 1
                Rec2.Close
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Buscando..."
            
            'CARGO CABECERA DE LA NOTA DE PEDIDO
            FechaNotaPedido.Value = Rec2!NPE_FECHA
            'txtNroVendedor.Text = Rec2!VEN_CODIGO
            'BUSCA FORMA DE PAGO
            If Not IsNull(Rec2!FPG_CODIGO) Then
                chkDetalle.Value = Checked
                Call BuscaCodigoProxItemData(Rec2!FPG_CODIGO, cboCondicion)
            Else
                chkDetalle.Value = Unchecked
            End If


            
            txtNroVendedor_LostFocus
            TxtCodigoCli.Text = Rec2!CLI_CODIGO
            TxtCodigoCli_LostFocus
            
            Call BuscoEstado(Rec2!EST_CODIGO, lblEstadoNota)
            If Rec2!EST_CODIGO <> 1 Then
                cmdGrabar.Enabled = False
                CmdBorrar.Enabled = False
                FramePedido.Enabled = False
                fraDatos.Enabled = False
                grdGrilla.SetFocus
            Else
                cmdGrabar.Enabled = True
                CmdBorrar.Enabled = True
                FramePedido.Enabled = True
                fraDatos.Enabled = True
            End If
            
            
            txtObservaciones = IIf(IsNull(Rec2!NPE_OBSERV), "", Rec2!NPE_OBSERV)
            
            'BUSCO LOS DATOS DEL DETALLE DE LA NOTA DE PEDIDO - PRESUPUESTO
            sql = "SELECT DNP.*,P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM DETALLE_NOTA_PEDIDO DNP, PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE DNP.NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND DNP.NPE_FECHA=" & XDQ(FechaNotaPedido)
            sql = sql & " AND DNP.PTO_CODIGO=P.PTO_CODIGO"
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
                    If IsNull(Rec1!DNP_PRECIO) Then
                        grdGrilla.TextMatrix(I, 3) = ""
                    Else
                        grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DNP_PRECIO)
                    End If
                    If IsNull(Rec1!DNP_IMPORTE) Then
                        grdGrilla.TextMatrix(I, 4) = ""
                    Else
                        grdGrilla.TextMatrix(I, 4) = Valido_Importe(Rec1!DNP_IMPORTE)
                    End If
                    grdGrilla.TextMatrix(I, 5) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 6) = Rec1!LNA_DESCRI
                    grdGrilla.TextMatrix(I, 7) = Rec1!DNP_NROITEM
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
        Dim TOTAL As Double
        Dim subtotal As Double
        Dim impIva As Double
        txtSubtotal.Text = Valido_Importe(SumaTotal)
        txtImporteIva.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
        txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
        subtotal = txtSubtotal
        impIva = txtImporteIva
        TOTAL = subtotal + impIva
        txtTotal.Text = Format(TOTAL, "0.00")
    Else
        MsgBox "Debe ingresar el Número del Presupuesto", vbExclamation, TIT_MSGBOX
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

