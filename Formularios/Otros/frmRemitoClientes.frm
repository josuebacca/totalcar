VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Object = "{F09A78C8-7814-11D2-8355-4854E82A9183}#1.0#0"; "CUIT32.OCX"
Begin VB.Form frmRemitoClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remito de Clientes"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmRemitoClientes.frx":0000
      Height          =   420
      Left            =   8490
      Picture         =   "frmRemitoClientes.frx":030A
      TabIndex        =   60
      Top             =   7050
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmRemitoClientes.frx":0614
      Height          =   420
      Left            =   7605
      Picture         =   "frmRemitoClientes.frx":091E
      TabIndex        =   59
      Top             =   7050
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmRemitoClientes.frx":0C28
      Height          =   420
      Left            =   10260
      Picture         =   "frmRemitoClientes.frx":0F32
      TabIndex        =   58
      Top             =   7050
      Width           =   870
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "frmRemitoClientes.frx":123C
      Height          =   420
      Left            =   9375
      Picture         =   "frmRemitoClientes.frx":1546
      TabIndex        =   57
      Top             =   7050
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6930
      Left            =   60
      TabIndex        =   19
      Top             =   75
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   12224
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   529
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
      TabPicture(0)   =   "frmRemitoClientes.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDatos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmRemitoClientes.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "GrdModulos"
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
         TabIndex        =   45
         Top             =   540
         Width           =   10395
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoClientes.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Buscar"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarSuc 
            Height          =   315
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoClientes.frx":1B92
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Buscar"
            Top             =   615
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3360
            TabIndex        =   12
            Top             =   960
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
            Left            =   4410
            TabIndex        =   52
            Top             =   975
            Width           =   5070
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   300
            TabIndex        =   8
            Top             =   945
            Width           =   1035
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1395
            Left            =   9870
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoClientes.frx":1E9C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Buscar Nota de Pedido"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   5865
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   1320
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
            Left            =   4845
            MaxLength       =   50
            TabIndex        =   47
            Tag             =   "Descripción"
            Top             =   255
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3360
            MaxLength       =   40
            TabIndex        =   10
            Top             =   255
            Width           =   975
         End
         Begin VB.TextBox txtSucursal 
            Height          =   300
            Left            =   3360
            MaxLength       =   40
            TabIndex        =   11
            Top             =   615
            Width           =   975
         End
         Begin VB.TextBox txtDesSuc 
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
            TabIndex        =   46
            Tag             =   "Descripción"
            Top             =   615
            Width           =   4620
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   300
            TabIndex        =   9
            Top             =   1215
            Width           =   810
         End
         Begin VB.CheckBox chkSucursal 
            Caption         =   "Sucursal"
            Height          =   195
            Left            =   300
            TabIndex        =   7
            Top             =   690
            Width           =   960
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   300
            TabIndex        =   6
            Top             =   435
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   2535
            TabIndex        =   53
            Top             =   1005
            Width           =   735
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4815
            TabIndex        =   51
            Top             =   1380
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2265
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   300
            Width           =   525
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2610
            TabIndex        =   48
            Top             =   675
            Width           =   660
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
         Height          =   2175
         Left            =   3480
         TabIndex        =   24
         Top             =   420
         Width           =   7545
         Begin VB.TextBox txtDescripcionSuc 
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
            Left            =   2865
            MaxLength       =   50
            TabIndex        =   18
            Tag             =   "Descripción"
            Top             =   780
            Width           =   4590
         End
         Begin VB.TextBox txtCodigoSuc 
            Height          =   285
            Left            =   960
            MaxLength       =   40
            TabIndex        =   4
            Top             =   780
            Width           =   975
         End
         Begin VB.CommandButton cmdBuscarSucursal 
            Height          =   315
            Left            =   1980
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoClientes.frx":21A6
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Buscar"
            Top             =   780
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaSucursal 
            Height          =   315
            Left            =   2415
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoClientes.frx":24B0
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Agregar Sucursal"
            Top             =   780
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Height          =   315
            Left            =   2415
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoClientes.frx":283A
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Agregar Cliente"
            Top             =   405
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   1980
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoClientes.frx":2BC4
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Buscar"
            Top             =   405
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtCondicionIVA 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Left            =   2520
            TabIndex        =   32
            Top             =   1485
            Width           =   2790
         End
         Begin VB.TextBox TxtCodigoCli 
            Height          =   285
            Left            =   960
            MaxLength       =   40
            TabIndex        =   3
            Top             =   420
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
            Height          =   315
            Left            =   2865
            MaxLength       =   50
            TabIndex        =   17
            Tag             =   "Descripción"
            Top             =   405
            Width           =   4590
         End
         Begin VB.TextBox txtDomici 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            MaxLength       =   50
            TabIndex        =   27
            Top             =   1140
            Width           =   4335
         End
         Begin VB.TextBox txtIngBrutos 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            MaxLength       =   10
            TabIndex        =   25
            Top             =   1815
            Width           =   1005
         End
         Begin Control_CUIT.CUIT txtCUIT 
            Height          =   315
            Left            =   960
            TabIndex        =   26
            Top             =   1485
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            ConSeparador    =   0   'False
            Text            =   ""
            Enabled         =   0   'False
            MensajeErr      =   ""
            nacPF           =   0   'False
            nacPJ           =   0   'False
            extPF           =   0   'False
            extPJ           =   0   'False
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   43
            Top             =   840
            Width           =   660
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
            TabIndex        =   31
            Top             =   465
            Width           =   525
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   210
            TabIndex        =   30
            Top             =   1170
            Width           =   675
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   285
            TabIndex        =   29
            Top             =   1530
            Width           =   600
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos:"
            Height          =   195
            Left            =   75
            TabIndex        =   28
            Top             =   1845
            Width           =   810
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   2175
         Left            =   105
         TabIndex        =   21
         Top             =   420
         Width           =   3360
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
            TabIndex        =   44
            Top             =   1350
            Width           =   3165
         End
         Begin VB.TextBox txtNroVendedor 
            Height          =   300
            Left            =   1020
            TabIndex        =   2
            Top             =   1005
            Width           =   780
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
         Begin VB.TextBox txtNroNotaPedido 
            Height          =   300
            Left            =   1020
            TabIndex        =   0
            Top             =   315
            Width           =   1260
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   420
            TabIndex        =   62
            Top             =   1770
            Width           =   540
         End
         Begin VB.Label lblEstadoNota 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Est. NotaPed"
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
            Height          =   315
            Left            =   1020
            TabIndex        =   61
            Top             =   1740
            Width           =   2265
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   225
            TabIndex        =   35
            Top             =   1050
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   465
            TabIndex        =   23
            Top             =   690
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Pedido:"
            Height          =   195
            Left            =   75
            TabIndex        =   22
            Top             =   345
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4245
         Left            =   -74655
         TabIndex        =   16
         Top             =   2325
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7488
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame3 
         Height          =   4335
         Left            =   105
         TabIndex        =   36
         Top             =   2505
         Width           =   10920
         Begin VB.CommandButton cmdBuscarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoClientes.frx":2ECE
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
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoClientes.frx":31D8
            Style           =   1  'Graphical
            TabIndex        =   39
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
            Picture         =   "frmRemitoClientes.frx":34E2
            Style           =   1  'Graphical
            TabIndex        =   38
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
            TabIndex        =   37
            Top             =   510
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   4095
            Left            =   90
            TabIndex        =   5
            Top             =   165
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   3
            Cols            =   4
            FixedCols       =   0
            BackColorSel    =   16777215
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   2
            HighLight       =   0
            SelectionMode   =   1
            FormatString    =   ""
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   20
         Top             =   570
         Width           =   1065
      End
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
      Left            =   255
      TabIndex        =   56
      Top             =   7110
      Width           =   750
   End
End
Attribute VB_Name = "frmRemitoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer

Private Sub chkCliente_Click()
    If chkCliente.Value = Checked Then
        txtCliente.Enabled = True
    Else
        txtCliente.Enabled = False
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

Private Sub chkSucursal_Click()
    If chkSucursal.Value = Checked Then
        txtSucursal.Enabled = True
    Else
        txtSucursal.Enabled = False
    End If
End Sub

Private Sub chkVendedor_Click()
    If chkVendedor.Value = Checked Then
        txtVendedor.Enabled = True
    Else
        txtVendedor.Enabled = False
    End If
End Sub

Private Sub cmdAgregarProducto_Click()
    ABMProducto.Show vbModal
End Sub

Private Sub cmdBorrar_Click()
    If txtNroNotaPedido.Text <> "" Then
        If MsgBox("Seguro que desea eliminar la Nota de Pedido Nro.: " & Trim(txtNroNotaPedido), vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
           On Error GoTo Seclavose
           
           sql = "SELECT P.EST_CODIGO, E.EST_DESCRI "
           sql = sql & " FROM NOTA_PEDIDO P, ESTADO_DOCUMENTO E"
           sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido)
           sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
           sql = sql & " AND P.EST_CODIGO=E.EST_CODIGO"
           rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
           
           If rec.EOF = False Then
                If rec!EST_CODIGO <> 1 Then
                    MsgBox "La Nota de Pedido no puede ser eliminada," & Chr(13) & _
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
            cmdNuevo_Click
        End If
    End If
    Exit Sub
    
Seclavose:
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT NP.*, C.CLI_RAZSOC, S.SUC_DESCRI"
    sql = sql & " FROM NOTA_PEDIDO NP, CLIENTE C, SUCURSAL S"
    sql = sql & " WHERE"
    sql = sql & " NP.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND NP.SUC_CODIGO=S.SUC_CODIGO"
    sql = sql & " AND C.CLI_CODIGO=S.CLI_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente)
    If txtSucursal.Text <> "" Then sql = sql & "AND NP.SUC_CODIGO=" & XN(txtSucursal)
    If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
    If FechaDesde <> "" Then sql = sql & " AND NP.NPE_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql = sql & " AND NP.NPE_FECHA<=" & XDQ(FechaHasta)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!NPE_NUMERO & Chr(9) & rec!NPE_FECHA _
                            & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!SUC_DESCRI
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
    frmBuscar.Show vbModal
    frmBuscar.grdBuscar.Col = 0
    txtCliente.Text = frmBuscar.grdBuscar.Text
    frmBuscar.grdBuscar.Col = 1
    txtDesCli.Text = frmBuscar.grdBuscar.Text
    txtSucursal.SetFocus
End Sub

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.Show vbModal
    frmBuscar.grdBuscar.Col = 0
    TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
    frmBuscar.grdBuscar.Col = 1
    txtRazSocCli.Text = frmBuscar.grdBuscar.Text
    txtCodigoSuc.SetFocus
End Sub

Private Sub cmdBuscarProducto_Click()
    grdGrilla.SetFocus
    frmBuscar.TipoBusqueda = 2
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    grdGrilla.Col = 0
    EDITAR grdGrilla, txtEdit, 13
    If Trim(frmBuscar.grdBuscar.Text) <> "" Then txtEdit.Text = frmBuscar.grdBuscar.Text
    TxtEdit_KeyDown vbKeyReturn, 0
End Sub

Private Sub cmdBuscarSuc_Click()
    frmBuscar.TipoBusqueda = 3
    frmBuscar.Show vbModal
    frmBuscar.grdBuscar.Col = 3
    txtCliente.Text = frmBuscar.grdBuscar.Text
    txtCliente_LostFocus
    frmBuscar.grdBuscar.Col = 0
    txtSucursal.Text = frmBuscar.grdBuscar.Text
    txtSucursal_LostFocus
End Sub

Private Sub cmdBuscarSucursal_Click()
    frmBuscar.TipoBusqueda = 3
    frmBuscar.Show vbModal
    frmBuscar.grdBuscar.Col = 3
    TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
    TxtCodigoCli_LostFocus
    frmBuscar.grdBuscar.Col = 0
    txtCodigoSuc.Text = frmBuscar.grdBuscar.Text
    txtCodigoSuc_LostFocus
End Sub

Private Sub CmdGrabar_Click()
    If ValidarNotaPedido = False Then Exit Sub
    
    On Error GoTo HayErrorNota
    
    DBConn.BeginTrans
    sql = "SELECT * FROM NOTA_PEDIDO"
    sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido)
    sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = False Then
        If MsgBox("Seguro que modificar la Nota de Pedido Nro.: " & Trim(txtNroNotaPedido), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            sql = "UPDATE NOTA_PEDIDO"
            sql = sql & " SET CLI_CODIGO=" & XN(TxtCodigoCli)
            sql = sql & " ,SUC_CODIGO=" & XN(txtCodigoSuc)
            sql = sql & " ,VEN_CODIGO=" & XN(txtNroVendedor)
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
                    sql = sql & " (NPE_NUMERO,NPE_FECHA,DNP_NROITEM,PTO_CODIGO,DNP_CANTIDAD)"
                    sql = sql & " VALUES ("
                    sql = sql & XN(txtNroNotaPedido) & ","
                    sql = sql & XDQ(FechaNotaPedido) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 0)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ")"
                    DBConn.Execute sql
                End If
            Next
            DBConn.CommitTrans
        End If
    Else
        sql = "INSERT INTO NOTA_PEDIDO"
        sql = sql & " (NPE_NUMERO,NPE_FECHA,CLI_CODIGO,"
        sql = sql & "SUC_CODIGO,VEN_CODIGO,EST_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(txtNroNotaPedido) & ","
        sql = sql & XDQ(FechaNotaPedido) & ","
        sql = sql & XN(TxtCodigoCli) & ","
        sql = sql & XN(txtCodigoSuc) & ","
        sql = sql & XN(txtNroVendedor) & ","
        sql = sql & "1)" 'ESTADO PENDIENTE
        DBConn.Execute sql
           
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                sql = "INSERT INTO DETALLE_NOTA_PEDIDO"
                sql = sql & " (NPE_NUMERO,NPE_FECHA,DNP_NROITEM,PTO_CODIGO,DNP_CANTIDAD)"
                sql = sql & " VALUES ("
                sql = sql & XN(txtNroNotaPedido) & ","
                sql = sql & XDQ(FechaNotaPedido) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 0)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ")"
                DBConn.Execute sql
            End If
        Next
        DBConn.CommitTrans
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    cmdNuevo_Click
    Exit Sub
    
HayErrorNota:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description
End Sub
Private Function ValidarNotaPedido() As Boolean
    
    If txtNroNotaPedido.Text = "" Then
        MsgBox "El número de Nota de Pedido es requerido", vbExclamation, TIT_MSGBOX
        txtNroNotaPedido.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    If FechaNotaPedido.Text = "" Then
        MsgBox "La Fecha de la Nota de pedido es requerida", vbExclamation, TIT_MSGBOX
        FechaNotaPedido.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    If txtNroVendedor.Text = "" Then
        MsgBox "El Vendedor es requerido", vbExclamation, TIT_MSGBOX
        txtNroVendedor.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    If TxtCodigoCli.Text = "" Then
        MsgBox "El Cliente es requerido", vbExclamation, TIT_MSGBOX
        TxtCodigoCli.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    If txtCodigoSuc.Text = "" Then
        MsgBox "El Cliente es requerido", vbExclamation, TIT_MSGBOX
        TxtCodigoCli.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    ValidarNotaPedido = True
End Function

Private Sub cmdNuevaSucursal_Click()
    ABMSucursal.Show vbModal
    txtCodigoSuc.SetFocus
End Sub

Private Sub cmdNuevo_Click()
   For I = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(I, 0) = ""
        grdGrilla.TextMatrix(I, 1) = ""
        grdGrilla.TextMatrix(I, 2) = ""
        grdGrilla.TextMatrix(I, 3) = I
   Next
   LimpiarSucursal
   TxtCodigoCli.Text = ""
   txtRazSocCli.Text = ""
   txtNombreVendedor.Text = ""
   txtNroVendedor.Text = ""
   FechaNotaPedido.Text = ""
   txtNroNotaPedido.Text = ""
   lblEstadoNota.Caption = ""
   lblEstado.Caption = ""
   cmdGrabar.Enabled = True
   cmdBorrar.Enabled = True
   txtNroNotaPedido.SetFocus
End Sub

Private Sub LimpiarSucursal()
    txtCodigoSuc.Text = ""
    txtDescripcionSuc.Text = ""
    txtCUIT.Text = ""
    txtIngBrutos.Text = ""
    txtCondicionIVA.Text = ""
    txtdomici.Text = ""
End Sub
Private Sub cmdNuevoCliente_Click()
    ABMCliente.Show vbModal
    TxtCodigoCli.SetFocus
End Sub

Private Sub cmdQuitarProducto_Click()
    If MsgBox("Seguro que desea quitar el Producto: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 1), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = grdGrilla.RowSel
    End If
End Sub

Private Sub cmdSalir_Click()
    Set frmRemitoClientes = Nothing
    Unload Me
End Sub

Private Sub FechaNotaPedido_LostFocus()
    If FechaNotaPedido.Text = "" Then FechaNotaPedido.Text = Date
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
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    grdGrilla.FormatString = "Código|Descripción|Cantidad|orden"
    grdGrilla.ColWidth(0) = 1000
    grdGrilla.ColWidth(1) = 5900
    grdGrilla.ColWidth(2) = 1000
    grdGrilla.ColWidth(3) = 0
    grdGrilla.Cols = 4
    grdGrilla.Rows = 1
    For I = 2 To 14
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & (I - 1)
    Next
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = ">Número|^Fecha|Cliente|Sucursal"
    GrdModulos.ColWidth(0) = 1300
    GrdModulos.ColWidth(1) = 1000
    GrdModulos.ColWidth(2) = 4000
    GrdModulos.ColWidth(3) = 4000
    GrdModulos.Rows = 1
    '------------------------------------
    lblEstado.Caption = ""
    lblEstadoNota.Caption = ""
    tabDatos.Tab = 0
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1
            LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
            grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = grdGrilla.RowSel
            grdGrilla.Col = 0
        'Case Else
        '    grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, grdGrilla.Col)) = ""
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
        Case 2
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "" Then
                grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "1"
            End If
        End Select
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
       (grdGrilla.Col = 2) Then
        If KeyAscii = vbKeyReturn Then
            If grdGrilla.Col = 2 Then
                If grdGrilla.row < grdGrilla.Rows - 1 Then
                    grdGrilla.row = grdGrilla.row + 1
                    grdGrilla.Col = 0
                Else
                    SendKeys "{TAB}"
                End If
            Else
                grdGrilla.Col = grdGrilla.Col + 1 '3
            End If
        Else
            EDITAR grdGrilla, txtEdit, KeyAscii
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
        cmdNuevo_Click
        txtNroNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        FechaNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
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
    txtSucursal.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    txtVendedor.Enabled = False
    cmdGrabar.Enabled = False
    cmdBorrar.Enabled = False
    LimpiarBusqueda
    If Me.Visible = True Then chkCliente.SetFocus
  Else
    If Me.Visible = True Then txtNroNotaPedido.SetFocus
    cmdGrabar.Enabled = True
    cmdBorrar.Enabled = True
  End If
End Sub

Private Sub LimpiarBusqueda()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    txtSucursal.Text = ""
    txtDesSuc.Text = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    txtVendedor.Text = ""
    txtDesVen.Text = ""
    GrdModulos.Rows = 1
    chkCliente.Value = Unchecked
    chkSucursal.Value = Unchecked
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
    If chkSucursal.Value = Unchecked And chkFecha.Value = Unchecked _
        And chkVendedor.Value = Unchecked Then CmdBuscAprox.SetFocus
End Sub

Private Sub TxtCodigoCli_Change()
    If TxtCodigoCli.Text = "" Then
        txtRazSocCli.Text = ""
    End If
End Sub

Private Sub txtCodigoCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodigoSuc_Change()
    If txtCodigoSuc.Text = "" Then
        LimpiarSucursal
    End If
End Sub

Private Sub txtCodigoSuc_GotFocus()
    SelecTexto txtCodigoSuc
End Sub

Private Sub txtCodigoSuc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodigoSuc_LostFocus()
    If txtCodigoSuc.Text <> "" Then
        sql = "SELECT * FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(txtCodigoSuc)
        If TxtCodigoCli.Text <> "" Then
         sql = sql & " AND CLI_CODIGO=" & XN(TxtCodigoCli)
        End If
        lblEstado.Caption = "Buscando..."
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
                If Rec1.RecordCount > 1 Then
                    frmBuscar.TipoBusqueda = 3
                    frmBuscar.TxtDescriB = txtCodigoSuc
                    frmBuscar.Show vbModal
                    frmBuscar.grdBuscar.Col = 3
                    TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
                    TxtCodigoCli_LostFocus
                    frmBuscar.grdBuscar.Col = 0
                    txtCodigoSuc.Text = frmBuscar.grdBuscar.Text
                    txtCodigoSuc_LostFocus
                    Rec1.Close
                    lblEstado.Caption = ""
                    Exit Sub
                End If
            TxtCodigoCli.Text = Rec1!CLI_CODIGO
            TxtCodigoCli_LostFocus
            txtDescripcionSuc.Text = Rec1!SUC_DESCRI
            txtdomici.Text = Rec1!SUC_DOMICI
            txtCondicionIVA.Text = BuscoCondicionIVA(Rec1!IVA_CODIGO)
            txtCUIT.Text = Rec1!SUC_CUIT
            txtIngBrutos.Text = IIf(IsNull(Rec1!SUC_INGBRU), "", Rec1!SUC_INGBRU)
            txtNroVendedor.Text = Trim(Rec1!VEN_CODIGO)
            txtNroVendedor_LostFocus
            grdGrilla.SetFocus
            lblEstado.Caption = ""
        Else
            lblEstado.Caption = ""
            MsgBox "La Sucursal no existe", vbExclamation, TIT_MSGBOX
            txtDescripcionSuc.Text = ""
            txtCodigoSuc.SetFocus
        End If
        Rec1.Close
    End If
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
Private Sub TxtCodigoCli_GotFocus()
    SelecTexto TxtCodigoCli
End Sub

Private Sub TxtCodigoCli_LostFocus()
    If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Or ActiveControl.Name = "CmdSalir" Then Exit Sub
    If TxtCodigoCli.Text <> "" Then
        sql = "SELECT * FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(TxtCodigoCli)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtRazSocCli.Text = rec!CLI_RAZSOC
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
    If grdGrilla.Col = 0 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 2 Then KeyAscii = CarNumeroEntero(KeyAscii)
    CarTexto KeyAscii
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 Then
        frmBuscar.TipoBusqueda = 2
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
            sql = "SELECT PTO_CODIGO,PTO_DESCRI,PTO_PRECIO"
            sql = sql & " FROM PRODUCTO"
            sql = sql & " WHERE"
            If grdGrilla.Col = 0 Then
                sql = sql & " PTO_CODIGO=" & XN(txtEdit)
            Else
                sql = sql & " PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
            End If
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                If rec.RecordCount > 1 Then
                    grdGrilla.SetFocus
                    frmBuscar.TipoBusqueda = 2
                    frmBuscar.TxtDescriB.Text = txtEdit.Text
                    frmBuscar.Show vbModal
                    grdGrilla.Col = 0
                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
                    grdGrilla.Col = 1
                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = grdGrilla.RowSel
                    grdGrilla.Col = 2
                Else
                    grdGrilla.Col = 0
                    grdGrilla.Text = Trim(rec!PTO_CODIGO)
                    grdGrilla.Col = 1
                    grdGrilla.Text = Trim(rec!PTO_DESCRI)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = grdGrilla.RowSel
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
        If FechaNotaPedido.Text <> "" Then
            sql = sql & " AND NP.NPE_FECHA=" & XDQ(FechaNotaPedido)
        End If
        sql = sql & " AND NP.EST_CODIGO=E.EST_CODIGO"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec2.EOF = False Then
            If Rec2.RecordCount > 1 Then
                MsgBox "Hay mas de una Nota de Pedido con el Número: " & txtNroNotaPedido.Text, vbInformation, TIT_MSGBOX
                tabDatos.Tab = 1
                Rec2.Close
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Buscando..."
            
            'CARGO CABECERA DE LA NOTA DE PEDIDO
            FechaNotaPedido.Text = Rec2!NPE_FECHA
            txtNroVendedor.Text = Rec2!VEN_CODIGO
            txtNroVendedor_LostFocus
            TxtCodigoCli.Text = Rec2!CLI_CODIGO
            TxtCodigoCli_LostFocus
            txtCodigoSuc.Text = Rec2!SUC_CODIGO
            txtCodigoSuc_LostFocus
            lblEstadoNota.Caption = Rec2!EST_DESCRI
            If Rec2!EST_CODIGO <> 1 Then
                cmdGrabar.Enabled = False
                cmdBorrar.Enabled = False
            Else
                cmdGrabar.Enabled = True
                cmdBorrar.Enabled = True
            End If
            
            'BUSCO LOS DATOS DEL DETALLE DE LA NOTA DE PEDIDO
            sql = "SELECT DNP.*,P.PTO_DESCRI"
            sql = sql & " FROM DETALLE_NOTA_PEDIDO DNP, PRODUCTO P"
            sql = sql & " WHERE DNP.NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND DNP.NPE_FECHA=" & XDQ(FechaNotaPedido)
            sql = sql & " AND DNP.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " ORDER BY DNP.DNP_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(I, 1) = Rec1!PTO_DESCRI
                    grdGrilla.TextMatrix(I, 2) = Rec1!DNP_CANTIDAD
                    grdGrilla.TextMatrix(I, 3) = Rec1!DNP_NROITEM
                    I = I + 1
                    Rec1.MoveNext
                Loop
            End If
            Rec1.Close
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
        Else
            lblEstadoNota.Caption = "PENDIENTE"
        End If
        Rec2.Close
    Else
        MsgBox "Debe ingresar el Número de Nota de Pedido", vbExclamation, TIT_MSGBOX
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
    If txtNroVendedor.Text <> "" Then
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
    End If
End Sub

Private Sub txtSucursal_Change()
    If txtSucursal.Text = "" Then
        txtDesSuc.Text = ""
    End If
End Sub

Private Sub txtSucursal_GotFocus()
    SelecTexto txtSucursal
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtSucursal_LostFocus()
    
    If txtSucursal.Text <> "" Then
        sql = "SELECT CLI_CODIGO, SUC_DESCRI FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(txtSucursal)
        If TxtCodigoCli.Text <> "" Then
         sql = sql & " AND CLI_CODIGO=" & XN(txtCliente)
        End If
        lblEstado.Caption = "Buscando..."
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtCliente.Text = Rec1!CLI_CODIGO
            txtCliente_LostFocus
            txtDesSuc.Text = Rec1!SUC_DESCRI
            lblEstado.Caption = ""
        Else
            lblEstado.Caption = ""
            MsgBox "La Sucursal no existe", vbExclamation, TIT_MSGBOX
            txtDesSuc.Text = ""
            txtSucursal.SetFocus
        End If
        Rec1.Close
    End If
    If chkFecha.Value = Unchecked And chkVendedor.Value = Unchecked Then CmdBuscAprox.SetFocus
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
    If chkFecha.Value = Unchecked Then CmdBuscAprox.SetFocus
End Sub
