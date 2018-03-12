VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmImputarNCaFactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imputar Nota de Crédito Clientes a Facturas..."
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10725
      TabIndex        =   10
      Top             =   5880
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   8955
      TabIndex        =   8
      Top             =   5880
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   9840
      TabIndex        =   9
      Top             =   5880
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   5820
      Left            =   60
      TabIndex        =   24
      Top             =   15
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   10266
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
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
      TabPicture(0)   =   "frmImputarNCaFactura.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameNotaCredito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameCliente"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmImputarNCaFactura.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   105
         TabIndex        =   48
         Top             =   435
         Width           =   6270
         Begin VB.CommandButton cmdBuscarCliente1 
            Height          =   315
            Left            =   1755
            MaskColor       =   &H000000FF&
            Picture         =   "frmImputarNCaFactura.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Buscar Cliente"
            Top             =   450
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtDomici 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   930
            MaxLength       =   50
            TabIndex        =   54
            Top             =   840
            Width           =   5130
         End
         Begin VB.TextBox txtCliLocalidad 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   930
            TabIndex        =   52
            Top             =   1177
            Width           =   5130
         End
         Begin VB.TextBox txtProvincia 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Left            =   930
            TabIndex        =   50
            Top             =   1515
            Width           =   5130
         End
         Begin VB.TextBox txtCliRazSoc 
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
            Left            =   2190
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "Descripción"
            Top             =   450
            Width           =   3870
         End
         Begin VB.TextBox txtCodCliente 
            Height          =   300
            Left            =   930
            MaxLength       =   40
            TabIndex        =   0
            Top             =   450
            Width           =   780
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   255
            TabIndex        =   55
            Top             =   870
            Width           =   675
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   195
            TabIndex        =   53
            Top             =   1215
            Width           =   735
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   225
            TabIndex        =   51
            Top             =   1545
            Width           =   705
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   390
            TabIndex        =   49
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Frame frameBuscar 
         Caption         =   "Buscar Nota de Crédito por..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   -74610
         TabIndex        =   29
         Top             =   540
         Width           =   11025
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3360
            TabIndex        =   17
            Top             =   600
            Width           =   990
         End
         Begin VB.ComboBox cboBuscaRep 
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1620
            Width           =   2970
         End
         Begin VB.CheckBox chkRepresentada 
            Caption         =   "Representada"
            Height          =   195
            Left            =   300
            TabIndex        =   15
            Top             =   1515
            Width           =   1380
         End
         Begin VB.CommandButton cmdBuscarVen 
            Height          =   300
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmImputarNCaFactura.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Buscar Vendedor"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboNotaCredito1 
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1260
            Width           =   2970
         End
         Begin VB.CheckBox chkTipoFactura 
            Caption         =   "Tipo"
            Height          =   195
            Left            =   300
            TabIndex        =   14
            Top             =   1260
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmImputarNCaFactura.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Buscar Cliente"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
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
            TabIndex        =   34
            Top             =   615
            Width           =   4620
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   300
            TabIndex        =   12
            Top             =   735
            Width           =   1020
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1755
            Left            =   9690
            MaskColor       =   &H000000FF&
            Picture         =   "frmImputarNCaFactura.frx":0956
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
            TabIndex        =   30
            Tag             =   "Descripción"
            Top             =   255
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3360
            MaxLength       =   40
            TabIndex        =   16
            Top             =   255
            Width           =   975
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   300
            TabIndex        =   13
            Top             =   990
            Width           =   810
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   300
            TabIndex        =   11
            Top             =   480
            Width           =   855
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   3360
            TabIndex        =   18
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61276161
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   6360
            TabIndex        =   19
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61276161
            CurrentDate     =   41098
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Representada:"
            Height          =   195
            Left            =   2220
            TabIndex        =   64
            Top             =   1650
            Width           =   1050
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   2910
            TabIndex        =   47
            Top             =   1305
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   2535
            TabIndex        =   35
            Top             =   630
            Width           =   735
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   5295
            TabIndex        =   33
            Top             =   990
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2265
            TabIndex        =   32
            Top             =   975
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
            TabIndex        =   31
            Top             =   300
            Width           =   525
         End
      End
      Begin VB.Frame FrameNotaCredito 
         Caption         =   "Nota de Crédito..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   6390
         TabIndex        =   26
         Top             =   435
         Width           =   5100
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
            Left            =   1230
            MaxLength       =   4
            TabIndex        =   4
            Top             =   960
            Width           =   555
         End
         Begin VB.TextBox txtNroNotaCredito 
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
            Left            =   1815
            MaxLength       =   8
            TabIndex        =   5
            Top             =   960
            Width           =   1065
         End
         Begin VB.ComboBox cboRep 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   270
            Visible         =   0   'False
            Width           =   3090
         End
         Begin VB.CommandButton cmdBuscarNotaCRedito 
            Caption         =   "&Buscar"
            Height          =   315
            Left            =   4245
            TabIndex        =   6
            ToolTipText     =   "Buscar Factura"
            Top             =   1545
            UseMaskColor    =   -1  'True
            Width           =   750
         End
         Begin VB.ComboBox cboNotaCredito 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   615
            Width           =   2400
         End
         Begin MSComCtl2.DTPicker FechaNotaCredito 
            Height          =   315
            Left            =   1230
            TabIndex        =   66
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61276161
            CurrentDate     =   41098
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Representada:"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   300
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   810
            TabIndex        =   42
            Top             =   630
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   675
            TabIndex        =   41
            Top             =   1350
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   570
            TabIndex        =   40
            Top             =   990
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   630
            TabIndex        =   39
            Top             =   1665
            Width           =   540
         End
         Begin VB.Label lblEstadoNotaCredito 
            AutoSize        =   -1  'True
            Caption         =   "EST. NOTA CREDITO"
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
            Left            =   1230
            TabIndex        =   38
            Top             =   1680
            Width           =   1890
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3330
         Left            =   105
         TabIndex        =   27
         Top             =   2295
         Width           =   11400
         Begin VB.CommandButton cmdQuitar 
            Height          =   555
            Left            =   5445
            Picture         =   "frmImputarNCaFactura.frx":30F8
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Quitar Factura de la Imputación"
            Top             =   1455
            Width           =   540
         End
         Begin VB.CommandButton CmdAgregar 
            Height          =   555
            Left            =   5445
            Picture         =   "frmImputarNCaFactura.frx":353A
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Agregar Factura a la Imputación"
            Top             =   885
            Width           =   540
         End
         Begin VB.TextBox txtSaldoNC 
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
            Left            =   9750
            TabIndex        =   44
            Top             =   2895
            Width           =   1350
         End
         Begin VB.TextBox txtTotalNC 
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
            Left            =   9750
            TabIndex        =   43
            Top             =   2550
            Width           =   1350
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   6495
            TabIndex        =   28
            Top             =   990
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   1965
            Left            =   6000
            TabIndex        =   7
            Top             =   495
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   3466
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
         End
         Begin MSFlexGridLib.MSFlexGrid grillaFactura 
            Height          =   1965
            Left            =   60
            TabIndex        =   58
            Top             =   495
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   3466
            _Version        =   393216
            Rows            =   3
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   260
            BackColor       =   16777215
            BackColorSel    =   8388736
            BackColorBkg    =   -2147483633
            FocusRect       =   0
            SelectionMode   =   1
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Facturas asignadas a la Nota de Crédito"
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
            Left            =   6060
            TabIndex        =   62
            Top             =   210
            Width           =   4215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Facturas con Saldo"
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
            Left            =   120
            TabIndex        =   61
            Top             =   210
            Width           =   2025
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nota de Crédito"
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
            Left            =   7335
            TabIndex        =   57
            Top             =   2535
            Width           =   1650
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   9270
            TabIndex        =   46
            Top             =   2580
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Saldo:"
            Height          =   195
            Left            =   9225
            TabIndex        =   45
            Top             =   2940
            Width           =   450
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   2940
         Left            =   -74625
         TabIndex        =   23
         Top             =   2730
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   5186
         _Version        =   393216
         Cols            =   11
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
         TabIndex        =   25
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
      Left            =   225
      TabIndex        =   37
      Top             =   5940
      Width           =   750
   End
End
Attribute VB_Name = "frmImputarNCaFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim w As Integer
Dim TipoBusquedaDoc As Integer ' 1 Busca NC 0 Busca Imputaciones
Dim VBonificacion As Double
Dim VTotal As Double
Dim VEstadoNotaCredito As Integer
Dim vAccion As Integer  '0 consulta inicial 1 busqueda desde Buscar

Private Sub cboRep_LostFocus()
    If txtCodCliente <> "" Then
        Call BuscarFacturas(txtCodCliente) ', CStr(cboRep.ItemData(cboRep.ListIndex)))
    End If
End Sub

Private Sub chkRepresentada_Click()
    If chkRepresentada.Value = Checked Then
        cboBuscaRep.Enabled = True
        cboBuscaRep.ListIndex = 0
    Else
        cboBuscaRep.Enabled = False
        cboBuscaRep.ListIndex = -1
    End If
End Sub

Private Sub CmdAgregar_Click()
    If grillaFactura.Rows > 1 Then
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 6) = grillaFactura.TextMatrix(grillaFactura.RowSel, 5) And _
               grdGrilla.TextMatrix(I, 1) = grillaFactura.TextMatrix(grillaFactura.RowSel, 1) And _
               grdGrilla.TextMatrix(I, 2) = grillaFactura.TextMatrix(grillaFactura.RowSel, 2) Then
               
               MsgBox "La Factura ya fue elegida", vbExclamation, TIT_MSGBOX
               grillaFactura.SetFocus
               Exit Sub
            End If
        Next
        grdGrilla.AddItem grillaFactura.TextMatrix(grillaFactura.RowSel, 0) & Chr(9) & _
                          grillaFactura.TextMatrix(grillaFactura.RowSel, 1) & Chr(9) & _
                          grillaFactura.TextMatrix(grillaFactura.RowSel, 2) & Chr(9) & _
                          grillaFactura.TextMatrix(grillaFactura.RowSel, 3) & Chr(9) & _
                          grillaFactura.TextMatrix(grillaFactura.RowSel, 4) & Chr(9) & _
                          "" & Chr(9) & _
                          grillaFactura.TextMatrix(grillaFactura.RowSel, 5)
    End If
End Sub

Private Sub cmdBuscarCliente1_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCodCliente.Text = frmBuscar.grdBuscar.Text
        txtCodCliente.SetFocus
        txtCodCliente_LostFocus
    Else
        txtCodCliente.SetFocus
    End If
End Sub

Private Sub cmdBuscarNotaCRedito_Click()
    If txtCodCliente.Text <> "" Then
        TipoBusquedaDoc = 1
        tabDatos.Tab = 1
    Else
        MsgBox "Debe elegir un Cliente", vbExclamation, TIT_MSGBOX
        txtCodCliente.SetFocus
        Exit Sub
    End If
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

Private Sub QuitoFacturaImputadaAntes()
  Dim SaldoFac As Double
  SaldoFac = 0
  
    If MsgBox("¿Seguro que desea quitar la Factura Nro.: " _
               & grdGrilla.TextMatrix(grdGrilla.RowSel, 1) & "?" _
               & Chr(13) & "La misma ya fue imputada con anterioridad.", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        
        On Error GoTo QuePaso
        DBConn.BeginTrans
        
        SaldoFac = BuscoSaldoFactura(grdGrilla.TextMatrix(grdGrilla.RowSel, 6), Right(grdGrilla.TextMatrix(grdGrilla.RowSel, 1), 8) _
                                        , Left(grdGrilla.TextMatrix(grdGrilla.RowSel, 1), 4), grdGrilla.TextMatrix(grdGrilla.RowSel, 2), _
                                        CStr(cboRep.ItemData(cboRep.ListIndex)))
                                        
        If grdGrilla.Rows > 2 Then
            For I = 1 To grillaFactura.Rows - 1
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grillaFactura.TextMatrix(I, 5) And _
                   grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = grillaFactura.TextMatrix(I, 1) And _
                   grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = grillaFactura.TextMatrix(I, 2) Then
    
                    grillaFactura.TextMatrix(I, 4) = Valido_Importe(CStr(SaldoFac + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
                   Exit For
                End If
            Next
            txtSaldoNC.Text = Valido_Importe(CStr(CDbl(txtSaldoNC.Text) + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
            
            Call ActualizoSaldoFacturas(grdGrilla.TextMatrix(grdGrilla.RowSel, 6), Right(grdGrilla.TextMatrix(grdGrilla.RowSel, 1), 8) _
                                        , Left(grdGrilla.TextMatrix(grdGrilla.RowSel, 1), 4), grdGrilla.TextMatrix(grdGrilla.RowSel, 2) _
                                        , CStr(SaldoFac + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))), CStr(cboRep.ItemData(cboRep.ListIndex)))
                                        
            'FACTURAS POR NOTA DE CREDITO
            Call QuitoLaFacturaDeLaTabla(grdGrilla.TextMatrix(grdGrilla.RowSel, 6), Right(grdGrilla.TextMatrix(grdGrilla.RowSel, 1), 8) _
                                        , Left(grdGrilla.TextMatrix(grdGrilla.RowSel, 1), 4), grdGrilla.TextMatrix(grdGrilla.RowSel, 2), _
                                        CStr(cboRep.ItemData(cboRep.ListIndex)))
            
            grdGrilla.RemoveItem grdGrilla.RowSel
            
        Else
            
            For I = 1 To grillaFactura.Rows - 1
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grillaFactura.TextMatrix(I, 5) And _
                   grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = grillaFactura.TextMatrix(I, 1) And _
                   grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = grillaFactura.TextMatrix(I, 2) Then
    
                    grillaFactura.TextMatrix(I, 4) = Valido_Importe(CStr(SaldoFac + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
                   Exit For
                End If
            Next
            txtSaldoNC.Text = Valido_Importe(CStr(CDbl(txtSaldoNC.Text) + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
            
            Call ActualizoSaldoFacturas(grdGrilla.TextMatrix(grdGrilla.RowSel, 6), Right(grdGrilla.TextMatrix(grdGrilla.RowSel, 1), 8) _
                              , Left(grdGrilla.TextMatrix(grdGrilla.RowSel, 1), 4), grdGrilla.TextMatrix(grdGrilla.RowSel, 2), _
                              CStr(SaldoFac + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))), CStr(cboRep.ItemData(cboRep.ListIndex)))
                              
            'FACTURAS POR NOTA DE CREDITO
            Call QuitoLaFacturaDeLaTabla(grdGrilla.TextMatrix(grdGrilla.RowSel, 6), Right(grdGrilla.TextMatrix(grdGrilla.RowSel, 1), 8) _
                                        , Left(grdGrilla.TextMatrix(grdGrilla.RowSel, 1), 4), grdGrilla.TextMatrix(grdGrilla.RowSel, 2) _
                                        , CStr(cboRep.ItemData(cboRep.ListIndex)))
                                        
            grdGrilla.Rows = 1
            'grdGrilla.HighLight = flexHighlightNever
        End If
            
            'ACTUALIZO EL SALDO DE LA NOTA DE CREDITO
            sql = "UPDATE NOTA_CREDITO_CLIENTE"
            sql = sql & " SET NCC_SALDO=" & XN(txtSaldoNC.Text)
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
            sql = sql & " AND NCC_NUMERO=" & XN(txtNroNotaCredito)
            sql = sql & " AND NCC_SUCURSAL=" & XN(txtNroSucursal)
            sql = sql & " AND REP_CODIGO=" & XN(cboRep.ItemData(cboRep.ListIndex))
            sql = sql & " AND NCC_FECHA=" & XDQ(FechaNotaCredito)
            DBConn.Execute sql
            
        DBConn.CommitTrans
    End If
    Exit Sub
    
QuePaso:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    CmdNuevo_Click
End Sub

Private Sub QuitoLaFacturaDeLaTabla(TIPO As String, Numero As String, Sucursal As String, Fecha As String, Rep As String)
    'FACTURAS POR NOTA DE CREDITO
    sql = "DELETE FROM FACTURAS_NOTA_CREDITO_CLIENTE"
    sql = sql & " WHERE"
    sql = sql & " TCO_CODIGO=" & XN(cboNotaCredito.ItemData(cboNotaCredito.ListIndex))
    sql = sql & " AND NCC_NUMERO=" & XN(txtNroNotaCredito)
    sql = sql & " AND NCC_SUCURSAL=" & XN(txtNroSucursal)
    sql = sql & " AND NCC_FECHA=" & XDQ(FechaNotaCredito)
    sql = sql & " AND FCL_TCO_CODIGO=" & XN(TIPO)
    sql = sql & " AND FCL_NUMERO=" & XN(Numero)
    sql = sql & " AND FCL_SUCURSAL=" & XN(Sucursal)
    sql = sql & " AND REP_CODIGO=" & XN(Rep)
    sql = sql & " AND FCL_FECHA=" & XDQ(Fecha)
    
    DBConn.Execute sql
End Sub

Private Sub cmdQuitar_Click()
    If grdGrilla.Rows > 1 Then
        If grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = "X" Then
            QuitoFacturaImputadaAntes
            Exit Sub
        End If
        If MsgBox("¿Seguro que desea quitar la Factura Nro.: " _
                & grdGrilla.TextMatrix(grdGrilla.RowSel, 1) & "?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
                If grdGrilla.Rows > 2 Then
                    For I = 1 To grillaFactura.Rows - 1
                        If grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grillaFactura.TextMatrix(I, 5) And _
                           grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = grillaFactura.TextMatrix(I, 1) And _
                           grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = grillaFactura.TextMatrix(I, 2) Then
            
                            grillaFactura.TextMatrix(I, 4) = Valido_Importe(grdGrilla.TextMatrix(grdGrilla.RowSel, 4))
                            txtSaldoNC.Text = Valido_Importe(CStr(CDbl(txtSaldoNC.Text) + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
                           Exit For
                        End If
                    Next

                    grdGrilla.RemoveItem grdGrilla.RowSel
                Else
                    For I = 1 To grillaFactura.Rows - 1
                        If grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grillaFactura.TextMatrix(I, 5) And _
                           grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = grillaFactura.TextMatrix(I, 1) And _
                           grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = grillaFactura.TextMatrix(I, 2) Then
            
                            grillaFactura.TextMatrix(I, 4) = Valido_Importe(grdGrilla.TextMatrix(grdGrilla.RowSel, 4))
                            txtSaldoNC.Text = Valido_Importe(CStr(CDbl(txtSaldoNC.Text) + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
                           Exit For
                        End If
                    Next
                    grdGrilla.Rows = 1
                    'grdGrilla.HighLight = flexHighlightNever
                End If
        End If
        grdGrilla.SetFocus
    End If
End Sub

Private Function BuscoSaldoFactura(TIPO As String, Numero As String, Sucursal As String, Fecha As String, Rep As String) As Double
    sql = "SELECT FCL_SALDO"
    sql = sql & " FROM FACTURA_CLIENTE"
    sql = sql & " WHERE"
    sql = sql & " TCO_CODIGO=" & XN(TIPO)
    sql = sql & " AND FCL_NUMERO=" & XN(Numero)
    sql = sql & " AND FCL_SUCURSAL=" & XN(Sucursal)
    sql = sql & " AND REP_CODIGO=" & XN(Rep)
    sql = sql & " AND FCL_FECHA=" & XDQ(Fecha)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        BuscoSaldoFactura = rec!FCL_SALDO
    Else
        BuscoSaldoFactura = 0
    End If
    rec.Close
End Function

Private Sub ActualizoSaldoFacturas(TIPO As String, Numero As String, Sucursal As String, Fecha As String, Saldo As String, Rep As String)
    'ACTUALIZO EL SALDO DE LAS FACTURAS ELEGIDAS
    sql = "UPDATE FACTURA_CLIENTE"
    sql = sql & " SET FCL_SALDO=" & XN(Saldo)
    sql = sql & " WHERE"
    sql = sql & " TCO_CODIGO=" & XN(TIPO)
    sql = sql & " AND FCL_NUMERO=" & XN(Numero)
    sql = sql & " AND FCL_SUCURSAL=" & XN(Sucursal)
    sql = sql & " AND REP_CODIGO=" & XN(Rep)
    sql = sql & " AND FCL_FECHA=" & XDQ(Fecha)
    
    DBConn.Execute sql
End Sub

Private Sub txtCliRazSoc_GotFocus()
    SelecTexto txtCliRazSoc
End Sub

Private Sub txtCodCliente_Change()
    If txtCodCliente.Text = "" Then
        txtCliRazSoc.Text = ""
        txtprovincia.Text = ""
        txtCliLocalidad.Text = ""
        txtDomici.Text = ""
        txtNroNotaCredito.Text = ""
        txtNroSucursal.Text = ""
        FechaNotaCredito.Value = Null
        cboNotaCredito.ListIndex = 0
        grillaFactura.Rows = 1
        grillaFactura.HighLight = flexHighlightNever
        CmdNuevo_Click
    End If
End Sub

Private Sub txtCodCliente_GotFocus()
    SelecTexto txtCodCliente
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodCliente_LostFocus()
    If txtCodCliente.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        Rec1.Open BuscoCliente(txtCodCliente), DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtCliRazSoc.Text = Rec1!CLI_RAZSOC
            txtprovincia.Text = Rec1!PRO_DESCRI
            txtCliLocalidad.Text = Rec1!LOC_DESCRI
            txtDomici.Text = IIf(IsNull(Rec1!CLI_DOMICI), "", Rec1!CLI_DOMICI)
            If vAccion = 0 Then
                Call BuscarFacturas(txtCodCliente.Text)
            Else
                Call BuscaFacImputadas(txtCodCliente.Text)
            End If
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
        Rec1.Close
    End If
End Sub
Private Sub BuscaFacImputadas(Cliente As String)
grillaFactura.Rows = 1
    grillaFactura.HighLight = flexHighlightNever
    
    'If TipoBusquedaDoc = 0 Then
    ' ACA HAY QUE HACER ALGO PARA QUE BUSQUE EL SALDO DEL CLIENTE
        sql = "SELECT FL.TCO_CODIGO,TC.TCO_ABREVIA,FL.FCL_NUMERO,FL.FCL_SUCURSAL,"
        sql = sql & "FL.FCL_FECHA,FL.FCL_TOTAL,FL.FCL_SALDO" ',REP_CODIGO"
        sql = sql & " FROM FACTURAS_NOTA_CREDITO_CLIENTE FNC, FACTURA_CLIENTE FL"
        sql = sql & " ,TIPO_COMPROBANTE TC"
        sql = sql & " WHERE "
        sql = sql & " TC.TCO_CODIGO = FL.TCO_CODIGO"
        sql = sql & " AND FNC.FCL_TCO_CODIGO = FL.TCO_CODIGO"
        sql = sql & " AND FNC.FCL_SUCURSAL = FL.FCL_SUCURSAL"
        sql = sql & " AND FNC.FCL_NUMERO = FL.FCL_NUMERO"
        sql = sql & " AND CLI_CODIGO = " & XN(Cliente)
        sql = sql & " AND FNC.NCC_NUMERO = " & XN(txtNroNotaCredito.Text)
        sql = sql & " AND FNC.NCC_SUCURSAL = " & XN(txtNroSucursal.Text)
        sql = sql & " AND FNC.NCC_FECHA = " & XDQ(FechaNotaCredito.Value)

'    Else
'        sql = "SELECT TCO_CODIGO,TCO_ABREVIA,FCL_NUMERO,FCL_SUCURSAL,"
'        sql = sql & "FCL_FECHA,FCL_TOTAL,FCL_SALDO" ',REP_CODIGO"
'        sql = sql & " FROM SALDO_FACTURAS_CLIENTE_V"
'        sql = sql & " WHERE CLI_CODIGO=" & XN(Cliente)
'    End If
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            grillaFactura.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") _
                                  & Chr(9) & rec!FCL_FECHA & Chr(9) & Valido_Importe(rec!FCL_TOTAL) & Chr(9) & _
                                  Valido_Importe(rec!FCL_SALDO) & Chr(9) & rec!TCO_CODIGO
                                  
            rec.MoveNext
        Loop
        grillaFactura.HighLight = flexHighlightAlways
    Else
        MsgBox "El Cliente no tiene facturas imputadas a la Nota de Credito", vbExclamation, TIT_MSGBOX
    End If
    rec.Close
End Sub

Private Sub BuscarFacturas(Cliente As String)
    
    grillaFactura.Rows = 1
    grillaFactura.HighLight = flexHighlightNever
    
    sql = "SELECT TCO_CODIGO,TCO_ABREVIA,FCL_NUMERO,FCL_SUCURSAL,"
    sql = sql & "FCL_FECHA,FCL_TOTAL,FCL_SALDO" ',REP_CODIGO"
    sql = sql & " FROM SALDO_FACTURAS_CLIENTE_V"
    sql = sql & " WHERE CLI_CODIGO=" & XN(Cliente)
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            grillaFactura.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") _
                                  & Chr(9) & rec!FCL_FECHA & Chr(9) & Valido_Importe(rec!FCL_TOTAL) & Chr(9) & _
                                  Valido_Importe(rec!FCL_SALDO) & Chr(9) & rec!TCO_CODIGO
                                  
            rec.MoveNext
        Loop
        grillaFactura.HighLight = flexHighlightAlways
    Else
        MsgBox "El Cliente no tiene facturas con Saldo", vbExclamation, TIT_MSGBOX
    End If
    rec.Close
End Sub

Private Sub txtCliRazSoc_Change()
    If txtCliRazSoc.Text = "" Then
        txtCodCliente.Text = ""
        txtprovincia.Text = ""
        txtCliLocalidad.Text = ""
        txtDomici.Text = ""
    End If
End Sub

Private Sub txtCliRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtCliRazSoc_LostFocus()
    If ActiveControl.Name = "txtCodCliente" Then Exit Sub
    
    If txtCodCliente.Text = "" And txtCliRazSoc.Text <> "" Then
        rec.Open BuscoCliente(txtCliRazSoc), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 1
                frmBuscar.TxtDescriB.Text = txtCliRazSoc.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 0
                    txtCodCliente.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 1
                    txtCliRazSoc.Text = frmBuscar.grdBuscar.Text
                    txtCodCliente_LostFocus
                    FechaDesde.SetFocus
                Else
                    txtCodCliente.SetFocus
                End If
            Else
                txtCodCliente.Text = rec!CLI_CODIGO
                txtCliRazSoc.Text = rec!CLI_RAZSOC
                txtCodCliente_LostFocus
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
        rec.Close
    ElseIf txtCodCliente.Text = "" And txtCliRazSoc.Text = "" Then
        MsgBox "Debe elegir un cliente", vbExclamation, TIT_MSGBOX
        txtCodCliente.SetFocus
    End If
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

Private Sub chkTipoFactura_Click()
    If chkTipoFactura.Value = Checked Then
        cboNotaCredito1.Enabled = True
        cboNotaCredito1.ListIndex = 0
    Else
        cboNotaCredito1.Enabled = False
        cboNotaCredito1.ListIndex = -1
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

Private Sub CmdBuscAprox_Click()
    If txtCliente.Text = "" Then
        MsgBox "Debe elegir un Cliente", vbExclamation, TIT_MSGBOX
        chkCliente.SetFocus
        Exit Sub
    End If
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
     'esto funciona bien para las NC
     If TipoBusquedaDoc = 1 Then
        sql = "SELECT NC.*, C.CLI_RAZSOC, TC.TCO_ABREVIA" ', R.REP_RAZSOC"
        sql = sql & " FROM NOTA_CREDITO_CLIENTE NC,"
        sql = sql & " TIPO_COMPROBANTE TC, CLIENTE C"  ', REPRESENTADA R"
        sql = sql & " WHERE"
        sql = sql & " NC.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND NC.CLI_CODIGO=C.CLI_CODIGO"
        'sql = sql & " AND NC.REP_CODIGO=R.REP_CODIGO"
        If txtCliente.Text <> "" Then sql = sql & " AND NC.CLI_CODIGO=" & XN(txtCliente)
        If Not IsNull(FechaDesde) Then sql = sql & " AND NC.NCC_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND NC.NCC_FECHA<=" & XDQ(FechaHasta)
        If chkTipoFactura.Value = Checked Then sql = sql & " AND NC.TCO_CODIGO=" & XN(cboNotaCredito1.ItemData(cboNotaCredito1.ListIndex))
        If chkRepresentada.Value = Checked Then sql = sql & " AND NC.REP_CODIGO=" & XN(cboBuscaRep.ItemData(cboBuscaRep.ListIndex))
        sql = sql & " ORDER BY NC.NCC_SUCURSAL,NC.NCC_NUMERO"
    Else
        sql = "SELECT DISTINCT NC.*, C.CLI_RAZSOC, TC.TCO_ABREVIA" ', R.REP_RAZSOC"
        sql = sql & " FROM FACTURAS_NOTA_CREDITO_CLIENTE NCC,NOTA_CREDITO_CLIENTE NC,"
        sql = sql & " TIPO_COMPROBANTE TC, CLIENTE C"  ', REPRESENTADA R"
        sql = sql & " WHERE"
        sql = sql & " NCC.TCO_CODIGO=NC.TCO_CODIGO"
        sql = sql & " AND NCC.NCC_NUMERO=NC.NCC_NUMERO"
        sql = sql & " AND NCC.NCC_SUCURSAL=NC.NCC_SUCURSAL"
        sql = sql & " AND NCC.NCC_FECHA=NC.NCC_FECHA"
        sql = sql & " AND NC.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND NC.CLI_CODIGO=C.CLI_CODIGO"
        'sql = sql & " AND NC.REP_CODIGO=R.REP_CODIGO"
        If txtCliente.Text <> "" Then sql = sql & " AND NC.CLI_CODIGO=" & XN(txtCliente)
        If Not IsNull(FechaDesde) Then sql = sql & " AND NC.NCC_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND NC.NCC_FECHA<=" & XDQ(FechaHasta)
        If chkTipoFactura.Value = Checked Then sql = sql & " AND NC.TCO_CODIGO=" & XN(cboNotaCredito1.ItemData(cboNotaCredito1.ListIndex))
        If chkRepresentada.Value = Checked Then sql = sql & " AND NC.REP_CODIGO=" & XN(cboBuscaRep.ItemData(cboBuscaRep.ListIndex))
        sql = sql & " ORDER BY NC.NCC_SUCURSAL,NC.NCC_NUMERO"
    
    End If
    
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!NCC_SUCURSAL, "0000") & "-" & Format(rec!NCC_NUMERO, "00000000") _
                            & Chr(9) & rec!NCC_FECHA & Chr(9) & Format(rec!NCC_TOTAL, "#0.00") & Chr(9) & rec!CLI_RAZSOC & Chr(9) & "" _
                            & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!NCC_TOTAL _
                            & Chr(9) & rec!NCC_SALDO & Chr(9) & rec!TCO_CODIGO _
                            & Chr(9) & rec!CLI_CODIGO & Chr(9) & ""
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

Private Sub cmdGrabar_Click()
    
    If ValidarNotaCredito = False Then Exit Sub
    If MsgBox("¿Confirma la imputación de la Nota de Crédito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayErrorFactura
    
    DBConn.BeginTrans
    sql = "SELECT * FROM FACTURAS_NOTA_CREDITO_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
    sql = sql & " AND NCC_NUMERO = " & XN(txtNroNotaCredito)
    sql = sql & " AND NCC_SUCURSAL = " & XN(txtNroSucursal)
    'sql = sql & " AND REP_CODIGO = " & XN(cboRep.ItemData(cboRep.ListIndex))
    sql = sql & " AND NCC_FECHA=" & XDQ(FechaNotaCredito)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = False Then
        'MODIFICA LA IMPUTACION
        If MsgBox("Seguro que modificar la imputación de la Nota de Crédito Nro.: " & Trim(txtNroNotaCredito), vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
        
            'FACTURAS POR NOTA DE CREDITO
            sql = "DELETE FROM FACTURAS_NOTA_CREDITO_CLIENTE"
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
            sql = sql & " AND NCC_NUMERO=" & XN(txtNroNotaCredito)
            sql = sql & " AND NCC_SUCURSAL = " & XN(txtNroSucursal)
            'sql = sql & " AND REP_CODIGO = " & XN(cboRep.ItemData(cboRep.ListIndex))
            sql = sql & " AND NCC_FECHA=" & XDQ(FechaNotaCredito)
            DBConn.Execute sql
         Else
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
            rec.Close
            DBConn.CommitTrans
            Exit Sub
         End If
    End If
    rec.Close
    
    'FACTURAS POR NOTA DE CREDITO
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            sql = "INSERT INTO FACTURAS_NOTA_CREDITO_CLIENTE"
            sql = sql & " (TCO_CODIGO,NCC_NUMERO,NCC_SUCURSAL,NCC_FECHA,"
            sql = sql & " FCL_TCO_CODIGO,FCL_NUMERO,FCL_SUCURSAL,FCL_FECHA,FNC_IMPORTE)"
            sql = sql & " VALUES ("
            'sql = sql & XN(cboRep.ItemData(cboRep.ListIndex)) & ","
            sql = sql & XN(cboNotaCredito.ItemData(cboNotaCredito.ListIndex)) & ","
            sql = sql & XN(txtNroNotaCredito) & ","
            sql = sql & XN(txtNroSucursal) & ","
            sql = sql & XDQ(FechaNotaCredito) & ","
            sql = sql & XN(grdGrilla.TextMatrix(I, 6)) & ","
            sql = sql & XN(Right(grdGrilla.TextMatrix(I, 1), 8)) & ", "
            sql = sql & XN(Left(grdGrilla.TextMatrix(I, 1), 4)) & ", "
            sql = sql & XDQ(grdGrilla.TextMatrix(I, 2)) & ","
            sql = sql & XN(grdGrilla.TextMatrix(I, 5)) & ")"
            DBConn.Execute sql
        End If
    Next
            
    'ACTUALIZO EL SALDO DE LA NOTA DE CREDITO
    sql = "UPDATE NOTA_CREDITO_CLIENTE"
    sql = sql & " SET NCC_SALDO=" & XN(txtSaldoNC.Text)
    sql = sql & " WHERE"
    sql = sql & " TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
    sql = sql & " AND NCC_NUMERO=" & XN(txtNroNotaCredito)
    sql = sql & " AND NCC_SUCURSAL = " & XN(txtNroSucursal)
    'sql = sql & " AND REP_CODIGO = " & XN(cboRep.ItemData(cboRep.ListIndex))
    sql = sql & " AND NCC_FECHA=" & XDQ(FechaNotaCredito)
    DBConn.Execute sql
    
    'ACTUALIZO EL SALDO DE LAS FACTURAS ELEGIDAS
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 7) <> "X" Then
            sql = "UPDATE FACTURA_CLIENTE"
            sql = sql & " SET FCL_SALDO=" & XN(CStr(CDbl(grdGrilla.TextMatrix(I, 4)) - CDbl(grdGrilla.TextMatrix(I, 5))))
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & XN(grdGrilla.TextMatrix(I, 6))
            sql = sql & " AND FCL_NUMERO=" & XN(Right(grdGrilla.TextMatrix(I, 1), 8))
            sql = sql & " AND FCL_SUCURSAL=" & XN(Left(grdGrilla.TextMatrix(I, 1), 4))
            'sql = sql & " AND REP_CODIGO=" & XN(cboRep.ItemData(cboRep.ListIndex))
            sql = sql & " AND FCL_FECHA=" & XDQ(grdGrilla.TextMatrix(I, 2))
            DBConn.Execute sql
        End If
    Next
    DBConn.CommitTrans
        
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

Private Function ValidarNotaCredito() As Boolean
    If txtNroNotaCredito.Text = "" Then
        MsgBox "La Nota de Crédito es requerida", vbExclamation, TIT_MSGBOX
        txtNroNotaCredito.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If IsNull(FechaNotaCredito.Value) Then
        MsgBox "La Fecha de la Nota de Crédito es requerida", vbExclamation, TIT_MSGBOX
        FechaNotaCredito.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If txtCodCliente.Text = "" Then
        MsgBox "Debe ingresar un Cliente", vbExclamation, TIT_MSGBOX
        txtCodCliente.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If grillaFactura.Rows = 1 Then
        MsgBox "El Cliente no tiene Facturas con Saldo", vbExclamation, TIT_MSGBOX
        txtCodCliente.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    ValidarNotaCredito = True
End Function

Private Sub CmdNuevo_Click()
   grdGrilla.Rows = 1
   txtCodCliente.Text = ""
   txtNroNotaCredito.Text = ""
   txtNroSucursal.Text = ""
   'cboRep.ListIndex = 0
   FechaNotaCredito.Value = Null
   lblEstadoNotaCredito.Caption = ""
   lblEstado.Caption = ""
   txtTotalNC.Text = ""
   txtSaldoNC.Text = ""
   cmdGrabar.Enabled = True
   'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoNotaCredito) 'ESTADO PENDIENTE
    VEstadoNotaCredito = 1
    '--------------
    FrameNotaCredito.Enabled = True
    FrameCliente.Enabled = True
    tabDatos.Tab = 0
    cboNotaCredito.ListIndex = 0
    txtCodCliente.SetFocus
    vAccion = 0
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmImputarNCaFactura = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        TipoBusquedaDoc = 0 'BUSCA Imputaciones
        GrdModulos.ColWidth(0) = 1300 'TIPO NOTA CREDITO O TIPO FACTURA
        chkTipoFactura.Visible = True
        lbltipoFac.Visible = True
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
    'Por defecto busca IMputaciones
    TipoBusquedaDoc = 0
    vAccion = 0
    
    grdGrilla.FormatString = "^Comp.|^Número|^Fecha|>Total|>Saldo|>Importe|Tipo Comp.|MARCA"
    grdGrilla.ColWidth(0) = 800 'COMPROBANTE
    grdGrilla.ColWidth(1) = 1300 'NUMERO
    grdGrilla.ColWidth(2) = 1100 'FECHA
    grdGrilla.ColWidth(3) = 1000 'TOTAL
    grdGrilla.ColWidth(4) = 0    'SALDO
    grdGrilla.ColWidth(5) = 1000 'IMPORTE A ASIGNAR
    grdGrilla.ColWidth(6) = 0    'TIPO COMPROBANTE
    grdGrilla.ColWidth(7) = 0    'MARCA
    grdGrilla.Cols = 8
    grdGrilla.Rows = 1
    
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "^Tipo|^Número|^Fecha|Importe|Cliente|Representada|Cod_Estado|" _
                              & "TOTAL|SALDO|COD TIPO COMPROBANTE FACTURA O NC|" _
                              & "COD CLIENTE|REPRESENTADA"
    GrdModulos.ColWidth(0) = 900 'TIPO NOTA CREDITO O TIPO FACTURA
    GrdModulos.ColWidth(1) = 1300 'NUMERO
    GrdModulos.ColWidth(2) = 1000 'FECHA
    GrdModulos.ColWidth(3) = 1100 'IMPORTE
    GrdModulos.ColWidth(4) = 4000 'CLIENTE
    GrdModulos.ColWidth(5) = 0    'REPRESENTADA
    GrdModulos.ColWidth(6) = 0    'COD_ESTADO
    GrdModulos.ColWidth(7) = 0    'TOTAL
    GrdModulos.ColWidth(8) = 0    'SALDO
    GrdModulos.ColWidth(9) = 0    'COD TIPO COMPROBANTE NOTA CREDITO O FACTURA
    GrdModulos.ColWidth(10) = 0    'CODIGO CLIENTE
    GrdModulos.ColWidth(11) = 0   'REPRESENTADA
    GrdModulos.Cols = 12
    GrdModulos.Rows = 1
    'GRILLA PARA MOSTRAR LAS FACTURAS PENDIENTES DE LOS CLIENTES
    grillaFactura.FormatString = "^Comp.|^Número|^Fecha|>Total|>Saldo|Tipo Comp."
    grillaFactura.ColWidth(0) = 800 'COMPROBANTE
    grillaFactura.ColWidth(1) = 1300 'NUMERO
    grillaFactura.ColWidth(2) = 1100 'FECHA
    grillaFactura.ColWidth(3) = 1000 'TOTAL
    grillaFactura.ColWidth(4) = 1000 'SALDO
    grillaFactura.ColWidth(5) = 0    'TIPO COMPROBANTE
    grillaFactura.Rows = 1
    grillaFactura.Cols = 6
     grillaFactura.HighLight = flexHighlightNever
    '------------------------------------
    lblEstado.Caption = ""
    'CARGO COMBO CON LOS TIPOS DE NOTA DE CREDITO
    LlenarComboNotaCredito
    
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoNotaCredito) 'ESTADO PENDIENTE
    VEstadoNotaCredito = 1
    TipoBusquedaDoc = 1 'ESTO ES PARA BUSCAR FACTURA(1), (2)PARA BUSCAR REMITOS
    tabDatos.Tab = 0
End Sub
Private Sub LlenarComboNotaCredito()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'NOTA DE CRE%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboNotaCredito.AddItem rec!TCO_DESCRI
            cboNotaCredito.ItemData(cboNotaCredito.NewIndex) = rec!TCO_CODIGO
            cboNotaCredito1.AddItem rec!TCO_DESCRI
            cboNotaCredito1.ItemData(cboNotaCredito1.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboNotaCredito.ListIndex = 0
        cboNotaCredito1.ListIndex = -1
    End If
    rec.Close
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If grdGrilla.Rows > 1 Then
        If KeyCode = vbKeyDelete Then
            Select Case grdGrilla.Col
            Case 5
                grdGrilla.Col = 5
                grdGrilla.Text = ""
                grdGrilla.Col = 5
            End Select
        End If
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If grdGrilla.Rows > 1 Then
        If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
           (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Or _
           (grdGrilla.Col = 4) Or (grdGrilla.Col = 5) Then
            If KeyAscii = vbKeyReturn Then
                If grdGrilla.Col = 5 Then
                    If grdGrilla.row < grdGrilla.Rows - 1 Then
                        grdGrilla.row = grdGrilla.row + 1
                        grdGrilla.Col = 5
                    Else
                        SendKeys "{TAB}"
                    End If
                Else
                    grdGrilla.Col = grdGrilla.Col + 1
                End If
            Else
                If grdGrilla.Col = 5 Then
                    If KeyAscii > 47 And KeyAscii < 58 And grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = "" Then
                        EDITAR grdGrilla, txtEdit, KeyAscii
                    End If
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
            grdGrilla.Col = 5
            Exit Sub
        End If
        grdGrilla = txtEdit.Text
        txtEdit.Visible = False
    End If
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        Set Rec1 = New ADODB.Recordset
        
        Select Case TipoBusquedaDoc
        Case 0 'BUSCA IMPUTACIONES
            lblEstado.Caption = "Buscando..."
            Screen.MousePointer = vbHourglass
            'CABEZA NOTA CREDITO
            Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 9)), cboNotaCredito)
            txtNroNotaCredito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
            txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
            FechaNotaCredito.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
            Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6)), lblEstadoNotaCredito)
            VEstadoNotaCredito = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6))
            txtCodCliente.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 10)
            vAccion = 1
            txtCodCliente_LostFocus
            txtTotalNC.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 7))
            txtSaldoNC.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 8))
            txtNroNotaCredito_LostFocus
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            '--------------
            FrameNotaCredito.Enabled = False
            FrameCliente.Enabled = False
            '--------------
            tabDatos.Tab = 0
            '----------------------------------------------------------
        
        Case 1 'BUSCA NOTA CREDITO
            lblEstado.Caption = "Buscando..."
            Screen.MousePointer = vbHourglass
            'CABEZA NOTA CREDITO
            Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 9)), cboNotaCredito)
            txtNroNotaCredito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
            txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
            FechaNotaCredito.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
            Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6)), lblEstadoNotaCredito)
            VEstadoNotaCredito = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6))
            txtCodCliente.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 10)
            txtCodCliente_LostFocus
            txtTotalNC.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 7))
            txtSaldoNC.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 8))
            txtNroNotaCredito_LostFocus
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            '--------------
            FrameNotaCredito.Enabled = False
            FrameCliente.Enabled = False
            '--------------
            tabDatos.Tab = 0
            '----------------------------------------------------------
        End Select
    End If
End Sub

Private Function BuscarTipoDocAbre(Codigo As String) As String
    sql = "SELECT TCO_ABREVIA"
    sql = sql & " FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_CODIGO=" & XN(Codigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscarTipoDocAbre = rec!TCO_ABREVIA
    Else
        BuscarTipoDocAbre = ""
    End If
    rec.Close
End Function
Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    If TipoBusquedaDoc = 1 Then
        frameBuscar.Caption = "Buscar Notas de Credito por...."
    Else
        frameBuscar.Caption = "Buscar Imputaciones de NC por...."
    End If
    
    GrdModulos.Rows = 2
    txtCliente.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    txtVendedor.Enabled = False
    cboNotaCredito1.Enabled = False
    cmdBuscarCli.Enabled = False
    cmdBuscarVen.Enabled = False
    cmdGrabar.Enabled = False
    LimpiarBusqueda
    
    chkVendedor.Enabled = False
    txtVendedor.Enabled = False
    chkCliente.Value = Checked
    txtCliente.Text = txtCodCliente.Text
    txtCliente.Enabled = True
    txtCliente_LostFocus
    cmdBuscarCli.Enabled = False
    'If Me.Visible = True Then chkCliente.SetFocus
  Else
    TipoBusquedaDoc = 0
    chkTipoFactura.Visible = True
    lbltipoFac.Visible = True
    If VEstadoNotaCredito = 2 Then
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

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    If grdGrilla.Col = 5 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
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
            Case 5 'IMPORTE
                If txtEdit.Text <> "" Then
                    txtEdit.Text = Valido_Importe(txtEdit)
                    If CDbl(Trim(txtEdit.Text)) <= CDbl(Trim(txtSaldoNC.Text)) Then
                        If CDbl(Trim(txtEdit.Text)) <= CDbl(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 4))) Then
                            For I = 1 To grillaFactura.Rows - 1
                                If grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grillaFactura.TextMatrix(I, 5) And _
                                   grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = grillaFactura.TextMatrix(I, 1) And _
                                   grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = grillaFactura.TextMatrix(I, 2) Then
                                   
                                    grillaFactura.TextMatrix(I, 4) = Valido_Importe(CStr(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 4)) - CDbl(txtEdit.Text)))
                                    txtSaldoNC.Text = Valido_Importe(CStr(CDbl(txtSaldoNC.Text) - CDbl(txtEdit.Text)))
                                   Exit For
                                End If
                            Next
                        Else
                            MsgBox "No puede ingresar un importe Mayor al Saldo de la Factura", vbExclamation, TIT_MSGBOX
                            txtEdit.Text = ""
                        End If
                    Else
                        MsgBox "No puede ingresar un importe Mayor al Saldo de la Nota de Crédito", vbExclamation, TIT_MSGBOX
                        txtEdit.Text = ""
                    End If
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

'Private Sub LimpiarFactura()
'    txtNroFactura.Text = ""
'    FechaFactura.Text = ""
'    grillaFactura.TextMatrix(0, 1) = ""
'    grillaFactura.TextMatrix(1, 1) = ""
'    grillaFactura.TextMatrix(2, 1) = ""
'    txtNroFactura.SetFocus
'End Sub

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

Private Function BuscoCliente(Cli As String) As String
    sql = "SELECT C.CLI_CODIGO, C.CLI_RAZSOC, C.CLI_DOMICI, P.PRO_DESCRI, L.LOC_DESCRI"
    sql = sql & " FROM CLIENTE C,  PROVINCIA P, LOCALIDAD L"
    sql = sql & " WHERE"
    If txtCodCliente.Text <> "" Then
        sql = sql & " C.CLI_CODIGO=" & XN(Cli)
    Else
        sql = sql & " C.CLI_RAZSOC LIKE '" & Cli & "%'"
    End If
    sql = sql & " AND C.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"

    BuscoCliente = sql
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

'Private Sub txtNroFactura_GotFocus()
'    SelecTexto txtNroFactura
'End Sub
'
'Private Sub txtNroFactura_KeyPress(KeyAscii As Integer)
'    KeyAscii = CarNumeroEntero(KeyAscii)
'End Sub
'
'Private Sub txtNroFactura_LostFocus()
'    If txtNroFactura.Text <> "" Then
'
'        Set Rec2 = New ADODB.Recordset
'        sql = "SELECT FC.FCL_NUMERO, FC.FCL_FECHA, FC.FCL_BONIFICA, FC.FCL_IVA, FC.FCL_BONIPESOS,"
'        sql = sql & "RC.EST_CODIGO, RC.STK_CODIGO, E.EST_DESCRI"
'        sql = sql & " ,NP.CLI_CODIGO, NP.SUC_CODIGO, NP.VEN_CODIGO"
'        sql = sql & " FROM FACTURA_CLIENTE FC, REMITO_CLIENTE RC, NOTA_PEDIDO NP, ESTADO_DOCUMENTO E"
'        sql = sql & " WHERE FC.FCL_NUMERO=" & XN(txtNroFactura)
'        If FechaFactura.Text <> "" Then
'            sql = sql & " AND FC.FCL_FECHA=" & XDQ(FechaFactura)
'        End If
'        'sql = sql & " AND FC.TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
'        sql = sql & " AND FC.RCL_NUMERO=RC.RCL_NUMERO"
'        sql = sql & " AND FC.RCL_FECHA=RC.RCL_FECHA"
'        sql = sql & " AND RC.NPE_NUMERO=NP.NPE_NUMERO"
'        sql = sql & " AND RC.NPE_FECHA=NP.NPE_FECHA"
'        sql = sql & " AND FC.EST_CODIGO=E.EST_CODIGO"
'
'        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'        If Rec2.EOF = False Then
'            If Rec2.RecordCount > 1 Then
'                MsgBox "Hay mas de una Factura con el Número: " & txtNroFactura.Text, vbInformation, TIT_MSGBOX
'                Rec2.Close
'                cmdBuscarFactura_Click
'                Exit Sub
'            End If
'            Screen.MousePointer = vbHourglass
'            lblEstado.Caption = "Buscando..."
'
'            'CARGO CABECERA DE LA FACTURA
'            FechaFactura.Text = Rec2!FCL_FECHA
'            grillaFactura.TextMatrix(0, 1) = BuscoCliente(Rec2!CLI_CODIGO)
'            grillaFactura.TextMatrix(1, 1) = BuscoSucursal(Rec2!SUC_CODIGO, Rec2!CLI_CODIGO)
'            grillaFactura.TextMatrix(2, 1) = BuscoVendedor(Rec2!VEN_CODIGO)
'            grillaFactura.TextMatrix(0, 2) = Rec2!CLI_CODIGO
'            If Rec2!EST_CODIGO = 2 Then
'                MsgBox "La Factura número: " & txtNroFactura.Text & Chr(13) & Chr(13) & _
'                       "No puede ser asignado a la Nota de Crédito por su estado (" & Rec2!EST_DESCRI & ")", vbExclamation, TIT_MSGBOX
'                cmdGrabar.Enabled = False
'                Screen.MousePointer = vbNormal
'                lblEstado.Caption = ""
'                Rec2.Close
'                LimpiarFactura
'                Exit Sub
'            Else
'                cmdGrabar.Enabled = True
'            End If
'
'            '----BUSCO DETALLE DE LA FACTURA------------------
'            sql = "SELECT DFC.*, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
'            sql = sql & " FROM DETALLE_FACTURA_CLIENTE DFC, PRODUCTO P, RUBROS R, LINEAS L"
'            sql = sql & " WHERE DFC.FCL_NUMERO=" & XN(txtNroFactura)
'            sql = sql & " AND DFC.FCL_FECHA=" & XDQ(FechaFactura)
'            'sql = sql & " AND DFC.TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
'            sql = sql & " AND DFC.PTO_CODIGO=P.PTO_CODIGO"
'            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
'            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
'            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
'            sql = sql & " ORDER BY DFC.DFC_NROITEM"
'            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'            If Rec1.EOF = False Then
'                I = 1
'                Do While Rec1.EOF = False
'                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
'                    grdGrilla.TextMatrix(I, 1) = Rec1!PTO_DESCRI
'                    grdGrilla.TextMatrix(I, 2) = Rec1!DFC_CANTIDAD
'                    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DFC_PRECIO)
'                    If IsNull(Rec1!DFC_BONIFICA) Then
'                        grdGrilla.TextMatrix(I, 4) = ""
'                    Else
'                        grdGrilla.TextMatrix(I, 4) = Valido_Importe(Rec1!DFC_BONIFICA)
'                    End If
'                    VBonificacion = 0
'                    If Not IsNull(Rec1!DFC_BONIFICA) Then
'                        VBonificacion = (((CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO)) * CDbl(Rec1!DFC_BONIFICA)) / 100)
'                        VBonificacion = ((CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO)) - VBonificacion)
'                        grdGrilla.TextMatrix(I, 5) = Valido_Importe(CStr(VBonificacion))
'                        grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
'                    Else
'                        VBonificacion = (CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO))
'                        grdGrilla.TextMatrix(I, 5) = ""
'                        grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
'                    End If
'                    grdGrilla.TextMatrix(I, 7) = Rec1!RUB_DESCRI
'                    grdGrilla.TextMatrix(I, 8) = Rec1!LNA_DESCRI
'                    grdGrilla.TextMatrix(I, 9) = Rec1!DFC_NROITEM
'                    I = I + 1
'                    Rec1.MoveNext
'                Loop
'                VBonificacion = 0
'            End If
'            Rec1.Close
'            '--CARGO LOS TOTALES----
'            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
'            txtTotal.Text = txtSubtotal.Text
'            If Rec2!FCL_BONIPESOS = "S" Then
'                chkBonificaEnPesos.Value = Checked
'            ElseIf Rec2!FCL_BONIPESOS = "N" Then
'                chkBonificaEnPorsentaje.Value = Checked
'            Else
'                chkBonificaEnPesos.Value = Unchecked
'                chkBonificaEnPorsentaje.Value = Unchecked
'            End If
'            If Not IsNull(Rec2!FCL_BONIFICA) Then
'                txtPorcentajeBoni.Text = Rec2!FCL_BONIFICA
'                txtPorcentajeBoni_LostFocus
'            End If
'            If Not IsNull(Rec2!FCL_IVA) Then
'                txtPorcentajeIva = Rec2!FCL_IVA
'                txtPorcentajeIva_LostFocus
'            End If
'            Rec2.Close
'            lblEstado.Caption = ""
'            Screen.MousePointer = vbNormal
'            '--------------
'            FrameNotaCredito.Enabled = False
'            FrameFactura.Enabled = False
'            '--------------
'        Else
'            MsgBox "La Factura no existe", vbExclamation, TIT_MSGBOX
'            If Rec2.State = 1 Then Rec2.Close
'            LimpiarFactura
'        End If
'    End If
'End Sub

Private Function SumaTotal() As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 6) <> "" Then
            VTotal = VTotal + (CInt(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3)))
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

Private Sub txtNroNotaCredito_Change()
    If txtNroNotaCredito.Text = "" Then
        FechaNotaCredito.Value = Null
    End If
End Sub

Private Sub txtNroNotaCredito_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaCredito_LostFocus()
    If txtCodCliente.Text = "" Then
        MsgBox "Debe elegir un Cliente", vbExclamation, TIT_MSGBOX
        txtNroNotaCredito.Text = ""
        txtCodCliente.SetFocus
        Exit Sub
    End If
    txtNroNotaCredito.Text = Format(txtNroNotaCredito.Text, "00000000")
    
    Set Rec1 = New ADODB.Recordset
    
    'ME FIJO SI EXISTE LA NOTA DE CREDITO
    sql = "SELECT NCC_FECHA"
    sql = sql & " FROM NOTA_CREDITO_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & XN(cboNotaCredito.ItemData(cboNotaCredito.ListIndex))
    sql = sql & " AND NCC_NUMERO=" & XN(txtNroNotaCredito)
    sql = sql & " AND NCC_SUCURSAL=" & XN(txtNroSucursal)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Rec1.EOF = True Then
        MsgBox "La Nota de Crédito no Existe", vbExclamation, TIT_MSGBOX
        Rec1.Close
        txtNroSucursal.Text = ""
        txtNroNotaCredito.Text = ""
        cboNotaCredito.SetFocus
        Exit Sub
    End If
    Rec1.Close
    
    grdGrilla.Rows = 1
    'BUSCO LAS FACTURAS RELACIONADAS A LA NOTA DE CREDITO
    sql = "SELECT FN.FCL_TCO_CODIGO, FN.FCL_NUMERO, FN.FCL_SUCURSAL, FN.FCL_FECHA,"
    sql = sql & " FN.FNC_IMPORTE, TC.TCO_ABREVIA, FC.FCL_TOTAL, FC.FCL_SALDO"
    sql = sql & " FROM FACTURAS_NOTA_CREDITO_CLIENTE FN, FACTURA_CLIENTE FC, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " FN.TCO_CODIGO=" & XN(cboNotaCredito.ItemData(cboNotaCredito.ListIndex))
    sql = sql & " AND FN.NCC_NUMERO=" & XN(txtNroNotaCredito)
    sql = sql & " AND FN.NCC_SUCURSAL=" & XN(txtNroSucursal)
    
    sql = sql & " AND FN.FCL_TCO_CODIGO=FC.TCO_CODIGO"
    sql = sql & " AND FN.FCL_NUMERO=FC.FCL_NUMERO"
    sql = sql & " AND FN.FCL_SUCURSAL=FC.FCL_SUCURSAL"
    'sql = sql & " AND FN.REP_CODIGO=FC.REP_CODIGO"
    sql = sql & " AND FN.FCL_FECHA=FC.FCL_FECHA"
    sql = sql & " AND FN.FCL_TCO_CODIGO=TC.TCO_CODIGO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            grdGrilla.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!FCL_SUCURSAL, "0000") & "-" & Format(Rec1!FCL_NUMERO, "00000000") _
                                  & Chr(9) & Rec1!FCL_FECHA & Chr(9) & Valido_Importe(Rec1!FCL_TOTAL) & Chr(9) & _
                                  "" & Chr(9) & Valido_Importe(Chk0(Rec1!FNC_IMPORTE)) & Chr(9) & _
                                  Rec1!FCL_TCO_CODIGO & Chr(9) & "X"
            Rec1.MoveNext
            CambiaColorAFilaDeGrilla grdGrilla, grdGrilla.Rows - 1, vbRed
        Loop
        grdGrilla.HighLight = flexHighlightAlways
    End If
    Rec1.Close
    
    'BUSCO EL TOTAL Y SALDO DE LA NOTA DE CREDITO
    sql = "SELECT NCC_FECHA,NCC_TOTAL,NCC_SALDO"
    sql = sql & " FROM NOTA_CREDITO_CLIENTE"
    sql = sql & " WHERE"
    sql = sql & " TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
    sql = sql & " AND NCC_NUMERO=" & XN(txtNroNotaCredito)
    sql = sql & " AND NCC_SUCURSAL=" & XN(txtNroSucursal)
    'sql = sql & " AND REP_CODIGO=" & XN(cboRep.ItemData(cboRep.ListIndex))
    sql = sql & " AND CLI_CODIGO=" & XN(txtCodCliente)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        FechaNotaCredito.Value = rec!NCC_FECHA
        txtTotalNC.Text = Valido_Importe(Chk0(rec!NCC_TOTAL))
        txtSaldoNC.Text = Valido_Importe(Chk0(rec!NCC_SALDO))
    End If
    rec.Close
    'VEO EL SALDO DE LA NOTA DE CREDITO
    If CDbl(txtSaldoNC.Text) > 0 Then
        cmdGrabar.Enabled = True
    Else
        cmdGrabar.Enabled = False
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
    If chkFecha.Value = Unchecked And chkTipoFactura.Value = Unchecked _
    And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub
