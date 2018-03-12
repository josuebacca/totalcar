VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCargaGastosGenerales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga Gastos Generales"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      Height          =   450
      Left            =   6960
      TabIndex        =   21
      Top             =   7155
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7050
      Left            =   45
      TabIndex        =   35
      Top             =   60
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   12435
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
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
      TabPicture(0)   =   "frmCargaGastosGenerales.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameProveedor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmCargaGastosGenerales.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
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
         Left            =   -74835
         TabIndex        =   52
         Top             =   375
         Width           =   8355
         Begin VB.ComboBox cboBuscaTipoGasto 
            Height          =   315
            Left            =   2385
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   960
            Width           =   4125
         End
         Begin VB.ComboBox cboBuscaTipoProveedor 
            Height          =   315
            Left            =   2385
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   270
            Width           =   3750
         End
         Begin VB.CheckBox chkTipoProveedor 
            Caption         =   "Tipo Prov"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   450
            Width           =   1050
         End
         Begin VB.CheckBox chkProveedor 
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   705
            Width           =   1125
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1230
            Width           =   810
         End
         Begin VB.TextBox txtDesProv 
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
            Left            =   3825
            MaxLength       =   50
            TabIndex        =   54
            Tag             =   "Descripción"
            Top             =   615
            Width           =   4440
         End
         Begin VB.TextBox txtProveedor 
            Height          =   300
            Left            =   2385
            MaxLength       =   40
            TabIndex        =   28
            Top             =   615
            Width           =   975
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   465
            Left            =   6810
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmCargaGastosGenerales.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Buscar "
            Top             =   1155
            UseMaskColor    =   -1  'True
            Width           =   1455
         End
         Begin VB.CheckBox chkTipoGasto 
            Caption         =   "Tipo Gasto"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1155
         End
         Begin VB.CommandButton cmdBuscarProveedor 
            Height          =   300
            Left            =   3390
            MaskColor       =   &H000000FF&
            Picture         =   "frmCargaGastosGenerales.frx":27DA
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Buscar Proveedor"
            Top             =   615
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   2400
            TabIndex        =   30
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61734913
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   5040
            TabIndex        =   31
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61734913
            CurrentDate     =   41098
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Gasto:"
            Height          =   195
            Left            =   1860
            TabIndex        =   59
            Top             =   990
            Width           =   465
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Prov.:"
            Height          =   195
            Left            =   1545
            TabIndex        =   58
            Top             =   315
            Width           =   780
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1545
            TabIndex        =   57
            Top             =   645
            Width           =   780
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1320
            TabIndex        =   56
            Top             =   1350
            Width           =   1005
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4080
            TabIndex        =   55
            Top             =   1365
            Width           =   960
         End
      End
      Begin VB.Frame FrameProveedor 
         Caption         =   "Proveedor..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2025
         Left            =   165
         TabIndex        =   45
         Top             =   585
         Width           =   8355
         Begin VB.CommandButton cmdBuscarProveedor1 
            Height          =   300
            Left            =   2295
            MaskColor       =   &H000000FF&
            Picture         =   "frmCargaGastosGenerales.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Buscar Proveedor"
            Top             =   765
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtDomici 
            BackColor       =   &H00C0C0C0&
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
            Left            =   1275
            MaxLength       =   50
            TabIndex        =   47
            Top             =   1425
            Width           =   4860
         End
         Begin VB.TextBox txtCliLocalidad 
            BackColor       =   &H00C0C0C0&
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
            Left            =   1275
            TabIndex        =   46
            Top             =   1110
            Width           =   4860
         End
         Begin VB.TextBox txtProvRazSoc 
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
            Left            =   2730
            MaxLength       =   50
            TabIndex        =   2
            Tag             =   "Descripción"
            Top             =   765
            Width           =   5310
         End
         Begin VB.TextBox txtCodProveedor 
            Height          =   300
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   1
            Top             =   765
            Width           =   975
         End
         Begin VB.ComboBox cboTipoProveedor 
            Height          =   315
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   405
            Width           =   3375
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Dom.:"
            Height          =   195
            Left            =   765
            TabIndex        =   51
            Top             =   1455
            Width           =   420
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Loc.:"
            Height          =   180
            Left            =   825
            TabIndex        =   50
            Top             =   1155
            Width           =   360
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
            Left            =   645
            TabIndex        =   49
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Prov.:"
            Height          =   195
            Left            =   405
            TabIndex        =   48
            Top             =   435
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Comprobantes..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   165
         TabIndex        =   36
         Top             =   2610
         Width           =   8355
         Begin VB.TextBox txtNroSucursal 
            Height          =   285
            Left            =   1275
            MaxLength       =   4
            TabIndex        =   5
            Top             =   1140
            Width           =   480
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   540
            Left            =   1275
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   3600
            Width           =   6825
         End
         Begin VB.ComboBox cboCondicion 
            Height          =   315
            Left            =   4125
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1320
            Width           =   2910
         End
         Begin VB.CommandButton cmdNuevoRubro 
            Height          =   315
            Left            =   7125
            MaskColor       =   &H000000FF&
            Picture         =   "frmCargaGastosGenerales.frx":2DEE
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Agregar Condición de Venta"
            Top             =   1320
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.TextBox txtimp2IVA 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4845
            MaxLength       =   40
            TabIndex        =   70
            Top             =   2145
            Width           =   900
         End
         Begin VB.TextBox txtimp1IVA 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1995
            MaxLength       =   40
            TabIndex        =   69
            Top             =   2160
            Width           =   780
         End
         Begin VB.TextBox txtIva1 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4125
            MaxLength       =   40
            TabIndex        =   12
            Top             =   2145
            Width           =   660
         End
         Begin VB.TextBox txtNeto1 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4125
            MaxLength       =   40
            TabIndex        =   11
            Top             =   1800
            Width           =   1620
         End
         Begin VB.TextBox txtSubTotal1 
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
            Height          =   300
            Left            =   4125
            MaxLength       =   40
            TabIndex        =   63
            Top             =   2490
            Width           =   1620
         End
         Begin VB.TextBox txtSubTotal 
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
            Height          =   300
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   62
            Top             =   2490
            Width           =   1500
         End
         Begin VB.TextBox txtImpuestos 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   13
            Top             =   2835
            Width           =   1500
         End
         Begin VB.CommandButton cmdNuevoGasto 
            Height          =   315
            Left            =   5115
            MaskColor       =   &H000000FF&
            Picture         =   "frmCargaGastosGenerales.frx":3178
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Agregar País"
            Top             =   405
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CheckBox chkCreditoFiscal 
            Caption         =   "Genera Crédito Fiscal"
            Height          =   210
            Left            =   4125
            TabIndex        =   15
            Top             =   2940
            Width           =   1815
         End
         Begin VB.ComboBox cboComprobante 
            Height          =   315
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   780
            Width           =   3375
         End
         Begin VB.TextBox txtNroComprobante 
            Height          =   285
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   6
            Top             =   1140
            Width           =   960
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
            Height          =   300
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   14
            Top             =   3180
            Width           =   1500
         End
         Begin VB.TextBox txtIva 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   10
            Top             =   2160
            Width           =   660
         End
         Begin VB.TextBox txtNeto 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   9
            Top             =   1800
            Width           =   1500
         End
         Begin VB.ComboBox CboGastos 
            Height          =   315
            Left            =   1275
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   405
            Width           =   3765
         End
         Begin MSComCtl2.DTPicker FechaComprobante 
            Height          =   315
            Left            =   1275
            TabIndex        =   7
            Top             =   1440
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61734913
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker Periodo 
            Height          =   315
            Left            =   4125
            TabIndex        =   16
            Top             =   3165
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61734913
            CurrentDate     =   41098
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   3645
            Width           =   1110
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Forma Pago:"
            Height          =   195
            Left            =   3150
            TabIndex        =   72
            Top             =   1365
            Width           =   900
         End
         Begin VB.Label Label11 
            Caption         =   "Neto:"
            Height          =   195
            Left            =   3660
            TabIndex        =   68
            Top             =   1860
            Width           =   390
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Iva:"
            Height          =   195
            Left            =   3780
            TabIndex        =   67
            Top             =   2190
            Width           =   270
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Sub Total:"
            Height          =   195
            Left            =   465
            TabIndex        =   66
            Top             =   2535
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Sub Total:"
            Height          =   195
            Left            =   3315
            TabIndex        =   65
            Top             =   2535
            Width           =   735
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Impuestos:"
            Height          =   195
            Left            =   435
            TabIndex        =   64
            Top             =   2880
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante:"
            Height          =   195
            Left            =   210
            TabIndex        =   44
            Top             =   825
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   600
            TabIndex        =   43
            Top             =   1170
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   705
            TabIndex        =   42
            Top             =   1515
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Neto:"
            Height          =   195
            Left            =   810
            TabIndex        =   41
            Top             =   1860
            Width           =   390
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Iva:"
            Height          =   195
            Left            =   930
            TabIndex        =   40
            Top             =   2190
            Width           =   270
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   795
            TabIndex        =   39
            Top             =   3210
            Width           =   405
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Periodo:"
            Height          =   195
            Left            =   3465
            TabIndex        =   38
            Top             =   3225
            Width           =   585
         End
         Begin VB.Label lblPeriodo1 
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   17
            Top             =   3180
            Width           =   1785
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Gasto:"
            Height          =   195
            Left            =   735
            TabIndex        =   37
            Top             =   450
            Width           =   465
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4695
         Left            =   -74865
         TabIndex        =   33
         Top             =   2160
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   19
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   6075
      TabIndex        =   20
      Top             =   7155
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   7845
      TabIndex        =   22
      Top             =   7155
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   5190
      TabIndex        =   19
      Top             =   7155
      Width           =   870
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
      Left            =   150
      TabIndex        =   34
      Top             =   7275
      Width           =   750
   End
End
Attribute VB_Name = "frmCargaGastosGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rec1 As ADODB.Recordset

Private Sub chkCreditoFiscal_Click()
    If chkCreditoFiscal.Value = Checked Then
        Periodo.Enabled = True
    Else
        Periodo.Enabled = False
    End If
End Sub

Private Sub chkCreditoFiscal_LostFocus()
    If chkCreditoFiscal.Value = Unchecked Then cmdGrabar.SetFocus
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

Private Sub chkProveedor_Click()
    If chkProveedor.Value = Checked Then
        txtProveedor.Enabled = True
        cmdBuscarProveedor.Enabled = True
    Else
        txtProveedor.Text = ""
        txtProveedor.Enabled = False
        cmdBuscarProveedor.Enabled = False
    End If
End Sub

Private Sub chkTipoGasto_Click()
    If chkTipoGasto.Value = Checked Then
        cboBuscaTipoGasto.Enabled = True
        cboBuscaTipoGasto.ListIndex = 0
    Else
        cboBuscaTipoGasto.Enabled = False
        cboBuscaTipoGasto.ListIndex = -1
    End If
End Sub

Private Sub chkTipoProveedor_Click()
    If chkTipoProveedor.Value = Checked Then
        cboBuscaTipoProveedor.Enabled = True
        cboBuscaTipoProveedor.ListIndex = 0
    Else
        cboBuscaTipoProveedor.Enabled = False
        cboBuscaTipoProveedor.ListIndex = -1
    End If
End Sub

Private Sub CmdBorrar_Click()
    
    If MsgBox("¿Seguro que desea eliminar el Gasto?", vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
        On Error GoTo Seclavose
         lblEstado.Caption = "Eliminando..."
         Screen.MousePointer = vbHourglass
         DBConn.BeginTrans
         
         sql = "DELETE FROM GASTOS_GENERALES"
         sql = sql & " WHERE TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
         sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
         sql = sql & " AND TCO_CODIGO=" & cboComprobante.ItemData(cboComprobante.ListIndex)
         sql = sql & " AND GGR_NROSUC=" & XN(txtNroSucursal)
         sql = sql & " AND GGR_NROCOMP=" & XN(txtNroComprobante)
         DBConn.Execute sql
                                        
         DBConn.CommitTrans
         lblEstado.Caption = ""
         Screen.MousePointer = vbNormal
         CmdNuevo_Click
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
    sql = "SELECT GG.*,P.PROV_RAZSOC,TC.TCO_ABREVIA,TG.TGT_DESCRI"
    sql = sql & " FROM GASTOS_GENERALES GG, TIPO_GASTO TG, TIPO_COMPROBANTE TC, PROVEEDOR P"
    sql = sql & " WHERE"
    sql = sql & " GG.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND GG.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND GG.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND GG.TGT_CODIGO=TG.TGT_CODIGO"
    sql = sql & " AND GG.EST_CODIGO=3"
    If (chkTipoProveedor.Value = Checked And chkProveedor.Value = Checked) Or _
       (chkTipoProveedor.Value = Unchecked And chkProveedor.Value = Checked) Then
        
        If cboBuscaTipoProveedor.ListIndex <> -1 Then
            sql = sql & " AND GG.TPR_CODIGO=" & cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.ListIndex)
        End If
        If txtProveedor.Text <> "" Then
            sql = sql & " AND GG.PROV_CODIGO=" & XN(txtProveedor)
        End If
        
    ElseIf chkTipoProveedor.Value = Checked And chkProveedor.Value = Unchecked Then
        sql = sql & " AND GG.TPR_CODIGO=" & cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.ListIndex)
    End If
    If chkTipoGasto.Value = Checked Then sql = sql & " AND GG.TGT_CODIGO=" & cboBuscaTipoGasto.ItemData(cboBuscaTipoGasto.ListIndex)
    If Not IsNull(FechaDesde) Then sql = sql & " AND GG.GGR_FECHACOMP>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND GG.GGR_FECHACOMP<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY GG.GGR_FECHACOMP DESC"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            GrdModulos.AddItem Rec1!PROV_RAZSOC & Chr(9) & Rec1!TGT_DESCRI & Chr(9) & Rec1!TCO_ABREVIA & Chr(9) & _
                               Rec1!GGR_FECHACOMP & Chr(9) & Rec1!TPR_CODIGO & Chr(9) & Rec1!PROV_CODIGO & Chr(9) & _
                               Rec1!TGT_CODIGO & Chr(9) & Rec1!TCO_CODIGO & Chr(9) & Rec1!GGR_NROSUC & Chr(9) & _
                               Rec1!GGR_NROCOMP & Chr(9) & Rec1!GGR_NETO & Chr(9) & _
                               Rec1!GGR_IVA & Chr(9) & Rec1!GGR_NETO1 & Chr(9) & Rec1!GGR_IVA1 & Chr(9) & _
                               Rec1!GGR_IMPUESTOS & Chr(9) & Rec1!GGR_PERIODO & Chr(9) & _
                               Rec1!GGR_LIBROIVA & Chr(9) & Rec1!GGR_IMP1IVA & Chr(9) & Rec1!GGR_IMP2IVA
                               
            Rec1.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        MsgBox "No se encontraron Datos", vbExclamation, TIT_MSGBOX
        chkTipoProveedor.SetFocus
    End If
    Rec1.Close
End Sub

Private Sub cmdBuscarProveedor_Click()
    frmBuscar.TipoBusqueda = 5
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 1
        txtProveedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 2
        txtDesProv.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 3
        Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboBuscaTipoProveedor)
    Else
        txtProveedor.SetFocus
    End If
End Sub

Private Sub cmdBuscarProveedor1_Click()
    frmBuscar.TipoBusqueda = 5
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 1
        txtCodProveedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 2
        txtProvRazSoc.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 3
        Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboTipoProveedor)
        txtCodProveedor_LostFocus
        txtProvRazSoc.SetFocus
    Else
        txtCodProveedor.SetFocus
    End If
End Sub

Private Sub cmdGrabar_Click()
    
    If ValidarGastosGenerales = False Then Exit Sub
    If MsgBox("¿Confirma Gasto?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    On Error GoTo HayErrorCarga
    
    DBConn.BeginTrans
    sql = "SELECT GGR_NETO FROM GASTOS_GENERALES"
    sql = sql & " WHERE TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
    sql = sql & " AND TCO_CODIGO=" & cboComprobante.ItemData(cboComprobante.ListIndex)
    sql = sql & " AND GGR_NROSUC=" & XN(txtNroSucursal)
    sql = sql & " AND GGR_NROCOMP=" & XN(txtNroComprobante)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = False Then
        If MsgBox("El gasto ya fue ingresado!!!" & Chr(13) & _
                  "Seguro que modificar el Gasto", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
'            'MODIFICO UN GASTO YA REGISTRADO
'            sql = "UPDATE GASTOS_PROVEEDORES"
'            sql = sql & " SET"
'            sql = sql & " GPR_FECHACOMP="
'            sql = sql & " ,GPR_NETO="
'            sql = sql & " ,GPR_IVA="
'            sql = sql & " ,GPR_TOTAL="
'            sql = sql & " ,GPR_PERIODO="
'            sql = sql & " ,TGT_CODIGO="
'            sql = sql & " WHERE TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
'            sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
'            sql = sql & " AND PROV_CODIGO=" & cboComprobante.ItemData(cboComprobante.ListIndex)
'            sql = sql & " AND GPR_NROSUC=" & XS(txtNroSucursal)
'            sql = sql & " AND GPR_NROCOMP=" & XS(txtNroComprobante)
'            DBConn.Execute sql
        End If
        
    Else 'NUEVO GASTO
        
        sql = "INSERT INTO GASTOS_GENERALES"
        sql = sql & " (TPR_CODIGO,PROV_CODIGO,TCO_CODIGO,GGR_NROSUC,GGR_NROCOMP,"
        sql = sql & "GGR_FECHACOMP,GGR_NETO,GGR_IVA,GGR_NETO1,GGR_IVA1,GGR_IMPUESTOS,"
        sql = sql & "GGR_TOTAL,GGR_LIBROIVA,"
        sql = sql & "GGR_PERIODO,TGT_CODIGO,GGR_NROSUCTXT,GGR_NROCOMPTXT,GGR_IMP1IVA,GGR_IMP2IVA,EST_CODIGO,FPG_CODIGO,GGR_SALDO,GGR_OBSER)"
        sql = sql & " VALUES ("
        sql = sql & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex) & ","
        sql = sql & XN(txtCodProveedor) & ","
        sql = sql & cboComprobante.ItemData(cboComprobante.ListIndex) & ","
        sql = sql & XN(txtNroSucursal) & ","
        sql = sql & XN(txtNroComprobante) & ","
        sql = sql & XDQ(FechaComprobante.Value) & ","
        sql = sql & XN(txtNeto) & ","
        sql = sql & XN(txtIva) & ","
        sql = sql & XN(txtNeto1) & ","
        sql = sql & XN(txtIva1) & ","
        sql = sql & XN(txtImpuestos) & ","
        sql = sql & XN(txtTotal) & ","
        If chkCreditoFiscal.Value = Checked Then
            sql = sql & "'S'," 'ENTRA EN EL LIBRO DE IVA COMPRAS
        Else
            sql = sql & "'N'," 'NO ENTRA EN EL LIBRO DE IVA COMPRAS
        End If
        sql = sql & XDQ(Periodo) & ","
        sql = sql & CboGastos.ItemData(CboGastos.ListIndex) & ","
        sql = sql & XS(txtNroSucursal) & ","
        sql = sql & XS(txtNroComprobante) & ","
        sql = sql & XN(txtimp1IVA) & ","
        sql = sql & XN(txtimp2IVA) & ","
        sql = sql & 3 & ","
        sql = sql & cboCondicion.ItemData(cboCondicion.ListIndex) & ","
        If cboCondicion.ItemData(cboCondicion.ListIndex) = 1 Then
            sql = sql & 0 & ","
        Else
            sql = sql & XN(txtTotal) & ","
        End If
        sql = sql & XS(txtObservaciones.Text) & ")"
        DBConn.Execute sql
           
    End If
    rec.Close
        
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.CommitTrans
    CmdNuevo_Click
    Exit Sub
    
HayErrorCarga:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarGastosGenerales() As Boolean
    
    If txtCodProveedor.Text = "" Then
        MsgBox "Debe ingresar un Proveedor", vbExclamation, TIT_MSGBOX
        txtCodProveedor.SetFocus
        ValidarGastosGenerales = False
        Exit Function
    End If
    If CboGastos.ListIndex = -1 Then
        MsgBox "Debe elegir un Tipo de Gasto", vbExclamation, TIT_MSGBOX
        CboGastos.SetFocus
        ValidarGastosGenerales = False
        Exit Function
    End If
    If cboComprobante.ListIndex = -1 Then
        MsgBox "Debe elegir un Tipo de Comprobante", vbExclamation, TIT_MSGBOX
        cboComprobante.SetFocus
        ValidarGastosGenerales = False
        Exit Function
    End If
    If txtNroSucursal.Text = "" Then
        MsgBox "La número de Sucursal es requerida", vbExclamation, TIT_MSGBOX
        txtNroSucursal.SetFocus
        ValidarGastosGenerales = False
        Exit Function
    End If
    If txtNroComprobante.Text = "" Then
        MsgBox "El número de comprobante es requerido", vbExclamation, TIT_MSGBOX
        txtNroComprobante.SetFocus
        ValidarGastosGenerales = False
        Exit Function
    End If
    If IsNull(FechaComprobante.Value) Then
        MsgBox "La Fecha del comprobate es requerida", vbExclamation, TIT_MSGBOX
        FechaComprobante.SetFocus
        ValidarGastosGenerales = False
        Exit Function
    End If
    If txtNeto.Text = "" Then
        MsgBox "El Neto del comprobante es requerido", vbExclamation, TIT_MSGBOX
        txtNeto.SetFocus
        ValidarGastosGenerales = False
        Exit Function
    End If
    If txtIva.Text = "" Then
        MsgBox "El Procentaje del I.V.A. es requerido", vbExclamation, TIT_MSGBOX
        txtIva.SetFocus
        ValidarGastosGenerales = False
        Exit Function
    End If
    If txtTotal.Text = "" Then
        MsgBox "El Total del comprobante es requerido", vbExclamation, TIT_MSGBOX
        txtTotal.SetFocus
        ValidarGastosGenerales = False
        Exit Function
    End If
    If chkCreditoFiscal.Value = Checked Then
        If IsNull(Periodo.Value) Then
            MsgBox "El Periodo es requerido (Libro I.V.A. Compras)", vbExclamation, TIT_MSGBOX
            Periodo.SetFocus
            ValidarGastosGenerales = False
            Exit Function
        End If
    End If
    
    ValidarGastosGenerales = True
End Function

Private Sub CmdNuevo_Click()
    LimpiarBusqueda
    limpiar_datos
End Sub
Private Sub limpiar_datos()

    Call CambioEstado(True)
    cboTipoProveedor.ListIndex = 0
    txtCodProveedor.Text = ""
    CboGastos.ListIndex = 0
    cboComprobante.ListIndex = 0
    txtNroSucursal.Text = ""
    txtNroComprobante.Text = ""
    FechaComprobante.Value = Null
    txtNeto.Text = "0,00"
    txtIva.Text = ""
    txtNeto1.Text = "0,00"
    txtIva1.Text = ""
    txtSubtotal.Text = "0,00"
    txtSubTotal1.Text = "0,00"
    txtImpuestos.Text = "0,00"
    txtTotal.Text = "0,00"
    Periodo.Value = Null
    chkCreditoFiscal.Value = Unchecked
    CmdBorrar.Enabled = False
    cmdGrabar.Enabled = True
    cboTipoProveedor.SetFocus
    tabDatos.Tab = 0
    txtimp1IVA.Text = ""
    txtimp2IVA.Text = ""
    cboCondicion.ListIndex = 0
    txtObservaciones.Text = ""
    lblPeriodo1.Caption = ""
End Sub


Private Sub cmdNuevoGasto_Click()
    ABMTipoGasto.Show vbModal
    CboGastos.Clear
    'CARGO COMBO GASTOS
    LlenarComboGastos
    CboGastos.SetFocus
End Sub

Private Sub cmdNuevoRubro_Click()
    ABMFormaPago.Show vbModal
End Sub

Private Sub CmdSalir_Click()
    Set frmCargaGastosGenerales = Nothing
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub LimpiarBusqueda()
    chkTipoProveedor.Value = Unchecked
    chkProveedor.Value = Unchecked
    chkTipoGasto.Value = Unchecked
    chkFecha.Value = Unchecked
    GrdModulos.Rows = 1
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    
    'CARGO COMBO TIPO PROVEEDOR
    LlenarComboTipoProv
    'CARGO COMBO COMPROBANTES
    LlenarComboComprobante
    'CARGO COMBO GASTOS
    LlenarComboGastos
   
    LlenarComboFormaPago
     'CONFIGURO GRILLA BUSQUEDA
    GrdModulos.FormatString = "Proveedor|Gasto|Comprobante|^Fecha|TIPO PROVEEDOR|" _
                            & "COD PROVEEDOR|COD TIPO GASTO|COD TIP COMPROBANTE|" _
                            & "NRO SUCURSAL|NRO COMPROBANTE|NETO|IVA|NETO1|IVA1|IMPUESTOS|PERIODO|ENTRA LIBRO IVA"
                            
    GrdModulos.ColWidth(0) = 3200 'Proveedor
    GrdModulos.ColWidth(1) = 3000 'Gasto
    GrdModulos.ColWidth(2) = 1100 'Comprobante
    GrdModulos.ColWidth(3) = 1000 'Fecha
    GrdModulos.ColWidth(4) = 0    'TIPO PROVEEDOR
    GrdModulos.ColWidth(5) = 0    'COD PROVEEDOR
    GrdModulos.ColWidth(6) = 0    'COD TIPO GASTO
    GrdModulos.ColWidth(7) = 0    'COD TIP COMPROBANTE
    GrdModulos.ColWidth(8) = 0    'NRO SUCURSAL
    GrdModulos.ColWidth(9) = 0    'NRO COMPROBANTE
    GrdModulos.ColWidth(10) = 0   'NETO
    GrdModulos.ColWidth(11) = 0   'IVA
    GrdModulos.ColWidth(12) = 0   'NETO1
    GrdModulos.ColWidth(13) = 0   'IVA1
    GrdModulos.ColWidth(14) = 0   'IMPUESTOS
    GrdModulos.ColWidth(15) = 0   'PERIODO
    GrdModulos.ColWidth(16) = 0   'ENTRA LIBRO IVA
    GrdModulos.ColWidth(17) = 0   'IMPORTE IVA 1
    GrdModulos.ColWidth(18) = 0   'IMPORTE IVA 2
    GrdModulos.Cols = 19
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    cmdGrabar.Enabled = True
    CmdBorrar.Enabled = False
    Periodo.Enabled = False
    lblEstado.Caption = ""
    txtNeto.Text = "0,00"
    txtNeto1.Text = "0,00"
    txtSubtotal.Text = "0,00"
    txtSubTotal1.Text = "0,00"
    txtImpuestos.Text = "0,00"
    txtTotal.Text = "0,00"
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
Private Sub LlenarComboGastos()
    sql = "SELECT * FROM TIPO_GASTO ORDER BY TGT_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            CboGastos.AddItem rec!TGT_DESCRI
            CboGastos.ItemData(CboGastos.NewIndex) = rec!TGT_CODIGO
            cboBuscaTipoGasto.AddItem rec!TGT_DESCRI
            cboBuscaTipoGasto.ItemData(cboBuscaTipoGasto.NewIndex) = rec!TGT_CODIGO
            rec.MoveNext
        Loop
        CboGastos.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboComprobante()
    sql = "SELECT TCO_CODIGO,TCO_DESCRI"
    sql = sql & " FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_CODIGO NOT IN (14,15,16)"
    sql = sql & " ORDER BY TCO_DESCRI"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboComprobante.AddItem rec!TCO_DESCRI
            cboComprobante.ItemData(cboComprobante.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboComprobante.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboTipoProv()
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboTipoProveedor.AddItem "TODOS"
        Do While rec.EOF = False
            cboTipoProveedor.AddItem rec!TPR_DESCRI
            cboTipoProveedor.ItemData(cboTipoProveedor.NewIndex) = rec!TPR_CODIGO
            cboBuscaTipoProveedor.AddItem rec!TPR_DESCRI
            cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.NewIndex) = rec!TPR_CODIGO
            rec.MoveNext
        Loop
        cboTipoProveedor.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        limpiar_datos
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 4)), cboTipoProveedor)
        txtCodProveedor.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
        txtCodProveedor_LostFocus
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6)), CboGastos)
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 7)), cboComprobante)
        txtNroSucursal.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
        txtNroSucursal_LostFocus
        txtNroComprobante.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 9)
        txtNroComprobante_LostFocus
        FechaComprobante.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 3)
        txtNeto.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 10))
        txtIva.Text = Format(GrdModulos.TextMatrix(GrdModulos.RowSel, 11), "0.00")
        txtIva_LostFocus
        txtNeto1.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 12))
        txtIva1.Text = Format(GrdModulos.TextMatrix(GrdModulos.RowSel, 13), "0.00")
        txtIva1_LostFocus
        txtImpuestos.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 14))
        txtImpuestos_LostFocus
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 16) = "S" Then
            chkCreditoFiscal.Value = Checked
            Periodo.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 15)
            Periodo_LostFocus
        Else
            chkCreditoFiscal.Value = Unchecked
        End If
        FrameProveedor.Enabled = False
        cboComprobante.Enabled = False
        'pongo enable falso (los campos clave) ya que consulto
        Call CambioEstado(False)
        
        BuscoFormaPago
        
        CboGastos.SetFocus
        CmdBorrar.Enabled = True
        cmdGrabar.Enabled = False
        tabDatos.Tab = 0
    
    End If
End Sub
Private Function BuscoFormaPago()
'busco forma de pago y observaciones
    sql = "SELECT FPG_CODIGO, GGR_OBSER FROM GASTOS_GENERALES"
    sql = sql & " WHERE TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
    sql = sql & " AND TCO_CODIGO=" & cboComprobante.ItemData(cboComprobante.ListIndex)
    sql = sql & " AND GGR_NROSUC=" & XN(txtNroSucursal)
    sql = sql & " AND GGR_NROCOMP=" & XN(txtNroComprobante)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        If Not IsNull(rec!FPG_CODIGO) Then
            Call BuscaCodigoProxItemData(rec!FPG_CODIGO, cboCondicion)
            txtObservaciones.Text = IIf(IsNull(rec!GGR_OBSER), "", rec!GGR_OBSER)
        End If
    End If
    rec.Close

End Function
Private Sub CambioEstado(Estado As Boolean)
    FrameProveedor.Enabled = Estado
    cboComprobante.Enabled = Estado
    txtNroSucursal.Enabled = Estado
    txtNroComprobante.Enabled = Estado
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GrdModulos_DblClick
    End If
End Sub

Private Sub Periodo_Change()
    If IsNull(Periodo.Value) Then
        lblPeriodo1.Caption = ""
    End If
End Sub

Private Sub Periodo_LostFocus()
    If Trim(Periodo.Value) <> "" Then
        lblPeriodo1.Caption = UCase(Format(Periodo.Value, "mmmm/yyyy"))
    Else
        lblPeriodo1.Caption = ""
    End If
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    cboBuscaTipoProveedor.ListIndex = -1
    cboBuscaTipoGasto.ListIndex = -1
    If tabDatos.Tab = 1 Then
      cboBuscaTipoProveedor.Enabled = False
      txtProveedor.Enabled = False
      FechaDesde.Enabled = False
      FechaHasta.Enabled = False
      cboBuscaTipoGasto.Enabled = False
      cmdGrabar.Enabled = False
      CmdBorrar.Enabled = False
      cmdBuscarProveedor.Enabled = False
      If Me.Visible = True Then chkTipoProveedor.SetFocus
    Else
        If Me.Visible = True Then
          If FrameProveedor.Enabled = True Then
              cboTipoProveedor.SetFocus
          Else
              CboGastos.SetFocus
          End If
        End If
    End If
End Sub

Private Sub txtCodProveedor_Change()
    If txtCodProveedor.Text = "" Then
        txtProvRazSoc.Text = ""
        txtCliLocalidad.Text = ""
        txtDomici.Text = ""
    End If
End Sub

Private Sub txtCodProveedor_GotFocus()
    SelecTexto txtCodProveedor
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodProveedor_LostFocus()
    If txtCodProveedor.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        sql = "SELECT PRO.TPR_CODIGO,PRO.PROV_CODIGO, PRO.PROV_RAZSOC,"
        sql = sql & " PRO.PROV_DOMICI, L.LOC_DESCRI"
        sql = sql & " FROM PROVEEDOR PRO,LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE"
        sql = sql & " PRO.PROV_CODIGO=" & XN(txtCodProveedor.Text)
        'sql = sql & " PRO.PROV_RAZSOC LIKE '" & Pro & "%'"
        If cboTipoProveedor.List(cboTipoProveedor.ListIndex) <> "TODOS" Then
            sql = sql & " AND PRO.TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
        End If
        sql = sql & " AND PRO.LOC_CODIGO=L.LOC_CODIGO"
        sql = sql & " AND PRO.PRO_CODIGO=P.PRO_CODIGO"
        sql = sql & " AND P.PRO_CODIGO=L.PRO_CODIGO"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtProvRazSoc.Text = Rec1!PROV_RAZSOC
            txtCliLocalidad.Text = Rec1!LOC_DESCRI
            txtDomici.Text = Rec1!PROV_DOMICI
            Call BuscaCodigoProxItemData(CInt(Rec1!TPR_CODIGO), cboTipoProveedor)
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtImpuestos_GotFocus()
    SelecTexto txtImpuestos
End Sub

Private Sub txtImpuestos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImpuestos, KeyAscii)
End Sub

Private Sub txtImpuestos_LostFocus()
    If txtImpuestos.Text <> "" Then
        txtImpuestos.Text = Valido_Importe(txtImpuestos.Text)
        txtTotal.Text = CDbl(txtImpuestos.Text) + CDbl(txtSubtotal.Text) + CDbl(txtSubTotal1.Text)
        txtTotal.Text = Valido_Importe(txtTotal)
    Else
        txtImpuestos.Text = "0,00"
    End If
End Sub

Private Sub txtIva_GotFocus()
    SelecTexto txtIva
End Sub

Private Sub txtIva_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtIva, KeyAscii)
End Sub

Private Sub txtIva_LostFocus()
    If txtIva.Text <> "" Then
        If ValidarPorcentaje(txtIva) = False Then
            txtIva.SetFocus
            Exit Sub
        End If
        txtimp1IVA.Text = Valido_Importe((CDbl(txtNeto.Text) * CDbl(txtIva.Text)) / 100)
        txtSubtotal.Text = CDbl(txtNeto.Text) + ((CDbl(txtNeto.Text) * CDbl(txtIva.Text)) / 100)
        txtSubtotal.Text = Valido_Importe(txtSubtotal)
        txtTotal.Text = CDbl(txtSubTotal1.Text) + CDbl(txtSubtotal.Text) + CDbl(txtImpuestos.Text)
        txtTotal.Text = Valido_Importe(txtTotal)
    End If
End Sub

Private Sub txtIva1_GotFocus()
     SelecTexto txtIva
End Sub

Private Sub txtIva1_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroDecimal(txtIva1, KeyAscii)
End Sub

Private Sub txtIva1_LostFocus()
    If txtIva1.Text <> "" Then
        If ValidarPorcentaje(txtIva1) = False Then
            txtIva1.SetFocus
            Exit Sub
        End If
        txtimp2IVA.Text = Valido_Importe((CDbl(txtNeto1.Text) * CDbl(txtIva1.Text)) / 100)
        txtSubTotal1.Text = CDbl(txtNeto1.Text) + ((CDbl(txtNeto1.Text) * CDbl(txtIva1.Text)) / 100)
        txtSubTotal1.Text = Valido_Importe(txtSubTotal1)
        txtTotal.Text = CDbl(txtSubTotal1.Text) + CDbl(txtSubtotal.Text) + CDbl(txtImpuestos.Text)
        txtTotal.Text = Valido_Importe(txtTotal)
    End If
End Sub


Private Sub txtNeto_GotFocus()
    SelecTexto txtNeto
End Sub

Private Sub txtNeto_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtNeto, KeyAscii)
End Sub

Private Sub txtNeto_LostFocus()
    If txtNeto.Text <> "" Then
        txtNeto.Text = Valido_Importe(txtNeto)
    Else
        txtNeto.Text = "0,00"
    End If
End Sub

Private Sub txtNeto1_GotFocus()
    SelecTexto txtNeto1
End Sub

Private Sub txtNeto1_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtNeto1, KeyAscii)
End Sub

Private Sub txtNeto1_LostFocus()
    If txtNeto1.Text <> "" Then
        txtNeto1.Text = Valido_Importe(txtNeto1.Text)
    Else
        txtNeto1.Text = "0,00"
    End If
End Sub

Private Sub txtNroComprobante_GotFocus()
    SelecTexto txtNroComprobante
End Sub

Private Sub txtNroComprobante_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroComprobante_LostFocus()
    If txtNroComprobante.Text <> "" Then
        txtNroComprobante.Text = Format(txtNroComprobante.Text, "00000000")
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
        txtNroSucursal.Text = "1"
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    Else
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    End If
End Sub

Private Sub txtObservaciones_GotFocus()
    SelecTexto txtObservaciones
End Sub

Private Sub txtProveedor_Change()
    If txtProveedor.Text = "" Then
        txtDesProv.Text = ""
    End If
End Sub

Private Sub txtProveedor_GotFocus()
    SelecTexto txtProveedor
End Sub

Private Sub txtProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtProveedor_LostFocus()
    If txtProveedor.Text <> "" Then
        sql = "SELECT TPR_CODIGO,PROV_CODIGO,PROV_RAZSOC,"
        
        
        
        sql = sql & " FROM PROVEEDOR"
        sql = sql & " WHERE"
        sql = sql & " PROV_CODIGO=" & XN(txtProveedor)
        
        Rec1.Open BuscoProveedor(txtProveedor.Text), DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtDesProv.Text = Rec1!PROV_RAZSOC
            Call BuscaCodigoProxItemData(CInt(Rec1!TPR_CODIGO), cboBuscaTipoProveedor)
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtDesProv.Text = ""
            cboTipoProveedor.ListIndex = 0
            txtProveedor.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtProvRazSoc_Change()
    If txtProvRazSoc.Text = "" Then
        txtCodProveedor.Text = ""
        txtCliLocalidad.Text = ""
        txtDomici.Text = ""
    End If
End Sub

Private Sub txtProvRazSoc_GotFocus()
    SelecTexto txtProvRazSoc
End Sub

Private Sub txtProvRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtProvRazSoc_LostFocus()
    If txtCodProveedor.Text = "" And txtProvRazSoc.Text <> "" Then
        rec.Open BuscoProveedor(txtProvRazSoc), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 5
                frmBuscar.TxtDescriB.Text = txtProvRazSoc.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 1
                    txtCodProveedor.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 2
                    txtProvRazSoc.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 3
                    Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboTipoProveedor)
                    txtCodProveedor_LostFocus
                Else
                    txtCodProveedor.SetFocus
                End If
            Else
                txtCodProveedor.Text = rec!PROV_CODIGO
                txtProvRazSoc.Text = rec!PROV_RAZSOC
                txtCodProveedor_LostFocus
            End If
        Else
            MsgBox "No se encontro el Proveedor", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        rec.Close
    ElseIf txtCodProveedor.Text = "" And txtProvRazSoc.Text = "" Then
        MsgBox "Debe elegir un Proveedor", vbExclamation, TIT_MSGBOX
        txtCodProveedor.SetFocus
    End If
End Sub

Private Function BuscoProveedor(Pro As String) As String
    sql = "SELECT PRO.TPR_CODIGO,PRO.PROV_CODIGO, PRO.PROV_RAZSOC,"
    sql = sql & " PRO.PROV_DOMICI, L.LOC_DESCRI"
    sql = sql & " FROM PROVEEDOR PRO,LOCALIDAD L"
    sql = sql & " WHERE"
    If txtProveedor.Text <> "" Then
        sql = sql & " PRO.PROV_CODIGO=" & XN(Pro)
    Else
        sql = sql & " PRO.PROV_RAZSOC LIKE '" & Pro & "%'"
    End If
    If cboTipoProveedor.List(cboTipoProveedor.ListIndex) <> "TODOS" Then
        sql = sql & " AND PRO.TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    End If
    sql = sql & " AND PRO.LOC_CODIGO=L.LOC_CODIGO"

    BuscoProveedor = sql
End Function

Private Sub txtTotal_GotFocus()
    SelecTexto txtTotal
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtTotal, KeyAscii)
End Sub

Private Sub txtTotal_LostFocus()
    If txtTotal.Text <> "" Then
        txtTotal.Text = Valido_Importe(txtTotal)
    Else
        txtTotal.Text = "0,00"
    End If
End Sub
