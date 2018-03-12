VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNotaDebitoProveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Débito de Proveedores..."
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      Height          =   450
      Left            =   6960
      TabIndex        =   14
      Top             =   6180
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6060
      Left            =   45
      TabIndex        =   26
      Top             =   60
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   10689
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
      TabPicture(0)   =   "frmNotaDebitoProveedores.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameProveedor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmNotaDebitoProveedores.frx":001C
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
         Height          =   1590
         Left            =   -74835
         TabIndex        =   43
         Top             =   375
         Width           =   8355
         Begin VB.ComboBox cboBuscaTipoProveedor 
            Height          =   315
            Left            =   2385
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   330
            Width           =   3750
         End
         Begin VB.CheckBox chkTipoProveedor 
            Caption         =   "Tipo Prov"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   450
            Width           =   1050
         End
         Begin VB.CheckBox chkProveedor 
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   705
            Width           =   1125
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   960
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
            TabIndex        =   45
            Tag             =   "Descripción"
            Top             =   675
            Width           =   4440
         End
         Begin VB.TextBox txtProveedor 
            Height          =   300
            Left            =   2385
            MaxLength       =   40
            TabIndex        =   20
            Top             =   675
            Width           =   975
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   450
            Left            =   6780
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmNotaDebitoProveedores.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Buscar "
            Top             =   1065
            UseMaskColor    =   -1  'True
            Width           =   1455
         End
         Begin VB.CommandButton cmdBuscarProveedor 
            Height          =   300
            Left            =   3390
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDebitoProveedores.frx":27DA
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Buscar Proveedor"
            Top             =   675
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   2400
            TabIndex        =   21
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   17235969
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   5010
            TabIndex        =   22
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   17235969
            CurrentDate     =   41098
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Prov.:"
            Height          =   195
            Left            =   1545
            TabIndex        =   49
            Top             =   375
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
            TabIndex        =   48
            Top             =   705
            Width           =   780
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1320
            TabIndex        =   47
            Top             =   1140
            Width           =   1005
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   3960
            TabIndex        =   46
            Top             =   1140
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
         TabIndex        =   36
         Top             =   585
         Width           =   8355
         Begin VB.CommandButton cmdBuscarProveedor1 
            Height          =   300
            Left            =   2295
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDebitoProveedores.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   51
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            Width           =   5400
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
            TabIndex        =   42
            Top             =   1455
            Width           =   420
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Loc.:"
            Height          =   180
            Left            =   825
            TabIndex        =   41
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
            TabIndex        =   40
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Prov.:"
            Height          =   195
            Left            =   405
            TabIndex        =   39
            Top             =   435
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Nota de Débito..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3330
         Left            =   165
         TabIndex        =   27
         Top             =   2610
         Width           =   8355
         Begin VB.TextBox txtConcepto 
            Height          =   300
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   11
            Top             =   2850
            Width           =   6825
         End
         Begin VB.ComboBox cboComprobante 
            Height          =   315
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   465
            Width           =   2400
         End
         Begin VB.TextBox txtNroSucursal 
            Height          =   285
            Left            =   1275
            MaxLength       =   4
            TabIndex        =   4
            Top             =   825
            Width           =   480
         End
         Begin VB.TextBox txtNroComprobante 
            Height          =   285
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   5
            Top             =   825
            Width           =   960
         End
         Begin VB.TextBox txtTotal 
            Height          =   300
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   9
            Top             =   2175
            Width           =   1140
         End
         Begin VB.TextBox txtIva 
            Height          =   300
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   8
            Top             =   1830
            Width           =   660
         End
         Begin VB.TextBox txtNeto 
            Height          =   300
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   7
            Top             =   1485
            Width           =   1140
         End
         Begin MSComCtl2.DTPicker FechaComprobante 
            Height          =   315
            Left            =   1275
            TabIndex        =   6
            Top             =   1140
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   17235969
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker Periodo 
            Height          =   315
            Left            =   1275
            TabIndex        =   10
            Top             =   2505
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   17235969
            CurrentDate     =   41098
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Concepto:"
            Height          =   195
            Left            =   465
            TabIndex        =   50
            Top             =   2895
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante:"
            Height          =   195
            Left            =   210
            TabIndex        =   35
            Top             =   510
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   600
            TabIndex        =   34
            Top             =   855
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   705
            TabIndex        =   33
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Neto:"
            Height          =   195
            Left            =   810
            TabIndex        =   32
            Top             =   1545
            Width           =   390
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Iva:"
            Height          =   195
            Left            =   930
            TabIndex        =   31
            Top             =   1875
            Width           =   270
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   795
            TabIndex        =   30
            Top             =   2220
            Width           =   405
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Periodo:"
            Height          =   195
            Left            =   615
            TabIndex        =   29
            Top             =   2550
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
            Left            =   2790
            TabIndex        =   28
            Top             =   2520
            Width           =   1785
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   3825
         Left            =   -74865
         TabIndex        =   24
         Top             =   2025
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   6747
         _Version        =   393216
         Cols            =   14
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
      TabIndex        =   13
      Top             =   6180
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   5190
      TabIndex        =   12
      Top             =   6180
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   7845
      TabIndex        =   15
      Top             =   6180
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
      Left            =   120
      TabIndex        =   25
      Top             =   6225
      Width           =   750
   End
End
Attribute VB_Name = "frmNotaDebitoProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
        If chkTipoProveedor.Value = Unchecked And cboBuscaTipoProveedor.ListIndex <> -1 Then
            cboBuscaTipoProveedor.ListIndex = -1
        End If
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
    
    If MsgBox("¿Seguro que desea eliminar la Nota de Débito?", vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
        On Error GoTo Seclavose
         lblEstado.Caption = "Eliminando..."
         Screen.MousePointer = vbHourglass
         DBConn.BeginTrans
         
         sql = "DELETE FROM NOTA_DEBITO_PROVEEDOR"
         sql = sql & " WHERE TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
         sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
         sql = sql & " AND TCO_CODIGO=" & cboComprobante.ItemData(cboComprobante.ListIndex)
         sql = sql & " AND DPR_NROSUC=" & XN(txtNroSucursal)
         sql = sql & " AND DPR_NUMERO=" & XN(txtNroComprobante)
         DBConn.Execute sql
         
         'BORRO DE LA CUENTA CORRIENTE DEL PROVEEDOR
         DBConn.Execute QuitoCtaCteProveedores(CStr(cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)), txtCodProveedor, _
                                          CStr(cboComprobante.ItemData(cboComprobante.ListIndex)), txtNroSucursal, txtNroComprobante)
                                          
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
    MsgBox Err.Description, vbCritical, TIT_MSGBOX, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
     GrdModulos.Rows = 1
    sql = "SELECT DP.*,P.PROV_RAZSOC,TC.TCO_ABREVIA"
    sql = sql & " FROM NOTA_DEBITO_PROVEEDOR DP, TIPO_COMPROBANTE TC, PROVEEDOR P"
    sql = sql & " WHERE"
    sql = sql & " DP.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND DP.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND DP.TCO_CODIGO=TC.TCO_CODIGO"
    If (chkTipoProveedor.Value = Checked And chkProveedor.Value = Checked) Or _
       (chkTipoProveedor.Value = Unchecked And chkProveedor.Value = Checked) Then
        
        If cboBuscaTipoProveedor.ListIndex <> -1 Then
            sql = sql & " AND DP.TPR_CODIGO=" & cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.ListIndex)
        End If
        If txtProveedor.Text <> "" Then
            sql = sql & " AND DP.PROV_CODIGO=" & XN(txtProveedor)
        End If
        
    ElseIf chkTipoProveedor.Value = Checked And chkProveedor.Value = Unchecked Then
        sql = sql & " AND DP.TPR_CODIGO=" & cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.ListIndex)
    End If
    If Not IsNull(FechaDesde) Then sql = sql & " AND DP.DPR_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND DP.DPR_FECHA<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY DP.DPR_FECHA,DP.PROV_CODIGO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            GrdModulos.AddItem Rec1!PROV_RAZSOC & Chr(9) & Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!DPR_NROSUC, "0000") & "-" & Format(Rec1!DPR_NUMERO, "00000000") & Chr(9) & _
                               Rec1!DPR_FECHA & Chr(9) & Rec1!TPR_CODIGO & Chr(9) & Rec1!PROV_CODIGO & Chr(9) & _
                               Rec1!DPR_CONCEPTO & Chr(9) & Rec1!TCO_CODIGO & Chr(9) & Rec1!DPR_NROSUC & Chr(9) & _
                               Rec1!DPR_NUMERO & Chr(9) & Rec1!DPR_NETO & Chr(9) & _
                               Rec1!DPR_IVA & Chr(9) & Rec1!DPR_TOTAL & Chr(9) & Rec1!DPR_PERIODO
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
    
    If ValidarGastosProveedor = False Then Exit Sub
    If MsgBox("¿Confirma Nota de Débito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    On Error GoTo HayErrorCarga
    
    DBConn.BeginTrans
    sql = "SELECT DPR_NETO FROM NOTA_DEBITO_PROVEEDOR"
    sql = sql & " WHERE TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
    sql = sql & " AND TCO_CODIGO=" & cboComprobante.ItemData(cboComprobante.ListIndex)
    sql = sql & " AND DPR_NROSUC=" & XN(txtNroSucursal)
    sql = sql & " AND DPR_NUMERO=" & XN(txtNroComprobante)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = False Then
        If MsgBox("La Nota de Débito ya fue ingresada!!!" & Chr(13) & _
                  "Seguro que modifica la Nota de Débito", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
             'MODIFICO UN GASTO YA REGISTRADO
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
        
        sql = "INSERT INTO NOTA_DEBITO_PROVEEDOR"
        sql = sql & " (TPR_CODIGO,PROV_CODIGO,TCO_CODIGO,DPR_NROSUC,DPR_NUMERO,"
        sql = sql & "DPR_FECHA,DPR_NETO,DPR_IVA,DPR_TOTAL,DPR_SALDO,"
        sql = sql & "DPR_PERIODO,DPR_CONCEPTO,DPR_NROSUCTXT,DPR_NUMEROTXT)"
        sql = sql & " VALUES ("
        sql = sql & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex) & ","
        sql = sql & XN(txtCodProveedor) & ","
        sql = sql & cboComprobante.ItemData(cboComprobante.ListIndex) & ","
        sql = sql & XN(txtNroSucursal) & ","
        sql = sql & XN(txtNroComprobante) & ","
        sql = sql & XDQ(FechaComprobante) & ","
        sql = sql & XN(txtNeto) & ","
        sql = sql & XN(txtIva) & ","
        sql = sql & XN(txtTotal) & ","
        sql = sql & XN(txtTotal) & "," 'SALDO COMRPOBANTE
        sql = sql & XDQ(Periodo) & ","
        sql = sql & XS(txtConcepto) & ","
        sql = sql & XS(txtNroSucursal) & ","
        sql = sql & XS(txtNroComprobante) & ")"
        DBConn.Execute sql
           
    End If
    rec.Close
    
    'ACTUALIZO CUNETA CORRIENTE DEL PROVEEDOR
    DBConn.Execute AgregoCtaCteProveedores(CStr(cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)), txtCodProveedor, CStr(cboComprobante.ItemData(cboComprobante.ListIndex)) _
                                        , txtNroSucursal, txtNroComprobante, FechaComprobante, txtTotal, "D", CStr(Date))
    
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.CommitTrans
    CmdNuevo_Click
    Exit Sub
    
HayErrorCarga:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarGastosProveedor() As Boolean
    
    If txtCodProveedor.Text = "" Then
        MsgBox "Debe ingresar un Proveedor", vbExclamation, TIT_MSGBOX
        txtCodProveedor.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If cboComprobante.ListIndex = -1 Then
        MsgBox "Debe elegir un Tipo de Comprobante", vbExclamation, TIT_MSGBOX
        cboComprobante.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If txtNroSucursal.Text = "" Then
        MsgBox "La número de Sucursal es requerida", vbExclamation, TIT_MSGBOX
        txtNroSucursal.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If txtNroComprobante.Text = "" Then
        MsgBox "El número de comprobante es requerido", vbExclamation, TIT_MSGBOX
        txtNroComprobante.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If IsNull(FechaComprobante.Value) Then
        MsgBox "La Fecha del comprobate es requerida", vbExclamation, TIT_MSGBOX
        FechaComprobante.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If txtNeto.Text = "" Then
        MsgBox "El Neto del comprobante es requerido", vbExclamation, TIT_MSGBOX
        txtNeto.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If txtIva.Text = "" Then
        MsgBox "El Procentaje del I.V.A. es requerido", vbExclamation, TIT_MSGBOX
        txtIva.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If txtTotal.Text = "" Then
        MsgBox "El Total del comprobante es requerido", vbExclamation, TIT_MSGBOX
        txtTotal.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If IsNull(Periodo.Value) Then
        MsgBox "El Periodo es requerido (Libro I.V.A. Compras)", vbExclamation, TIT_MSGBOX
        Periodo.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If txtConcepto.Text = "" Then
        MsgBox "El Concepto de la Nota de Crédito es requerido", vbExclamation, TIT_MSGBOX
        txtConcepto.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    ValidarGastosProveedor = True
End Function

Private Sub CmdNuevo_Click()
    LimpiarBusqueda
    Call CambioEstado(True)
    cboTipoProveedor.ListIndex = 0
    txtCodProveedor.Text = ""
    cboComprobante.ListIndex = 0
    txtNroSucursal.Text = ""
    txtNroComprobante.Text = ""
    FechaComprobante.Value = Null
    txtNeto.Text = ""
    txtIva.Text = ""
    txtTotal.Text = ""
    Periodo.Value = Null
    txtConcepto.Text = ""
    cmdBorrar.Enabled = False
    cmdGrabar.Enabled = True
    cboTipoProveedor.SetFocus
    tabDatos.Tab = 0
End Sub

Private Sub CmdSalir_Click()
    Set frmNotaDebitoProveedores = Nothing
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
    'CONFIGURO GRILLA BUSQUEDA
    GrdModulos.FormatString = "Proveedor|Comprobante|^Número|^Fecha|TIPO PROVEEDOR|" _
                            & "COD PROVEEDOR|Concepto|COD TIP COMPROBANTE|" _
                            & "NRO SUCURSAL|NRO COMPROBANTE|NETO|IVA|TOTAL|PERIODO"
                            
    GrdModulos.ColWidth(0) = 4000 'Proveedor
    GrdModulos.ColWidth(1) = 1200 'Comprobnate
    GrdModulos.ColWidth(2) = 1500 'Numero
    GrdModulos.ColWidth(3) = 1200 'Fecha
    GrdModulos.ColWidth(4) = 0    'TIPO PROVEEDOR
    GrdModulos.ColWidth(5) = 0    'COD PROVEEDOR
    GrdModulos.ColWidth(6) = 0    'CONCEPTO
    GrdModulos.ColWidth(7) = 0    'COD TIP COMPROBANTE
    GrdModulos.ColWidth(8) = 0    'NRO SUCURSAL
    GrdModulos.ColWidth(9) = 0    'NRO COMPROBANTE
    GrdModulos.ColWidth(10) = 0    'NETO
    GrdModulos.ColWidth(11) = 0   'IVA
    GrdModulos.ColWidth(12) = 0   'TOTAL
    GrdModulos.ColWidth(13) = 0   'PERIODO

    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    cmdGrabar.Enabled = True
    cmdBorrar.Enabled = False
    lblEstado.Caption = ""
End Sub

Private Sub LlenarComboComprobante()
    sql = "SELECT TCO_CODIGO,TCO_DESCRI"
    sql = sql & " FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'NOTA DE DEB%'"
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
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 4)), cboTipoProveedor)
        txtCodProveedor.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
        txtCodProveedor_LostFocus
        txtConcepto.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 7)), cboComprobante)
        txtNroSucursal.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
        txtNroSucursal_LostFocus
        txtNroComprobante.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 9)
        txtNroComprobante_LostFocus
        FechaComprobante.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 3)
        txtNeto.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 10))
        txtIva.Text = Format(GrdModulos.TextMatrix(GrdModulos.RowSel, 11), "0.00")
        txtTotal.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 12))
        Periodo.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 13)
        Periodo_LostFocus
        FrameProveedor.Enabled = False
        cboComprobante.Enabled = False
        'pongo enable falso (los campos clave) ya que consulto
        Call CambioEstado(False)
        'CboGastos.SetFocus
        cmdBorrar.Enabled = True
        cmdGrabar.Enabled = False
        tabDatos.Tab = 0
    End If
End Sub

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
    If tabDatos.Tab = 1 Then
      cboBuscaTipoProveedor.Enabled = False
      txtProveedor.Enabled = False
      FechaDesde.Enabled = False
      FechaHasta.Enabled = False
      cmdGrabar.Enabled = False
      cmdBorrar.Enabled = False
      cmdBuscarProveedor.Enabled = False
      If Me.Visible = True Then chkTipoProveedor.SetFocus
    Else
        If Me.Visible = True Then
          If FrameProveedor.Enabled = True Then
              cboTipoProveedor.SetFocus
          Else
              FechaComprobante.SetFocus
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
        Rec1.Open BuscoProveedor(txtCodProveedor), DBConn, adOpenStatic, adLockOptimistic
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

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
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
        txtTotal.Text = CDbl(txtNeto.Text) + ((CDbl(txtNeto.Text) * CDbl(txtIva.Text)) / 100)
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
    Set Rec1 = New ADODB.Recordset
    If txtProveedor.Text <> "" Then
        sql = "SELECT TPR_CODIGO,PROV_CODIGO,PROV_RAZSOC,"
        sql = sql & " FROM PROVEEDOR"
        sql = sql & " WHERE"
        sql = sql & " PROV_CODIGO=" & XN(txtProveedor)
        
        Rec1.Open BuscoProveedor(txtProveedor), DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtDesProv.Text = Rec1!PROV_RAZSOC
            Call BuscaCodigoProxItemData(CInt(Rec1!TPR_CODIGO), cboBuscaTipoProveedor)
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
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
    sql = sql & " FROM PROVEEDOR PRO,LOCALIDAD L,PROVINCIA P"
    sql = sql & " WHERE"
    If txtCodProveedor.Text <> "" Or txtProveedor.Text <> "" Then
        sql = sql & " PRO.PROV_CODIGO=" & XN(Pro)
    Else
        sql = sql & " PRO.PROV_RAZSOC LIKE '" & Pro & "%'"
    End If
    If cboTipoProveedor.List(cboTipoProveedor.ListIndex) <> "TODOS" Then
        sql = sql & " AND PRO.TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    End If
    sql = sql & " AND PRO.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND PRO.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"

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
