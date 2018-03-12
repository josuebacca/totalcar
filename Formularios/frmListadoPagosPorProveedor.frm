VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListadoPagosPorProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Recibo de Proveedores"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   720
      Left            =   4365
      Picture         =   "frmListadoPagosPorProveedor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2430
      Width           =   840
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   720
      Left            =   6090
      Picture         =   "frmListadoPagosPorProveedor.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2430
      Width           =   825
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoPagosPorProveedor.frx":0BD4
      Height          =   720
      Left            =   5220
      Picture         =   "frmListadoPagosPorProveedor.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2430
      Width           =   855
   End
   Begin VB.Frame Frame1 
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
      Height          =   1635
      Left            =   60
      TabIndex        =   11
      Top             =   -15
      Width           =   6885
      Begin VB.CommandButton cmdBuscarProveedor1 
         Height          =   300
         Left            =   1905
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoPagosPorProveedor.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Buscar Proveedor"
         Top             =   780
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.ComboBox cboTipoProveedor 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   3375
      End
      Begin VB.TextBox txtCodProveedor 
         Height          =   300
         Left            =   975
         MaxLength       =   40
         TabIndex        =   1
         Top             =   780
         Width           =   900
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
         Left            =   2340
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Descripción"
         Top             =   780
         Width           =   4410
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   1110
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   16842753
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   4650
         TabIndex        =   4
         Top             =   1110
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   16842753
         CurrentDate     =   41098
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   960
         TabIndex        =   18
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3675
         TabIndex        =   17
         Top             =   1170
         Width           =   960
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Prov.:"
         Height          =   195
         Left            =   105
         TabIndex        =   16
         Top             =   450
         Width           =   780
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
         Left            =   345
         TabIndex        =   15
         Top             =   795
         Width           =   540
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3330
      Top             =   2535
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2835
      Top             =   2445
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameImpresora 
      Caption         =   "impresora"
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
      Left            =   60
      TabIndex        =   12
      Top             =   1650
      Width           =   6885
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   6
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   210
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   315
         Width           =   585
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
      Height          =   345
      Left            =   150
      TabIndex        =   14
      Top             =   2610
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoPagosPorProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SumaSaldo As Double
Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
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

Private Sub cmdListar_Click()
    lblEstado.Caption = "Buscando Listado..."
    'Rep.WindowState = crptMaximized 'crptMinimized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.SelectionFormula = ""
    
    If txtCodProveedor.Text <> "" Then
        Rep.SelectionFormula = "{PAGOS_PROVEEDORES_V.TPR_CODIGO}=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
        Rep.SelectionFormula = Rep.SelectionFormula & " AND {PAGOS_PROVEEDORES_V.PROV_CODIGO}=" & txtCodProveedor.Text
    End If
    If Not IsNull(FechaDesde.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {PAGOS_PROVEEDORES_V.OPG_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PAGOS_PROVEEDORES_V.OPG_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        End If
    End If
    If Not IsNull(FechaHasta.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {PAGOS_PROVEEDORES_V.OPG_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                           
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PAGOS_PROVEEDORES_V.OPG_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        End If
    End If
    
    If Not IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf Not IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    Rep.WindowTitle = "Listado de Pagos a Proveedores"
    Rep.ReportFileName = DRIVE & DirReport & "rptpagosrealizadosaproveedores.rpt"
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     lblEstado.Caption = ""
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
End Sub

Private Sub CmdNuevo_Click()
    txtCodProveedor.Text = ""
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    cboTipoProveedor.ListIndex = 0
    cboTipoProveedor.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoPagosPorProveedor = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    'CARGO COMBO TIPO PROVEEDOR
    LlenarComboTipoProv
    lblEstado.Caption = ""
End Sub

Private Sub txtCodProveedor_Change()
    If txtCodProveedor.Text = "" Then
        txtProvRazSoc.Text = ""
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
            Call BuscaCodigoProxItemData(CInt(Rec1!TPR_CODIGO), cboTipoProveedor)
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtProvRazSoc_Change()
    If txtProvRazSoc.Text = "" Then
        txtCodProveedor.Text = ""
    End If
End Sub

Private Sub txtProvRazSoc_GotFocus()
    SelecTexto txtProvRazSoc
End Sub

Private Sub txtProvRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtProvRazSoc_LostFocus()
    If ActiveControl.Name = "txtCodProveedor" Or _
       cboTipoProveedor.List(cboTipoProveedor.ListIndex) = "TODOS" Then Exit Sub
       
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
    End If
End Sub

Private Function BuscoProveedor(Pro As String) As String
    sql = "SELECT TPR_CODIGO,PROV_CODIGO, PROV_RAZSOC,"
    sql = sql & " PROV_DOMICI"
    sql = sql & " FROM PROVEEDOR "
    sql = sql & " WHERE"
    If txtCodProveedor.Text <> "" Then
        sql = sql & " PROV_CODIGO=" & XN(Pro)
    Else
        sql = sql & " PROV_RAZSOC LIKE '" & Pro & "%'"
    End If
    If cboTipoProveedor.List(cboTipoProveedor.ListIndex) <> "TODOS" Then
        sql = sql & " AND TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    End If

    BuscoProveedor = sql
End Function

Private Sub LlenarComboTipoProv()
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboTipoProveedor.AddItem "TODOS"
        Do While rec.EOF = False
            cboTipoProveedor.AddItem rec!TPR_DESCRI
            cboTipoProveedor.ItemData(cboTipoProveedor.NewIndex) = rec!TPR_CODIGO
            rec.MoveNext
        Loop
        cboTipoProveedor.ListIndex = 0
    End If
    rec.Close
End Sub

