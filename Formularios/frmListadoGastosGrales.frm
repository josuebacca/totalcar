VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListadoGastosGrales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Gastos Generales"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      Top             =   6375
      Width           =   7845
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   435
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   660
         Width           =   1665
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   27
         Top             =   360
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   26
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label lblImpresora 
         AutoSize        =   -1  'True
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1965
         TabIndex        =   30
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   29
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ver..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      TabIndex        =   22
      Top             =   5760
      Width           =   10425
      Begin VB.OptionButton optGeneralTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado General"
         Height          =   210
         Left            =   5265
         TabIndex        =   24
         Top             =   240
         Width           =   1620
      End
      Begin VB.OptionButton optDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado"
         Height          =   255
         Left            =   2370
         TabIndex        =   23
         Top             =   225
         Value           =   -1  'True
         Width           =   1770
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   8025
      Picture         =   "frmListadoGastosGrales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6855
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   9780
      Picture         =   "frmListadoGastosGrales.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6855
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoGastosGrales.frx":0BD4
      Height          =   750
      Left            =   8895
      Picture         =   "frmListadoGastosGrales.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6855
      Width           =   870
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
      Left            =   100
      TabIndex        =   10
      Top             =   0
      Width           =   10755
      Begin VB.CommandButton cmdBuscarProveedor 
         Height          =   300
         Left            =   4110
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoGastosGrales.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Buscar Proveedor"
         Top             =   615
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CheckBox chkTipoGasto 
         Caption         =   "Tipo Gasto"
         Height          =   195
         Left            =   720
         TabIndex        =   2
         Top             =   1010
         Width           =   1155
      End
      Begin VB.CommandButton CmdBuscAprox 
         Height          =   465
         Left            =   9120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmListadoGastosGrales.frx":14F2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Buscar "
         Top             =   1155
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtProveedor 
         Height          =   300
         Left            =   3105
         MaxLength       =   40
         TabIndex        =   5
         Top             =   615
         Width           =   975
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
         Left            =   4545
         MaxLength       =   50
         TabIndex        =   11
         Tag             =   "Descripción"
         Top             =   615
         Width           =   4440
      End
      Begin VB.CheckBox chkFecha 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   720
         TabIndex        =   3
         Top             =   1350
         Width           =   810
      End
      Begin VB.CheckBox chkProveedor 
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Top             =   670
         Width           =   1125
      End
      Begin VB.CheckBox chkTipoProveedor 
         Caption         =   "Tipo Prov"
         Height          =   195
         Left            =   720
         TabIndex        =   0
         Top             =   330
         Width           =   1050
      End
      Begin VB.ComboBox cboBuscaTipoProveedor 
         Height          =   315
         Left            =   3105
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   270
         Width           =   3990
      End
      Begin VB.ComboBox cboBuscaTipoGasto 
         Height          =   315
         Left            =   3105
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   4005
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   3120
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   16711681
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   5685
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   16711681
         CurrentDate     =   41098
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   4680
         TabIndex        =   17
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   2040
         TabIndex        =   16
         Top             =   1350
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
         Index           =   2
         Left            =   2265
         TabIndex        =   15
         Top             =   645
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Prov.:"
         Height          =   195
         Left            =   2265
         TabIndex        =   14
         Top             =   315
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Gasto:"
         Height          =   195
         Left            =   2580
         TabIndex        =   13
         Top             =   990
         Width           =   465
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3855
      Left            =   75
      TabIndex        =   18
      Top             =   1785
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   19
      FixedCols       =   0
      BackColorSel    =   8388736
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   8565
      Top             =   6210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   8025
      Top             =   6255
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
      Left            =   195
      TabIndex        =   31
      Top             =   7680
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoGastosGrales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub CmdBuscAprox_Click()
    Set Rec1 = New ADODB.Recordset
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
                               Rec1!GGR_NROCOMP & Chr(9) & Valido_Importe(Rec1!GGR_NETO) & Chr(9) & _
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

Private Sub cmdListar_Click()
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    
    If optGeneralTodos.Value = True Then
        Rep.SelectionFormula = ""
        If txtProveedor.Text <> "" Then
            Rep.SelectionFormula = "{GASTOS_GENERALES.PROV_CODIGO}=" & txtProveedor.Text
            Rep.Formulas(0) = "PROVEEDOR='" & "Proveedor: " & txtDesProv & "'"
        Else
            Rep.Formulas(0) = "PROVEEDOR='" & "Proveedor: Todos'"
        End If
        If Not IsNull(FechaDesde.Value) Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {GASTOS_GENERALES.FECHACOM_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {GASTOS_GENERALES.FECHACOM_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            End If
        End If
        
        If Not IsNull(FechaHasta.Value) Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {GASTOS_GENERALES.FECHACOM_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {GASTOS_GENERALES.FECHACOM_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
            End If
        End If
'        If optPen.Value = True Then
'            If Rep.SelectionFormula = "" Then
'                Rep.SelectionFormula = "{ESTADO_DOCUMENTO.EST_CODIGO}= 1 "
'            Else
'                Rep.SelectionFormula = Rep.SelectionFormula & "AND {ESTADO_DOCUMENTO.EST_CODIGO}= 1 "
'            End If
'        Else
'            If optDef.Value = True Then
'                If Rep.SelectionFormula = "" Then
'                    Rep.SelectionFormula = "{ESTADO_DOCUMENTO.EST_CODIGO}= 3 "
'                Else
'                    Rep.SelectionFormula = Rep.SelectionFormula & "AND {ESTADO_DOCUMENTO.EST_CODIGO}= 3 "
'                End If
'            Else
'                If optAnu.Value = True Then
'                    If Rep.SelectionFormula = "" Then
'                        Rep.SelectionFormula = "{ESTADO_DOCUMENTO.EST_CODIGO}= 2 "
'                    Else
'                        Rep.SelectionFormula = Rep.SelectionFormula & "{ESTADO_DOCUMENTO.EST_CODIGO}= 2 "
'                    End If
'                End If
'
'            End If
'        End If
        
        
        If Not IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
        ElseIf Not IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
        ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
            Rep.Formulas(2) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
        End If
        
'            Rep.WindowTitle = "Remito de Cliente - General - Por cuenta y orden de Terceros"
'            Rep.ReportFileName = DRIVE & DirReport & "rptremitoclientegeneralTerceros.rpt"
        
        Rep.WindowTitle = "Gastos Generales - General..."
        Rep.ReportFileName = DRIVE & DirReport & "rptGastosgeneral.rpt"
    End If
    
    If optDetallado.Value = True Then
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 0) = "" Then
            MsgBox "Debe seleccionar un Remito", vbExclamation, TIT_MSGBOX
            chkCliente.SetFocus
            Exit Sub
        End If
        Rep.SelectionFormula = ""
        Rep.SelectionFormula = "{GASTOS_GENERALES.GGR_NROCOMP}=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 9), 8)) _
                               & " AND {GASTOS_GENERALES.GGR_NROSUC}=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 8), 4))
                               
        Rep.WindowTitle = "Gastos Generales - Detallado..."
        Rep.ReportFileName = DRIVE & DirReport & "rptGastosDetalle.rpt"
    End If
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     Rep.Formulas(2) = ""
End Sub

Private Sub CmdNuevo_Click()
    chkTipoProveedor.Value = Unchecked
    chkProveedor.Value = Unchecked
    chkTipoGasto.Value = Unchecked
    chkFecha.Value = Unchecked
    GrdModulos.Rows = 1
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

 
    
Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
       
    txtProveedor.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    cmdBuscarProveedor.Enabled = False
    GrdModulos.Rows = 1
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""
    
    cboBuscaTipoGasto.Enabled = False
    cboBuscaTipoProveedor.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    txtProveedor.Enabled = False
    'CARGO COMBO TIPO PROVEEDOR
    LlenarComboTipoProv
    'CARGO COMBO GASTOS
    LlenarComboGastos
   
    
    Call Centrar_pantalla(Me)
    GrdModulos.FormatString = "Proveedor|Gasto|Comprobante|^Fecha|TIPO PROVEEDOR|" _
                            & "COD PROVEEDOR|COD TIPO GASTO|COD TIP COMPROBANTE|" _
                            & "NRO SUCURSAL|NRO COMPROBANTE|Importe|IVA|NETO1|IVA1|IMPUESTOS|PERIODO|ENTRA LIBRO IVA"
                            
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
    GrdModulos.ColWidth(10) = 1500   'NETO
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
    
    
    sql = "UPDATE GASTOS_GENERALES SET FPG_CODIGO= 1"
    DBConn.Execute sql
End Sub
Private Sub LlenarComboTipoProv()
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboBuscaTipoProveedor.AddItem "TODOS"
        Do While rec.EOF = False
            cboBuscaTipoProveedor.AddItem rec!TPR_DESCRI
            cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.NewIndex) = rec!TPR_CODIGO
            'cboBuscaTipoProveedor.AddItem rec!TPR_DESCRI
            'cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.NewIndex) = rec!TPR_CODIGO
            rec.MoveNext
        Loop
        cboBuscaTipoProveedor.ListIndex = -1
    End If
    rec.Close
End Sub
Private Sub LlenarComboGastos()
    sql = "SELECT * FROM TIPO_GASTO ORDER BY TGT_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboBuscaTipoGasto.AddItem rec!TGT_DESCRI
            cboBuscaTipoGasto.ItemData(cboBuscaTipoGasto.NewIndex) = rec!TGT_CODIGO
            'cboBuscaTipoGasto.AddItem rec!TGT_DESCRI
            'cboBuscaTipoGasto.ItemData(cboBuscaTipoGasto.NewIndex) = rec!TGT_CODIGO
            rec.MoveNext
        Loop
        cboBuscaTipoGasto.ListIndex = -1
    End If
    rec.Close
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
    If cboBuscaTipoProveedor.List(cboBuscaTipoProveedor.ListIndex) <> "TODOS" Then
        sql = sql & " AND PRO.TPR_CODIGO=" & cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.ListIndex)
    End If
    sql = sql & " AND PRO.LOC_CODIGO=L.LOC_CODIGO"

    BuscoProveedor = sql
End Function
