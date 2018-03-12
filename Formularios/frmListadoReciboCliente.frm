VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListadoReciboCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Recibo de Cliente"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoReciboCliente.frx":0000
      Height          =   750
      Left            =   8865
      Picture         =   "frmListadoReciboCliente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5655
      Width           =   870
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   6615
      Top             =   5010
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   -15
      TabIndex        =   31
      Top             =   6420
      Width           =   10650
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   7290
      Top             =   5055
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   9750
      Picture         =   "frmListadoReciboCliente.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5655
      Width           =   840
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   7995
      Picture         =   "frmListadoReciboCliente.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5655
      Width           =   855
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
      Left            =   90
      TabIndex        =   29
      Top             =   4560
      Width           =   10425
      Begin VB.OptionButton optDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado"
         Height          =   255
         Left            =   5475
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton optGeneralTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado General"
         Height          =   210
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1620
      End
   End
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
      Left            =   90
      TabIndex        =   25
      Top             =   5175
      Width           =   7845
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   13
         Top             =   360
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   435
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   660
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   360
         Width           =   585
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
         TabIndex        =   27
         Top             =   840
         Width           =   840
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Recibo de Cliente por..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   105
      TabIndex        =   17
      Top             =   75
      Width           =   10395
      Begin VB.CheckBox chkCliente 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   315
         TabIndex        =   33
         Top             =   495
         Width           =   855
      End
      Begin VB.CheckBox chkTipoRecibo 
         Caption         =   "Tipo de Recibo"
         Height          =   195
         Left            =   315
         TabIndex        =   1
         Top             =   975
         Width           =   1485
      End
      Begin VB.ComboBox cboRecibo 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   930
         Width           =   2400
      End
      Begin VB.CheckBox chkFecha 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   315
         TabIndex        =   2
         Top             =   1230
         Width           =   810
      End
      Begin VB.TextBox txtCliente 
         Height          =   300
         Left            =   3360
         MaxLength       =   40
         TabIndex        =   3
         Top             =   255
         Width           =   990
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
         TabIndex        =   20
         Tag             =   "Descripción"
         Top             =   255
         Width           =   4620
      End
      Begin VB.CommandButton CmdBuscAprox 
         Height          =   1455
         Left            =   9660
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoReciboCliente.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Buscar Nota de Pedido"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CheckBox chkVendedor 
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   315
         TabIndex        =   0
         Top             =   735
         Width           =   1035
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
         TabIndex        =   19
         Top             =   615
         Width           =   5055
      End
      Begin VB.TextBox txtVendedor 
         Height          =   300
         Left            =   3360
         TabIndex        =   4
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton cmdBuscarCli 
         Height          =   315
         Left            =   4410
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoReciboCliente.frx":398A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Buscar"
         Top             =   255
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   3360
         TabIndex        =   6
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   56623105
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   5970
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   56623105
         CurrentDate     =   41098
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   2910
         TabIndex        =   32
         Top             =   960
         Width           =   360
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
         TabIndex        =   24
         Top             =   300
         Width           =   525
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   2265
         TabIndex        =   23
         Top             =   1335
         Width           =   1005
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   4935
         TabIndex        =   22
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   2535
         TabIndex        =   21
         Top             =   645
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   2775
      Left            =   90
      TabIndex        =   9
      Top             =   1875
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   8388736
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
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
      Left            =   165
      TabIndex        =   30
      Top             =   6555
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoReciboCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub cboRecibo_LostFocus()
    If chkFecha.Value = Unchecked And ActiveControl.Name <> "cmdNuevo" _
        And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub
Private Sub chkTipoRecibo_Click()
    If chkTipoRecibo.Value = Checked Then
        cboRecibo.Enabled = True
        cboRecibo.ListIndex = 0
    Else
        cboRecibo.Enabled = False
        cboRecibo.ListIndex = -1
    End If
End Sub

Private Sub CmdBuscAprox_Click()
        
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    GrdModulos.HighLight = flexHighlightNever
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT RC.REC_NUMERO, RC.REC_SUCURSAL, RC.REC_FECHA,"
    sql = sql & " RC.TCO_CODIGO, TC.TCO_ABREVIA,"
    sql = sql & " C.CLI_RAZSOC, V.VEN_NOMBRE"
    sql = sql & " FROM RECIBO_CLIENTE RC, CLIENTE C, VENDEDOR V, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " RC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND RC.VEN_CODIGO=V.VEN_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND RC.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND RC.REC_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND RC.REC_FECHA<=" & XDQ(FechaHasta)
    If chkTipoRecibo.Value = Checked Then sql = sql & " AND RC.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    sql = sql & " ORDER BY RC.REC_SUCURSAL,RC.REC_NUMERO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            GrdModulos.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!REC_SUCURSAL, "0000") & "-" & Format(Rec1!REC_NUMERO, "00000000") _
                               & Chr(9) & Rec1!REC_FECHA & Chr(9) & Rec1!CLI_RAZSOC _
                               & Chr(9) & Rec1!VEN_NOMBRE & Chr(9) & "" _
                               & Chr(9) & Rec1!TCO_CODIGO
            Rec1.MoveNext
        Loop
        GrdModulos.HighLight = flexHighlightAlways
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Recibos...", vbExclamation, TIT_MSGBOX
        chkCliente.SetFocus
    End If
    Rec1.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdListar_Click()
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    
    If optGeneralTodos.Value = True Then
        Rep.SelectionFormula = ""
        If txtCliente.Text <> "" Then
            Rep.SelectionFormula = "{RECIBO_CLIENTE.CLI_CODIGO}=" & txtCliente.Text
            Rep.Formulas(0) = "CLIENTE='" & "Cliente: " & txtDesCli & "'"
        Else
            Rep.Formulas(0) = "CLIENTE='" & "Cliente: Todos'"
        End If
        
        If Not IsNull(FechaDesde.Value) Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {RECIBO_CLIENTE.REC_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {RECIBO_CLIENTE.REC_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            End If
        End If
        
        If Not IsNull(FechaHasta.Value) Then
            If Rep.SelectionFormula = "" Then
                                                                                        
                Rep.SelectionFormula = " {RECIBO_CLIENTE.REC_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {RECIBO_CLIENTE.REC_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
            End If
        End If
                        
        If Not IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
        ElseIf Not IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
        ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
            Rep.Formulas(2) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
        ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
            Rep.Formulas(2) = "FECHA='" & "Desde: Inicio   Hasta: " & Date & "'"
        End If

        Rep.WindowTitle = "Recibo de Cliente - General..."
        Rep.ReportFileName = DRIVE & DirReport & "rptreciboclientegeneral.rpt"
    End If
    
    If optDetallado.Value = True Then
         'Exit Sub
         If GrdModulos.TextMatrix(GrdModulos.RowSel, 0) = "" Then
            MsgBox "Debe seleccionar un Recibo", vbExclamation, TIT_MSGBOX
            chkCliente.SetFocus
            Exit Sub
        End If
        Rep.SelectionFormula = ""
                                                                                'XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)
        Rep.SelectionFormula = "{RECIBO_CLIENTE.REC_NUMERO}=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)) _
                               & " AND DAY({RECIBO_CLIENTE.REC_FECHA})=" & Day(GrdModulos.TextMatrix(GrdModulos.RowSel, 2)) _
                               & " AND MONTH({RECIBO_CLIENTE.REC_FECHA})=" & Month(GrdModulos.TextMatrix(GrdModulos.RowSel, 2)) _
                               & " AND YEAR({RECIBO_CLIENTE.REC_FECHA})=" & Year(GrdModulos.TextMatrix(GrdModulos.RowSel, 2))
        Rep.WindowTitle = "Recibo de Cliente - Detallado..."
        Rep.ReportFileName = DRIVE & DirReport & "rptreciboclientedetalle.rpt"
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

Private Sub CmdNuevo_Click()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    txtVendedor.Text = ""
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    cboRecibo.ListIndex = -1
    GrdModulos.Rows = 1
    GrdModulos.Rows = 2
    txtCliente.Enabled = False
    txtVendedor.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    cboRecibo.Enabled = False
    chkCliente.Value = Unchecked
    chkVendedor.Value = Unchecked
    chkFecha.Value = Unchecked
    chkTipoRecibo.Value = Unchecked
    optGeneralTodos.Value = True
    optPantalla.Value = True
    chkCliente.SetFocus
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    
    optGeneralTodos.Value = True
    
    txtCliente.Enabled = False
    txtVendedor.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    cboRecibo.Enabled = False
    cmdBuscarCli.Enabled = False
    GrdModulos.Rows = 1
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""

    Call Centrar_pantalla(Me)
    GrdModulos.FormatString = "Tipo Rec|^Nro Recibo|^Fecha Recibo|Cliente|Vendedor|REPRESENTADA|TIPO RECIBO"
    GrdModulos.ColWidth(0) = 1000 'TIPO_RECIBO
    GrdModulos.ColWidth(1) = 1300 'NRO RECIBO
    GrdModulos.ColWidth(2) = 1200 'FECHA_RECIBO
    GrdModulos.ColWidth(3) = 4500 'CLIENTE
    GrdModulos.ColWidth(4) = 4500 'VENDEDOR
    GrdModulos.ColWidth(5) = 0    'REPRESENTADA
    GrdModulos.ColWidth(6) = 0    'TIPO RECIBO (TCO_CODIGO)
    GrdModulos.Rows = 2
    '------------------------------------
    'LLENAR COMBO RECIBO
    LlenarComboRecibo
End Sub

Private Sub LlenarComboRecibo()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'RECIB%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboRecibo.AddItem rec!TCO_DESCRI
            cboRecibo.ItemData(cboRecibo.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboRecibo.ListIndex = -1
    End If
    rec.Close
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
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

Private Sub chkVendedor_Click()
    If chkVendedor.Value = Checked Then
        txtVendedor.Enabled = True
    Else
        txtVendedor.Enabled = False
    End If
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
    'If chkTipoRecibo.Value = Unchecked And chkFecha.Value = Unchecked _
        And chkVendedor.Value = Unchecked And ActiveControl.Name <> "cmdBuscarCli" _
        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoReciboCliente = Nothing
    Unload Me
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
    'If chkTipoRecibo.Value = Unchecked And chkFecha.Value = Unchecked And ActiveControl.Name <> "cmdNuevo" _
        And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub

