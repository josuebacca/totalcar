VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListadoRemitoProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Remito de Proveedor"
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
      DisabledPicture =   "frmListadoRemitoProveedor.frx":0000
      Height          =   750
      Left            =   8865
      Picture         =   "frmListadoRemitoProveedor.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5535
      Width           =   870
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   6615
      Top             =   4890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   -15
      TabIndex        =   26
      Top             =   6300
      Width           =   10650
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   7275
      Top             =   4935
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   9750
      Picture         =   "frmListadoRemitoProveedor.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5535
      Width           =   840
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   7995
      Picture         =   "frmListadoRemitoProveedor.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5535
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
      TabIndex        =   24
      Top             =   4440
      Width           =   10425
      Begin VB.OptionButton optDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado"
         Height          =   255
         Left            =   2370
         TabIndex        =   7
         Top             =   225
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton optGeneralTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado General"
         Height          =   210
         Left            =   5265
         TabIndex        =   8
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
      TabIndex        =   20
      Top             =   5055
      Width           =   7845
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   10
         Top             =   360
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   435
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   660
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   840
         Width           =   840
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Remito de Proveedor por..."
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
      Left            =   105
      TabIndex        =   14
      Top             =   75
      Width           =   10395
      Begin VB.CheckBox chkCliente 
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   300
         TabIndex        =   0
         Top             =   615
         Width           =   1215
      End
      Begin VB.CheckBox chkFecha 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   960
         Width           =   810
      End
      Begin VB.TextBox txtCliente 
         Height          =   300
         Left            =   3360
         MaxLength       =   40
         TabIndex        =   2
         Top             =   495
         Width           =   975
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
         TabIndex        =   16
         Tag             =   "Descripción"
         Top             =   495
         Width           =   4620
      End
      Begin VB.CommandButton CmdBuscAprox 
         Height          =   1155
         Left            =   9660
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoRemitoProveedor.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Buscar Nota de Pedido"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton cmdBuscarCli 
         Height          =   300
         Left            =   4410
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoRemitoProveedor.frx":398A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Buscar"
         Top             =   510
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   3360
         TabIndex        =   3
         Top             =   960
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
         TabIndex        =   4
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   56950785
         CurrentDate     =   41098
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
         TabIndex        =   19
         Top             =   540
         Width           =   780
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   2265
         TabIndex        =   18
         Top             =   1005
         Width           =   1005
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   4935
         TabIndex        =   17
         Top             =   1020
         Width           =   960
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   2715
      Left            =   90
      TabIndex        =   6
      Top             =   1725
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   4789
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
      TabIndex        =   25
      Top             =   6555
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoRemitoProveedor"
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

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    GrdModulos.HighLight = flexHighlightNever
    sql = "SELECT RC.RPR_NUMERO, RC.RPR_SUCURSAL,"
    sql = sql & "RC.RPR_FECHA, C.PROV_RAZSOC, C.PROV_DOMICI,L.LOC_DESCRI, P.PRO_DESCRI"
    sql = sql & " FROM REMITO_PROVEEDOR RC, PROVEEDOR C, LOCALIDAD L, PROVINCIA P"
    sql = sql & " WHERE"
    sql = sql & " RC.PROV_CODIGO=C.PROV_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.PROV_CODIGO=" & XN(txtCliente)
    If Not IsNull(FechaDesde) Then sql = sql & " AND RC.RPR_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND RC.RPR_FECHA<=" & XDQ(FechaHasta)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        GrdModulos.HighLight = flexHighlightAlways
        Do While rec.EOF = False
            GrdModulos.AddItem Format(rec!RPR_SUCURSAL, "0000") & "-" & Format(rec!RPR_NUMERO, "00000000") _
                            & Chr(9) & rec!RPR_FECHA _
                            & Chr(9) & rec!PROV_RAZSOC & Chr(9) & rec!PROV_DOMICI _
                            & Chr(9) & rec!LOC_DESCRI & Chr(9) & rec!PRO_DESCRI
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
            Rep.SelectionFormula = "{REMITO_PROVEEDOR.PROV_CODIGO}=" & txtCliente.Text
            Rep.Formulas(0) = "CLIENTE='" & "Proveedor: " & txtDesCli & "'"
        Else
            Rep.Formulas(0) = "CLIENTE='" & "Proveedor: Todos'"
        End If
        If Not IsNull(FechaDesde.Value) Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {REMITO_PROVEEDOR.RPR_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_PROVEEDOR.RPR_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            End If
        End If
        
        If Not IsNull(FechaHasta.Value) Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {REMITO_PROVEEDOR.RPR_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_PROVEEDOR.RPR_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
            End If
        End If
        
        If Not IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
        ElseIf Not IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
        ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
            Rep.Formulas(2) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
        End If
        
'            Rep.WindowTitle = "Remito de Proveedor - General - Por cuenta y orden de Terceros"
'            Rep.ReportFileName = DRIVE & DirReport & "rptremitoclientegeneralTerceros.rpt"
        
        Rep.WindowTitle = "Remito de Proveedor - General..."
        Rep.ReportFileName = DRIVE & DirReport & "rptremitoproveedorgeneral.rpt"
    End If
    
    If optDetallado.Value = True Then
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 0) = "" Then
            MsgBox "Debe seleccionar un Remito", vbExclamation, TIT_MSGBOX
            chkCliente.SetFocus
            Exit Sub
        End If
        Rep.SelectionFormula = ""
        Rep.SelectionFormula = "{REMITO_PROVEEDOR.RPR_NUMERO}=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)) _
                               & " AND {REMITO_PROVEEDOR.RPR_SUCURSAL}=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4))
                               
        Rep.WindowTitle = "Remito de Proveedor - Detallado..."
        Rep.ReportFileName = DRIVE & DirReport & "rptremitoproveedordetalle.rpt"
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
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    GrdModulos.Rows = 1
    GrdModulos.Rows = 2
    txtCliente.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    chkCliente.Value = Unchecked
    chkFecha.Value = Unchecked
    optDetallado.Value = True
    optPantalla.Value = True
    chkCliente.SetFocus
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
       
    txtCliente.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    cmdBuscarCli.Enabled = False
    GrdModulos.Rows = 1
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""

    Call Centrar_pantalla(Me)
    GrdModulos.FormatString = "^Número|^Fecha|Proveedor|Domicilio|Localidad|Provincia"
    GrdModulos.ColWidth(0) = 1300
    GrdModulos.ColWidth(1) = 1200
    GrdModulos.ColWidth(2) = 3200
    GrdModulos.ColWidth(3) = 3200
    GrdModulos.ColWidth(4) = 3200
    GrdModulos.ColWidth(5) = 3200
    GrdModulos.Rows = 2
    '------------------------------------
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
    If chkFecha.Value = Unchecked _
        And ActiveControl.Name <> "cmdBuscarCli" _
        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub
Private Sub CmdSalir_Click()
    Set frmListadoRemitoProveedor = Nothing
    Unload Me
End Sub


