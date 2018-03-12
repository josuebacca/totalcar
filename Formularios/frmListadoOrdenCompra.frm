VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListadoOrdenCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Orden de Compra"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Nota de Pedido por..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   90
      TabIndex        =   24
      Top             =   0
      Width           =   10395
      Begin VB.CommandButton cmdBuscarCli 
         Height          =   315
         Left            =   4410
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoOrdenCompra.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Buscar"
         Top             =   345
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.TextBox txtVendedor 
         Height          =   300
         Left            =   3360
         TabIndex        =   4
         Top             =   810
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
         TabIndex        =   27
         Top             =   825
         Width           =   4620
      End
      Begin VB.CheckBox chkVendedor 
         Caption         =   "Empleado"
         Height          =   195
         Left            =   540
         TabIndex        =   1
         Top             =   840
         Width           =   1035
      End
      Begin VB.CommandButton CmdBuscAprox 
         Height          =   1455
         Left            =   9645
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoOrdenCompra.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Buscar Nota de Pedido"
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
         TabIndex        =   26
         Tag             =   "Descripción"
         Top             =   345
         Width           =   4620
      End
      Begin VB.TextBox txtCliente 
         Height          =   300
         Left            =   3360
         MaxLength       =   40
         TabIndex        =   3
         Top             =   345
         Width           =   975
      End
      Begin VB.CheckBox chkFecha 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   540
         TabIndex        =   2
         Top             =   1230
         Width           =   810
      End
      Begin VB.CheckBox chkCliente 
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   540
         TabIndex        =   0
         Top             =   450
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscarVendedor 
         Height          =   300
         Left            =   4410
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoOrdenCompra.frx":2AAC
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Buscar Vendedor"
         Top             =   825
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   3360
         TabIndex        =   5
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
         Left            =   5925
         TabIndex        =   6
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empleado:"
         Height          =   195
         Left            =   2535
         TabIndex        =   32
         Top             =   855
         Width           =   750
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   4935
         TabIndex        =   31
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   2265
         TabIndex        =   30
         Top             =   1335
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
         TabIndex        =   29
         Top             =   390
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoOrdenCompra.frx":2DB6
      Height          =   750
      Left            =   8865
      Picture         =   "frmListadoOrdenCompra.frx":30C0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5745
      Width           =   870
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   6615
      Top             =   5100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   -15
      TabIndex        =   23
      Top             =   6510
      Width           =   10650
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   7275
      Top             =   5145
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   9750
      Picture         =   "frmListadoOrdenCompra.frx":33CA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5745
      Width           =   840
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   7995
      Picture         =   "frmListadoOrdenCompra.frx":36D4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5745
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
      TabIndex        =   21
      Top             =   4635
      Width           =   10425
      Begin VB.OptionButton optDetalladoVarios 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado (Varios)"
         Height          =   255
         Left            =   3645
         TabIndex        =   10
         Top             =   240
         Width           =   2340
      End
      Begin VB.OptionButton optDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado (Una)"
         Height          =   255
         Left            =   735
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optGeneralTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado General"
         Height          =   210
         Left            =   7155
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
      TabIndex        =   17
      Top             =   5265
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
         TabIndex        =   18
         Top             =   660
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   840
         Width           =   840
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   2835
      Left            =   90
      TabIndex        =   8
      Top             =   1800
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   5001
      _Version        =   393216
      Cols            =   5
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
      TabIndex        =   22
      Top             =   6690
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CBImpresora_Click()
    If IMPRIMO = "EPSON" Then
        CDImpresora.PrinterDefault = True
        CDImpresora.ShowPrinter
        lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    End If
End Sub
Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    GrdModulos.HighLight = flexHighlightNever
    sql = "SELECT NP.*, C.PROV_RAZSOC,C.PROV_DOMICI,L.LOC_DESCRI,P.PRO_DESCRI"
    sql = sql & " FROM ORDEN_COMPRA NP, PROVEEDOR C"
    sql = sql & ", LOCALIDAD L, PROVINCIA P"
    sql = sql & " WHERE"
    sql = sql & " NP.PROV_CODIGO=C.PROV_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND NP.PROV_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
    'If chkRepresentada.Value = Checked Then sql = sql & " AND NP.REP_CODIGO=" & cboRepresentada.ItemData(cboRepresentada.ListIndex)
    If Not IsNull(FechaDesde) Then sql = sql & " AND NP.OC_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND NP.OC_FECHA<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY OC_NUMERO,OC_FECHA"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        GrdModulos.HighLight = flexHighlightAlways
        Do While rec.EOF = False
            GrdModulos.AddItem Format(rec!OC_NUMERO, "00000000") & Chr(9) & rec!OC_FECHA _
                            & Chr(9) & rec!PROV_RAZSOC & Chr(9) & rec!LOC_DESCRI _
                            & Chr(9) & rec!PRO_DESCRI & Chr(9) & ""
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

Private Sub cmdBuscarVendedor_Click()
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

Private Sub cmdListar_Click()
    'Rep.WindowState = crptMaximized 'crptMinimized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    Rep.Formulas(3) = ""
    
    'NOTA DE PEDIDO GENERAL
    If GrdModulos.Rows > 1 Then
        If optGeneralTodos.Value = True Then
            Rep.SelectionFormula = ""
            If txtCliente.Text <> "" Then
                Rep.SelectionFormula = "{ORDEN_COMPRA.PROV_CODIGO}=" & txtCliente.Text
                Rep.Formulas(0) = "CLIENTE='" & "Proveedor: " & txtDesCli & "'"
            Else
                Rep.Formulas(0) = "CLIENTE='" & "Proveedor: Todos'"
            End If
                    
            If Not IsNull(FechaDesde.Value) Then
                If Rep.SelectionFormula = "" Then
                    Rep.SelectionFormula = " {ORDEN_COMPRA.OC_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
                Else
                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {ORDEN_COMPRA.OC_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
                End If
            End If
            
            If Not IsNull(FechaHasta.Value) Then
                If Rep.SelectionFormula = "" Then
                    Rep.SelectionFormula = " {ORDEN_COMPRA.OC_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                Else
                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {ORDEN_COMPRA.OC_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                End If
            End If
            
            If Not IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
                Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
            ElseIf Not IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
                Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
            ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
                Rep.Formulas(2) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
            End If
        
            Rep.WindowTitle = "Orden de Compra - General..."
            Rep.ReportFileName = DRIVE & DirReport & "rptordencomprageneral.rpt"
        End If
        
        'NOTA DE PEDIDO DETALLADO (UNA NOTA DE PEDIDO SOLA)
        If optDetallado.Value = True Then
            Rep.Formulas(0) = ""
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 0) = "" Then
                MsgBox "Debe seleccionar una Nota de Pedido", vbExclamation, TIT_MSGBOX
                chkCliente.SetFocus
                Exit Sub
            End If
            Rep.SelectionFormula = ""
            Rep.SelectionFormula = "{ORDEN_COMPRA.OC_NUMERO}=" & Int(Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0))) _
                                   & " AND {ORDEN_COMPRA.OC_FECHA}= DATE (" & Mid(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 7, 4) & "," & Mid(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4, 2) & "," & Mid(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 1, 2) & ")"
            
            Rep.WindowTitle = "Orden de Compra - Detallado..."
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "" Then 'SI VA SIN DETALLE
                Rep.ReportFileName = DRIVE & DirReport & "rptordencompra.rpt"
            Else 'SI VA CON DETALLE
                Rep.ReportFileName = DRIVE & DirReport & "rptnotapedidodetallePrecio.rpt"
            End If
        End If
        
        'NOTA DE PEDIDO DETALLE (VARIOS)
        If optDetalladoVarios.Value = True Then
            Rep.Formulas(0) = ""
            Rep.SelectionFormula = ""
            If txtCliente.Text <> "" Then
                Rep.SelectionFormula = "{ORDEN_COMPRA.PROV_CODIGO}=" & txtCliente.Text
            End If
            
            
            If Not IsNull(FechaDesde.Value) Then
                If Rep.SelectionFormula = "" Then
                    Rep.SelectionFormula = " {ORDEN_COMPRA.OC_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
                Else
                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {ORDEN_COMPRA.OC_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
                End If
            End If
            
            If Not IsNull(FechaHasta.Value) Then
                If Rep.SelectionFormula = "" Then
                    Rep.SelectionFormula = " {ORDEN_COMPRA.OC_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                Else
                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {ORDEN_COMPRA.OC_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                End If
            End If
            
            If Not IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
                Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
            ElseIf Not IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
                Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
            ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
                Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
            End If
        
            Rep.WindowTitle = "Orden de Compra - Detallado..."
            Rep.ReportFileName = DRIVE & DirReport & "rptordencompra.rpt"
        End If
        
        If optPantalla.Value = True Then
            Rep.Destination = crptToWindow
        ElseIf optImpresora.Value = True Then
            Rep.Destination = crptToPrinter
        End If
    End If
    
     Rep.WindowState = crptMaximized
     Rep.Action = 1
     'Rep.WindowState = crptMaximized
     'Rep.WindowState = crptNormal
     
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     Rep.Formulas(2) = ""
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
Private Sub CmdNuevo_Click()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    txtVendedor.Text = ""
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    GrdModulos.Rows = 1
    GrdModulos.Rows = 2
    txtCliente.Enabled = False
    txtVendedor.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    chkCliente.Value = Unchecked
    chkVendedor.Value = Unchecked
    chkFecha.Value = Unchecked
    optDetallado.Value = True
    optPantalla.Value = True
    chkCliente.SetFocus
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
       
    txtCliente.Enabled = False
    txtVendedor.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    cmdBuscarCli.Enabled = False
    cmdBuscarVendedor.Enabled = False
    GrdModulos.Rows = 1
    If IMPRIMO = "EPSON" Then
        lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    End If
    lblEstado.Caption = ""

    Call Centrar_pantalla(Me)
    GrdModulos.FormatString = ">Número|^Fecha|Proveedor|Localidad|Provincia|Forma de pago"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 1000
    GrdModulos.ColWidth(2) = 3200
    GrdModulos.ColWidth(3) = 3200
    GrdModulos.ColWidth(4) = 3200
    GrdModulos.ColWidth(5) = 0
    GrdModulos.Rows = 2
    '------------------------------------
    
    
    optDetallado.Value = True
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
        cmdBuscarVendedor.Enabled = True
    Else
        txtVendedor.Enabled = False
        cmdBuscarVendedor.Enabled = False
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
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtDesCli.Text = ""
            txtCliente.SetFocus
        End If
        rec.Close
    End If
    If chkFecha.Value = Unchecked _
        And chkVendedor.Value = Unchecked And ActiveControl.Name <> "cmdBuscarCli" _
        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "cmdSalir" Then CmdBuscAprox.SetFocus
End Sub
Private Sub CmdSalir_Click()
    Set frmListadoOrdenCompra = Nothing
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
    If chkFecha.Value = Unchecked And ActiveControl.Name <> "cmdNuevo" _
        And ActiveControl.Name <> "cmdSalir" Then CmdBuscAprox.SetFocus
End Sub

