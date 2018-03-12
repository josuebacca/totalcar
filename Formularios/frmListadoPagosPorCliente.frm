VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListadoPagosPorCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Pagos por Cliente"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
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
      Height          =   780
      Left            =   60
      TabIndex        =   15
      Top             =   1920
      Width           =   6690
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   7
         Top             =   330
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   6
         Top             =   330
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   225
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   330
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   4140
      Picture         =   "frmListadoPagosPorCliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   5895
      Picture         =   "frmListadoPagosPorCliente.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoPagosPorCliente.frx":0BD4
      Height          =   750
      Left            =   5010
      Picture         =   "frmListadoPagosPorCliente.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listar por..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   60
      TabIndex        =   12
      Top             =   -15
      Width           =   6690
      Begin VB.CommandButton cmdBuscarCliente 
         Height          =   315
         Left            =   1800
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoPagosPorCliente.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Buscar Cliente"
         Top             =   390
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.ComboBox cboFactura 
         Height          =   315
         Left            =   1755
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   885
         Width           =   1485
      End
      Begin VB.TextBox txtNroFactura 
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
         Left            =   4365
         TabIndex        =   3
         Top             =   885
         Width           =   1365
      End
      Begin VB.TextBox txtDesCli 
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
         Left            =   2235
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Descripción"
         Top             =   390
         Width           =   4350
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   795
         MaxLength       =   40
         TabIndex        =   0
         Top             =   390
         Width           =   975
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1755
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   58982401
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   4365
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   58982401
         CurrentDate     =   41098
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   1320
         TabIndex        =   20
         Top             =   915
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   3570
         TabIndex        =   19
         Top             =   915
         Width           =   600
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
         Left            =   180
         TabIndex        =   17
         Top             =   435
         Width           =   525
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3345
         TabIndex        =   14
         Top             =   1380
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   645
         TabIndex        =   13
         Top             =   1380
         Width           =   1005
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3330
      Top             =   2865
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2835
      Top             =   2775
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   18
      Top             =   2940
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoPagosPorCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCliente.Text = frmBuscar.grdBuscar.Text
        txtCliente_LostFocus
        txtDesCli.SetFocus
    Else
        txtCliente.SetFocus
    End If
End Sub

Private Sub cmdListar_Click()
    
    If txtCliente.Text = "" And txtNroFactura.Text = "" Then
        MsgBox "Debe elegir un Cliente", vbExclamation, TIT_MSGBOX
        txtCliente.SetFocus
        Exit Sub
    End If
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.SelectionFormula = ""
    
    If txtCliente.Text <> "" Then
        Rep.SelectionFormula = "{PAGOS_CLIENTES_V.CLI_CODIGO}=" & txtCliente.Text
    End If
    
    If txtNroFactura.Text <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = "{PAGOS_CLIENTES_V.FCL_TCO_CODIGO}=" & cboFactura.ItemData(cboFactura.ListIndex)
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PAGOS_CLIENTES_V.FCL_NUMERO}=" & XN(txtNroFactura)
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PAGOS_CLIENTES_V.FCL_TCO_CODIGO}=" & cboFactura.ItemData(cboFactura.ListIndex)
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PAGOS_CLIENTES_V.FCL_NUMERO}=" & XN(txtNroFactura)
        End If
    End If
    
    If Not IsNull(FechaDesde.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {PAGOS_CLIENTES_V.REC_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PAGOS_CLIENTES_V.REC_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        End If
    End If
    
    If Not IsNull(FechaHasta.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {PAGOS_CLIENTES_V.REC_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                           
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PAGOS_CLIENTES_V.REC_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
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
    
    Rep.WindowTitle = "Listado de Pagos por Cliente"
    Rep.ReportFileName = DRIVE & DirReport & "listadopagosrealizadosporcliente.rpt"
    
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
    txtCliente.Text = ""
    cboFactura.ListIndex = 0
    txtNroFactura.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    txtCliente.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoPagosPorCliente = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    'CARGO COMBOS
    LlenarComboFactura
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""
End Sub
Private Sub LlenarComboFactura()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'FACT%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboFactura.AddItem rec!TCO_ABREVIA
            cboFactura.ItemData(cboFactura.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboFactura.ListIndex = 0
    End If
    rec.Close
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
        rec.Open BuscoCliente(txtCliente), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtDesCli_Change()
    If txtDesCli.Text = "" Then
        txtCliente.Text = ""
    End If
End Sub

Private Sub txtDesCli_GotFocus()
    SelecTexto txtDesCli
End Sub

Private Sub txtDesCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesCli_LostFocus()
    If txtCliente.Text = "" And txtDesCli.Text <> "" Then
        rec.Open BuscoCliente(txtDesCli), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 1
                frmBuscar.TxtDescriB.Text = txtDesCli.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 0
                    txtCliente.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 1
                    txtDesCli.Text = frmBuscar.grdBuscar.Text
                    FechaDesde.SetFocus
                Else
                    txtCliente.SetFocus
                End If
            Else
                txtCliente.Text = rec!CLI_CODIGO
                txtDesCli.Text = rec!CLI_RAZSOC
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Function BuscoCliente(Cli As String) As String
    sql = "SELECT CLI_CODIGO, CLI_RAZSOC"
    sql = sql & " FROM CLIENTE"
    sql = sql & " WHERE "
    If txtCliente.Text <> "" Then
        sql = sql & " CLI_CODIGO=" & XN(Cli)
    Else
        sql = sql & " CLI_RAZSOC LIKE '" & Cli & "%'"
    End If
    BuscoCliente = sql
End Function

Private Sub txtNroFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroFactura_LostFocus()
    If txtNroFactura.Text <> "" Then
        sql = "SELECT CLI_CODIGO"
        sql = sql & " FROM PAGOS_CLIENTES_V"
        sql = sql & " WHERE FCL_TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
        sql = sql & " AND FCL_NUMERO=" & XN(txtNroFactura)
        If txtCliente.Text <> "" Then sql = sql & " AND CLI_CODIGO=" & XN(txtCliente)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtNroFactura.Text = Format(txtNroFactura.Text, "00000000")
        Else
            MsgBox "La Factura no existe o no corresponde con el Cliente", vbExclamation, TIT_MSGBOX
            txtNroFactura.Text = ""
            txtNroFactura.SetFocus
        End If
        rec.Close
    End If
End Sub
