VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListadoFacturasPendientePorCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Facturas Pendientes de Pago por Cliente"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   60
      TabIndex        =   17
      Top             =   -15
      Width           =   7050
      Begin VB.OptionButton optTodas 
         Caption         =   "Todas"
         Height          =   195
         Left            =   4920
         TabIndex        =   7
         Top             =   1680
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optMaq 
         Caption         =   "de Maquinarias"
         Height          =   195
         Left            =   3000
         TabIndex        =   6
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optRep 
         Caption         =   "de Repuestos"
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   1155
         MaxLength       =   40
         TabIndex        =   0
         Top             =   420
         Width           =   975
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
         Left            =   2595
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Descripción"
         Top             =   420
         Width           =   4350
      End
      Begin VB.CommandButton cmdBuscarCliente 
         Height          =   315
         Left            =   2160
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoFacturasPendientePorCliente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Buscar Cliente"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   1185
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   795
         Width           =   3165
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   16777217
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   4200
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   16777217
         CurrentDate     =   41098
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   105
         TabIndex        =   22
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3135
         TabIndex        =   21
         Top             =   1215
         Width           =   960
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
         Left            =   540
         TabIndex        =   20
         Top             =   465
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   375
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
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
      TabIndex        =   14
      Top             =   2130
      Width           =   7050
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   9
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   8
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   4500
      Picture         =   "frmListadoFacturasPendientePorCliente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2955
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   6255
      Picture         =   "frmListadoFacturasPendientePorCliente.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2955
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoFacturasPendientePorCliente.frx":0EDE
      Height          =   750
      Left            =   5370
      Picture         =   "frmListadoFacturasPendientePorCliente.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2955
      Width           =   870
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3330
      Top             =   3060
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2835
      Top             =   2970
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
      TabIndex        =   16
      Top             =   3135
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoFacturasPendientePorCliente"
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
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.SelectionFormula = ""
    
    If txtCliente.Text <> "" Then
        Rep.SelectionFormula = "{SALDO_FACTURAS_CLIENTE_V.CLI_CODIGO}=" & txtCliente.Text
    End If
    
    If cboVendedor.List(cboVendedor.ListIndex) <> "<Todos>" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = "{SALDO_FACTURAS_CLIENTE_V.VEN_CODIGO}=" & XN(cboVendedor.ItemData(cboVendedor.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & "{SALDO_FACTURAS_CLIENTE_V.VEN_CODIGO}=" & XN(cboVendedor.ItemData(cboVendedor.ListIndex))
        End If
    End If
    
    If optMaq.Value = True Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = "{SALDO_FACTURAS_CLIENTE_V.FCL_IVA}<>21 "
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {SALDO_FACTURAS_CLIENTE_V.FCL_IVA}<>21 "
        End If
    End If
    If optRep.Value = True Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = "{SALDO_FACTURAS_CLIENTE_V.FCL_IVA}=21 "
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {SALDO_FACTURAS_CLIENTE_V.FCL_IVA}=21 "
        End If
    End If
        
    If Not IsNull(FechaDesde.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {SALDO_FACTURAS_CLIENTE_V.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {SALDO_FACTURAS_CLIENTE_V.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        End If
    End If
    If Not IsNull(FechaHasta.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {SALDO_FACTURAS_CLIENTE_V.FCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                           
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {SALDO_FACTURAS_CLIENTE_V.FCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        End If
    End If
    
'    If Rep.SelectionFormula = "" Then 'ESTADO DEFINITIVO
'        Rep.SelectionFormula = " {FACTURA_CLIENTE.EST_CODIGO}=3"
'    Else
'        Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.EST_CODIGO}=3"
'    End If
    
    If Not IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf Not IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value = Null And FechaHasta.Value = Null Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    
    
    Rep.WindowTitle = "Listado de Facturas Pendientes de pago por Cliente"
    Rep.ReportFileName = DRIVE & DirReport & "listadofacturaspendientespagoxcliente.rpt"
    
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
    txtDesCli.Text = ""
    cboVendedor.ListIndex = 0
    txtCliente.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoFacturasPendientePorCliente = Nothing
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
    lblEstado.Caption = ""
    
    CargoComboVendedor
End Sub
Private Sub CargoComboVendedor()
    sql = "SELECT VEN_CODIGO,VEN_NOMBRE"
    sql = sql & " FROM VENDEDOR ORDER BY VEN_NOMBRE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboVendedor.AddItem "<Todos>"
        Do While rec.EOF = False
            cboVendedor.AddItem rec!VEN_NOMBRE
            cboVendedor.ItemData(cboVendedor.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        cboVendedor.ListIndex = 0
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

