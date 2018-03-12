VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListadoClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Clientes"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
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
      TabIndex        =   9
      Top             =   1515
      Width           =   6690
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   3
         Top             =   330
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   2
         Top             =   330
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   330
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   4140
      Picture         =   "frmListadoClientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2385
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   5895
      Picture         =   "frmListadoClientes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2385
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoClientes.frx":0BD4
      Height          =   750
      Left            =   5010
      Picture         =   "frmListadoClientes.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2385
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
      Height          =   930
      Left            =   60
      TabIndex        =   8
      Top             =   -15
      Width           =   6690
      Begin VB.CommandButton cmdBuscarCliente 
         Height          =   315
         Left            =   1800
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoClientes.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Buscar Cliente"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   405
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
         Top             =   360
         Width           =   4305
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   795
         MaxLength       =   40
         TabIndex        =   0
         Top             =   360
         Width           =   975
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
         TabIndex        =   11
         Top             =   405
         Width           =   525
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3330
      Top             =   2490
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2775
      Top             =   2490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameVer 
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
      Height          =   585
      Left            =   60
      TabIndex        =   14
      Top             =   900
      Width           =   6705
      Begin VB.CheckBox chkBaja 
         Caption         =   "Dados de Baja "
         Height          =   240
         Left            =   4950
         TabIndex        =   17
         Top             =   255
         Width           =   1395
      End
      Begin VB.OptionButton optDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado"
         Height          =   255
         Left            =   210
         TabIndex        =   16
         Top             =   255
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton optGeneralTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado General"
         Height          =   210
         Left            =   2535
         TabIndex        =   15
         Top             =   270
         Width           =   1620
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
      Height          =   270
      Left            =   150
      TabIndex        =   12
      Top             =   2565
      Width           =   810
   End
End
Attribute VB_Name = "frmListadoClientes"
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
    
    If optDetallado.Value = True Then
        If txtCliente.Text <> "" Then
           Rep.SelectionFormula = "{CLIENTE.CLI_CODIGO}=" & txtCliente.Text
           If chkBaja.Value = Checked Then
               Rep.SelectionFormula = Rep.SelectionFormula & " AND {CLIENTE.CLI_ESTADO}=2"
           Else
               Rep.SelectionFormula = Rep.SelectionFormula & " AND {CLIENTE.CLI_ESTADO}=1"
           End If
        Else
           lblEstado.Caption = ""
           MsgBox "Debe seleccionar un Cliente", vbExclamation, TIT_MSGBOX
           txtCliente.SetFocus
           Exit Sub
        End If
        Rep.WindowTitle = "Listado de Clientes - Detallado"
        Rep.ReportFileName = DRIVE & DirReport & "rptclientes.rpt"
    End If
    
    If optGeneralTodos.Value = True Then
        If chkBaja.Value = Checked Then
             Rep.SelectionFormula = "{CLIENTE.CLI_ESTADO}=2"
        Else
             Rep.SelectionFormula = "{CLIENTE.CLI_ESTADO}=1"
        End If
        Rep.WindowTitle = "Listado de Clientes - General"
        Rep.ReportFileName = DRIVE & DirReport & "rptclientesgeneral.rpt"
    End If
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
    Rep.Action = 1
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    lblEstado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    chkBaja.Value = Unchecked
    txtCliente.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoClientes = Nothing
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

