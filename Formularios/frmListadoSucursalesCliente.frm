VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListadoSucursalesCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Sucursales por Cliente"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
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
      Height          =   810
      Left            =   60
      TabIndex        =   11
      Top             =   1980
      Width           =   6690
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   5
         Top             =   330
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   4
         Top             =   330
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   330
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   780
      Left            =   4140
      Picture         =   "frmListadoSucursalesCliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2850
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   780
      Left            =   5895
      Picture         =   "frmListadoSucursalesCliente.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2850
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoSucursalesCliente.frx":0BD4
      Height          =   780
      Left            =   5010
      Picture         =   "frmListadoSucursalesCliente.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2850
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
      Height          =   1380
      Left            =   60
      TabIndex        =   10
      Top             =   -30
      Width           =   6690
      Begin VB.TextBox txtDesSuc 
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
         Left            =   2250
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "Descripción"
         Top             =   780
         Width           =   4305
      End
      Begin VB.CommandButton cmdBuscarSucursal 
         Height          =   315
         Left            =   1815
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoSucursalesCliente.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Buscar Cliente"
         Top             =   780
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   300
         Left            =   810
         MaxLength       =   40
         TabIndex        =   2
         Top             =   780
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscarCliente 
         Height          =   315
         Left            =   1815
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoSucursalesCliente.frx":14F2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Buscar Cliente"
         Top             =   405
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
         Left            =   2250
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Descripción"
         Top             =   405
         Width           =   4305
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   810
         MaxLength       =   40
         TabIndex        =   0
         Top             =   405
         Width           =   975
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   825
         Width           =   660
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
         Left            =   225
         TabIndex        =   13
         Top             =   450
         Width           =   525
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3330
      Top             =   2955
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2775
      Top             =   2955
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
      Height          =   615
      Left            =   60
      TabIndex        =   16
      Top             =   1365
      Width           =   6705
      Begin VB.CheckBox chkBaja 
         Caption         =   "Dados de Baja "
         Height          =   240
         Left            =   5025
         TabIndex        =   21
         Top             =   270
         Width           =   1395
      End
      Begin VB.OptionButton optDetalle 
         Caption         =   "Sucursal (Detallado)"
         Height          =   240
         Left            =   2760
         TabIndex        =   20
         Top             =   270
         Width           =   1815
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Sucursales por Cliente"
         Height          =   270
         Left            =   360
         TabIndex        =   19
         Top             =   255
         Width           =   1905
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
      Height          =   300
      Left            =   150
      TabIndex        =   14
      Top             =   3015
      Width           =   810
   End
End
Attribute VB_Name = "frmListadoSucursalesCliente"
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

Private Sub cmdBuscarSucursal_Click()
    frmBuscar.TipoBusqueda = 3
    frmBuscar.TxtDescriB = ""
    If txtCliente.Text <> "" Then
        frmBuscar.CodigoCli = txtCliente.Text
    Else
        frmBuscar.CodigoCli = ""
    End If
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 3
        txtCliente.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 0
        txtcodigo.Text = frmBuscar.grdBuscar.Text
        txtcodigo.SetFocus
        TxtCodigo_LostFocus
    Else
        txtcodigo.SetFocus
    End If
End Sub

Private Sub cmdListar_Click()
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    
    If OptTodos.Value = True Then
        If txtCliente.Text <> "" Then
           Rep.SelectionFormula = "{SUCURSAL.CLI_CODIGO}=" & txtCliente.Text
           If chkBaja.Value = Unchecked Then
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {CLIENTE.CLI_ESTADO}=1"
           Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {CLIENTE.CLI_ESTADO}=2"
           End If
        Else
           MsgBox "Debe seleccionar un Cliente", vbExclamation, TIT_MSGBOX
           txtCliente.SetFocus
           Exit Sub
        End If
        Rep.ReportFileName = DRIVE & DirReport & "rptsucursalxcliente.rpt"
        
    ElseIf optDetalle.Value = True Then
        If txtCliente.Text <> "" Then
           Rep.SelectionFormula = "{SUCURSAL.CLI_CODIGO}=" & txtCliente.Text
           If chkBaja.Value = Unchecked Then
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {CLIENTE.CLI_ESTADO}=1"
           Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {CLIENTE.CLI_ESTADO}=2"
           End If
        Else
           MsgBox "Debe seleccionar un Cliente", vbExclamation, TIT_MSGBOX
           txtCliente.SetFocus
           Exit Sub
        End If
        If txtcodigo.Text <> "" Then
           Rep.SelectionFormula = Rep.SelectionFormula & " AND {SUCURSAL.SUC_CODIGO}=" & txtcodigo.Text
           If chkBaja.Value = Unchecked Then
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {CLIENTE.CLI_ESTADO}=1"
           Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {CLIENTE.CLI_ESTADO}=2"
           End If
        Else
           MsgBox "Debe seleccionar una Sucursal", vbExclamation, TIT_MSGBOX
           txtcodigo.SetFocus
           Exit Sub
        End If
        Rep.ReportFileName = DRIVE & DirReport & "rptsucursal.rpt"
    Else
        Exit Sub
    End If

     Rep.Destination = crptToWindow
     Rep.Action = 1
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     lblEstado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    txtcodigo.Text = ""
    txtDesSuc.Text = ""
    OptTodos.Value = True
    chkBaja.Value = Unchecked
    txtCliente.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoSucursalesCliente = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    OptTodos.Value = True
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""
End Sub

Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
        txtcodigo.Text = ""
        txtDesSuc.Text = ""
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

Private Sub TxtCodigo_Change()
    If txtcodigo.Text = "" Then
        txtCliente.Text = ""
        txtDesCli.Text = ""
        txtcodigo.Text = ""
        txtDesSuc.Text = ""
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto txtcodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If txtcodigo.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, SUC_DESCRI FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(txtcodigo)
        If txtCliente.Text <> "" Then
         sql = sql & " AND CLI_CODIGO=" & XN(txtCliente)
        End If
        lblEstado.Caption = "Buscando..."
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtCliente.Text = Rec1!CLI_CODIGO
            txtCliente_LostFocus
            txtDesSuc.Text = Rec1!SUC_DESCRI
            lblEstado.Caption = ""
        Else
            lblEstado.Caption = ""
            MsgBox "La Sucursal no existe", vbExclamation, TIT_MSGBOX
            txtDesSuc.Text = ""
            txtcodigo.SetFocus
             Rec1.Close
            Exit Sub
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtDesCli_Change()
    If txtDesCli.Text = "" Then
        txtCliente.Text = ""
        txtcodigo.Text = ""
        txtDesSuc.Text = ""
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

Private Function BuscoSucursal(Cli As String, Suc As String) As String
    sql = "SELECT C.CLI_CODIGO, C.CLI_RAZSOC, S.SUC_CODIGO, S.SUC_DESCRI"
    sql = sql & " FROM CLIENTE C, SUCURSAL S"
    sql = sql & " WHERE "
    If txtcodigo.Text <> "" Then
        sql = sql & " S.SUC_CODIGO=" & XN(Suc)
    Else
        sql = sql & " S.SUC_DESCRI LIKE '" & Suc & "%'"
    End If
    If txtCliente.Text <> "" Then
        sql = sql & " AND C.CLI_CODIGO=" & XN(Cli)
    End If
    sql = sql & " AND S.CLI_CODIGO=C.CLI_CODIGO"
    BuscoSucursal = sql
End Function

Private Sub txtDesSuc_Change()
    If txtDesCli.Text = "" Then
        txtCliente.Text = ""
        txtcodigo.Text = ""
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtDesSuc_GotFocus()
     SelecTexto txtDesSuc
End Sub

Private Sub txtDesSuc_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesSuc_LostFocus()
    If txtcodigo.Text = "" And txtDesSuc.Text <> "" Then
        Rec1.Open BuscoSucursal(txtCliente, txtDesSuc), DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            If Rec1.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 3
                frmBuscar.CodigoCli = ""
                frmBuscar.TxtDescriB.Text = txtDesSuc.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 3
                    txtCliente.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 0
                    txtcodigo.Text = frmBuscar.grdBuscar.Text
                    txtcodigo.SetFocus
                    TxtCodigo_LostFocus
                Else
                    txtCliente.SetFocus
                End If
            Else
                txtCliente.Text = Rec1!CLI_CODIGO
                txtDesCli.Text = Rec1!CLI_RAZSOC
                txtcodigo.Text = Rec1!SUC_CODIGO
                txtDesSuc.Text = Rec1!SUC_DESCRI
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        Rec1.Close
    End If
End Sub
