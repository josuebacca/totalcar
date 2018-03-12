VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListadoCantidadesVendidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Cantidades Vendidas"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6555
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
      Height          =   735
      Left            =   60
      TabIndex        =   16
      Top             =   2145
      Width           =   6435
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
         TabIndex        =   17
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   3885
      Picture         =   "frmListadoCantidadesVendidas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2925
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   5640
      Picture         =   "frmListadoCantidadesVendidas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2925
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoCantidadesVendidas.frx":0BD4
      Height          =   750
      Left            =   4755
      Picture         =   "frmListadoCantidadesVendidas.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2925
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
      Height          =   2130
      Left            =   60
      TabIndex        =   11
      Top             =   -15
      Width           =   6435
      Begin VB.ComboBox cboRep 
         Height          =   315
         Left            =   1635
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   4020
      End
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   1635
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   780
         Width           =   4020
      End
      Begin VB.ComboBox cboLinea 
         Height          =   315
         Left            =   1635
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   4020
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1635
         TabIndex        =   3
         Top             =   1680
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
         Left            =   4200
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   58982401
         CurrentDate     =   41098
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Rubro:"
         Height          =   195
         Left            =   1110
         TabIndex        =   15
         Top             =   825
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Línea:"
         Height          =   195
         Left            =   1125
         TabIndex        =   14
         Top             =   345
         Width           =   465
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3165
         TabIndex        =   13
         Top             =   1770
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   585
         TabIndex        =   12
         Top             =   1770
         Width           =   1005
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3330
      Top             =   3030
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2850
      Top             =   2940
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
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoCantidadesVendidas"
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

Private Sub cboLinea_Click()
    cborubro.Clear
End Sub

Private Sub cboLinea_LostFocus()
    If cbolinea.List(cbolinea.ListIndex) <> "<Todas>" Then
        cargocboRubro
    Else
        cborubro.Clear
        cborubro.AddItem "<Todos>"
        cborubro.ListIndex = 0
    End If
End Sub

Private Sub cmdListar_Click()
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    
   
    Rep.SelectionFormula = ""
    If cbolinea.List(cbolinea.ListIndex) <> "<Todas>" Then
        Rep.SelectionFormula = " {LINEAS.LNA_CODIGO}=" & XN(cbolinea.ItemData(cbolinea.ListIndex))
    End If
    If cborubro.List(cborubro.ListIndex) <> "<Todos>" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {RUBROS.RUB_CODIGO}=" & XN(cborubro.ItemData(cborubro.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {RUBROS.RUB_CODIGO}=" & XN(cborubro.ItemData(cborubro.ListIndex))
        End If
    End If
    If cboRep.List(cboRep.ListIndex) <> "<Todos>" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {TIPO_PRESENTACION.TPRE_CODIGO}=" & XN(cboRep.ItemData(cboRep.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {TIPO_PRESENTACION.TPRE_CODIGO}=" & XN(cboRep.ItemData(cboRep.ListIndex))
        End If
    End If
    
    If Not IsNull(FechaDesde.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {DETALLE_FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {DETALLE_FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        End If
    End If
    If Not IsNull(FechaHasta.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {DETALLE_FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                           
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {DETALLE_FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        End If
    End If
    
    If Rep.SelectionFormula = "" Then 'ESTADO DEFINITIVO
        Rep.SelectionFormula = " {FACTURA_CLIENTE.EST_CODIGO}=3"
    Else
        Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.EST_CODIGO}=3"
    End If
    
    If Not IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf Not IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value = Null And FechaHasta.Value = Null Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    Rep.WindowTitle = "Listado de Cantidades Vendidas"
    Rep.ReportFileName = DRIVE & DirReport & "rptlistadocantidadesvendidas.rpt"

    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     lblEstado.Caption = ""
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
End Sub

Private Sub CmdNuevo_Click()
    cbolinea.ListIndex = 0
    cborubro.Clear
    FechaDesde.Value = Null
    FechaHasta.Value = Null
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoCantidadesVendidas = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    cargocboLinea
    cborubro.AddItem "<Todos>"
    cborubro.ListIndex = 0
    cargocboRepres -1, -1
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub
Function cargocboRepres(codL As Integer, codR As Integer)
    cboRep.Clear
    sql = "SELECT * FROM TIPO_PRESENTACION WHERE TPRE_CODIGO <> 0 "
    If codL <> -1 Then
        sql = sql & " AND LNA_CODIGO = " & cbolinea.ItemData(cbolinea.ListIndex) & ""
    End If
    If codR <> -1 Then
        sql = sql & "AND RUB_CODIGO = " & cborubro.ItemData(cborubro.ListIndex) & ""
    End If
    sql = sql & " ORDER BY TPRE_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboRep.AddItem rec!TPRE_DESCRI
            cboRep.ItemData(cboRep.NewIndex) = rec!TPRE_CODIGO
            rec.MoveNext
        Loop
        cboRep.ListIndex = -1
    End If
    rec.Close
End Function

Private Sub cargocboLinea()
    lblEstado.Caption = ""
    sql = "SELECT * FROM LINEAS  ORDER BY LNA_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cbolinea.AddItem "<Todas>"
        Do While rec.EOF = False
            cbolinea.AddItem rec!LNA_DESCRI
            cbolinea.ItemData(cbolinea.NewIndex) = rec!LNA_CODIGO
            rec.MoveNext
        Loop
        cbolinea.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub cargocboRubro()
    sql = "SELECT * FROM RUBROS "
    sql = sql & " WHERE LNA_CODIGO= " & cbolinea.ItemData(cbolinea.ListIndex)
    sql = sql & " ORDER BY RUB_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cborubro.AddItem "<Todos>"
        Do While rec.EOF = False
            cborubro.AddItem rec!RUB_DESCRI
            cborubro.ItemData(cborubro.NewIndex) = rec!RUB_CODIGO
            rec.MoveNext
        Loop
        cborubro.ListIndex = 0
    End If
    rec.Close
End Sub


