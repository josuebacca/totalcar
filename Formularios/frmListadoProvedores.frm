VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListadoProvedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Proveedores...."
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoProvedores.frx":0000
      Height          =   705
      Left            =   6225
      Picture         =   "frmListadoProvedores.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5940
      Width           =   810
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   4140
      Top             =   5955
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
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
      Height          =   1110
      Left            =   45
      TabIndex        =   18
      Top             =   4785
      Width           =   7815
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   660
         Width           =   1665
      End
      Begin VB.OptionButton optExcel 
         Caption         =   "Excel"
         Height          =   225
         Left            =   4020
         TabIndex        =   10
         Top             =   330
         Width           =   780
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   8
         Top             =   330
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   9
         Top             =   330
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
         Left            =   1920
         TabIndex        =   20
         Top             =   780
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   330
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ver..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   45
      TabIndex        =   17
      Top             =   4200
      Width           =   7830
      Begin VB.OptionButton optGeneralTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado General (Todos)"
         Height          =   210
         Left            =   4935
         TabIndex        =   22
         Top             =   255
         Width           =   2280
      End
      Begin VB.OptionButton optDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado"
         Height          =   255
         Left            =   345
         TabIndex        =   7
         Top             =   255
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.OptionButton optGeneral 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado General"
         Height          =   210
         Left            =   2677
         TabIndex        =   6
         Top             =   255
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   705
      Left            =   5400
      Picture         =   "frmListadoProvedores.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5940
      Width           =   810
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   705
      Left            =   7050
      Picture         =   "frmListadoProvedores.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5940
      Width           =   810
   End
   Begin VB.Frame Frame2 
      Caption         =   "   Proveedor ......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   60
      TabIndex        =   13
      Top             =   90
      Width           =   7815
      Begin VB.CheckBox chkRazSoc 
         Alignment       =   1  'Right Justify
         Caption         =   "  Por Raz. Soc."
         Height          =   195
         Left            =   75
         TabIndex        =   1
         Top             =   750
         Width           =   1395
      End
      Begin VB.CheckBox chkPorTipo 
         Alignment       =   1  'Right Justify
         Caption         =   "   Por Tipo"
         Height          =   195
         Left            =   345
         TabIndex        =   0
         Top             =   465
         Width           =   1125
      End
      Begin VB.ComboBox cboBuscaTipoProv 
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   390
         Width           =   4095
      End
      Begin VB.TextBox txtBuscaProv 
         Height          =   315
         Left            =   2505
         MaxLength       =   30
         TabIndex        =   3
         Top             =   765
         Width           =   4095
      End
      Begin VB.CommandButton cmdBusProv 
         Caption         =   "B&uscar"
         Height          =   750
         Left            =   6735
         Picture         =   "frmListadoProvedores.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   225
         Index           =   28
         Left            =   2085
         TabIndex        =   15
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         Height          =   195
         Index           =   26
         Left            =   1665
         TabIndex        =   14
         Top             =   780
         Width           =   780
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgBuscaProv 
      Height          =   2745
      Left            =   45
      TabIndex        =   5
      Top             =   1395
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   4842
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorSel    =   8388736
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3690
      Top             =   5970
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
      Left            =   165
      TabIndex        =   16
      Top             =   6165
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoProvedores"
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

Private Sub chkPorTipo_Click()
    If chkPorTipo.Value = Checked Then
        cboBuscaTipoProv.Enabled = True
    Else
        cboBuscaTipoProv.Enabled = False
    End If
End Sub

Private Sub chkRazSoc_Click()
    If chkRazSoc.Value = Checked Then
        txtBuscaProv.Enabled = True
    Else
        txtBuscaProv.Enabled = False
    End If
End Sub

Private Sub cmdBusProv_Click()
    Screen.MousePointer = vbHourglass
    Me.lblEstado.Caption = "Buscando Proveedores...."
    Me.Refresh
    
    fgBuscaProv.Rows = 1
    sql = "SELECT TP.TPR_CODIGO,TP.TPR_DESCRI,"
    sql = sql & "P.PROV_CODIGO,P.PROV_RAZSOC"
    sql = sql & " FROM TIPO_PROVEEDOR TP, PROVEEDOR P"
    sql = sql & " WHERE TP.TPR_CODIGO=P.TPR_CODIGO"
   If chkPorTipo.Value = Checked Then
    sql = sql & " AND TP.TPR_CODIGO=" & XN(cboBuscaTipoProv.ItemData(cboBuscaTipoProv.ListIndex))
   End If
   If chkRazSoc.Value = Checked Then
    sql = sql & " AND P.PROV_RAZSOC LIKE '" & Trim(txtBuscaProv.Text) & "%'"
   End If
    sql = sql & " ORDER BY P.TPR_CODIGO,P.PROV_CODIGO,P.PROV_RAZSOC"
    
   rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
   If rec.EOF = False Then
        fgBuscaProv.HighLight = flexHighlightAlways
        rec.MoveFirst
        ' SI BUSCO POR RAZ SOCIAL
        Do While rec.EOF = False
         fgBuscaProv.AddItem rec!TPR_DESCRI & Chr(9) & rec!PROV_CODIGO _
                     & Chr(9) & rec!PROV_RAZSOC & Chr(9) & rec!TPR_CODIGO
                     
         rec.MoveNext
        Loop
         
        fgBuscaProv.SetFocus
        Screen.MousePointer = vbNormal
        Me.lblEstado.Caption = ""
    
   Else ' SI NO ENCONTRO NINGUNO
        Screen.MousePointer = vbNormal
        Me.lblEstado.Caption = ""
        fgBuscaProv.HighLight = flexHighlightNever
        fgBuscaProv.Rows = 1
        MsgBox "No se han encontrado Proveedores", vbExclamation, TIT_MSGBOX
        cboBuscaTipoProv.SetFocus
   End If
   rec.Close
End Sub

Private Sub cmdListar_Click()
    lblEstado.Caption = "Buscando Listado..."
     'Rep.WindowState = crptMaximized 'crptMinimized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""

    If optGeneral.Value = True Then 'LISTADO GENERAL DE DE UN TIPO DE PROVEEDOR SELECCIONADO
        
        If chkPorTipo.Value = Checked Then
            If chkPorTipo.Value = Checked And (chkRazSoc.Value = Unchecked) Then
                Rep.SelectionFormula = "{PROVEEDOR.TPR_CODIGO}=" & cboBuscaTipoProv.ItemData(cboBuscaTipoProv.ListIndex)
            ElseIf chkRazSoc.Value = Checked Then
                Rep.SelectionFormula = "{PROVEEDOR.TPR_CODIGO}=" & fgBuscaProv.TextMatrix(fgBuscaProv.RowSel, 3) _
                                       & " AND {PROVEEDOR.PROV_CODIGO}=" & fgBuscaProv.TextMatrix(fgBuscaProv.RowSel, 1)
            End If
            Rep.WindowTitle = "Maestro de Proveedores..."
            Rep.ReportFileName = DRIVE & DirReport & "MaestroProveedores.rpt"
        Else
            MsgBox "Debe seleccionar un Tipo de Proveedor", vbExclamation, TIT_MSGBOX
            lblEstado.Caption = ""
            chkPorTipo.SetFocus
            Exit Sub
        End If
            
    ElseIf optGeneralTodos.Value = True Then 'LISTADO GENERAL DE TODOS LOS TIPO DE PROVEEDOR
        
        Rep.SelectionFormula = ""
        Rep.WindowTitle = "Maestro de Proveedores..."
        Rep.ReportFileName = DRIVE & DirReport & "MaestroProveedores.rpt"
    
    ElseIf optDetallado.Value = True Then 'LISTADO DETALLADO DE UNPROVEEDOR
        
        If fgBuscaProv.Rows > 1 Then
            If fgBuscaProv.TextMatrix(fgBuscaProv.RowSel, 2) <> "" Then
                Rep.SelectionFormula = ""
                Rep.SelectionFormula = "{PROVEEDOR.TPR_CODIGO}=" & fgBuscaProv.TextMatrix(fgBuscaProv.RowSel, 3) _
                                          & " AND {PROVEEDOR.PROV_CODIGO}=" & fgBuscaProv.TextMatrix(fgBuscaProv.RowSel, 1)
                Rep.WindowTitle = "Maestro de Proveedores - Detallado..."
                Rep.ReportFileName = DRIVE & DirReport & "maestroproveedoresDetalle.rpt"
            Else
                Rep.SelectionFormula = ""
                Rep.Formulas(0) = ""
                lblEstado.Caption = ""
                Exit Sub
            End If
        Else
            MsgBox "Debe seleccionar un Proveedor", vbExclamation, TIT_MSGBOX
            lblEstado.Caption = ""
            chkPorTipo.SetFocus
            Exit Sub
        End If
        
    End If
    
    If optPantalla.Value = True Then
         Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
        Rep.PrintFileType = crptExcel50
    ElseIf optExcel.Value = True Then
        Rep.Destination = crptToFile
        Rep.PrintFileType = crptExcel50
    End If
    Rep.Action = 1
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    lblEstado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
    fgBuscaProv.Rows = 1
    fgBuscaProv.HighLight = flexHighlightNever
    txtBuscaProv.Text = ""
    cboBuscaTipoProv.ListIndex = 0
    chkPorTipo.Value = Unchecked
    chkRazSoc.Value = Unchecked
    chkPorTipo.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoProvedores = Nothing
    Unload Me
End Sub

Private Sub fgBuscaProv_DblClick()
    cmdListar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    lblEstado.Caption = ""
    fgBuscaProv.Clear
    fgBuscaProv.Rows = 2
    txtBuscaProv.Text = ""
    
    fgBuscaProv.FormatString = "Tipo Prov.|Nro. Proveedor|Proveedor|codigo tipo_proveedor"

    fgBuscaProv.ColWidth(0) = 2500
    fgBuscaProv.ColWidth(1) = 1200
    fgBuscaProv.ColWidth(2) = 4000
    fgBuscaProv.ColWidth(3) = 0
    
    CargoComboTipoProveedor
    'impresora actual
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    cboBuscaTipoProv.Enabled = False
    txtBuscaProv.Enabled = False
End Sub
Public Sub CargoComboTipoProveedor()
    'Cargo el combo Tipo de Proveedor
    cboBuscaTipoProv.Clear
    
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            cboBuscaTipoProv.AddItem rec.Fields!TPR_DESCRI
            cboBuscaTipoProv.ItemData(cboBuscaTipoProv.NewIndex) = rec.Fields!TPR_CODIGO
            rec.MoveNext
        Loop
        cboBuscaTipoProv.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub txtBuscaProv_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
