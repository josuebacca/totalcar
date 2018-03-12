VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCtaCteCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cta-Cte Clientes"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   60
      TabIndex        =   15
      Top             =   6555
      Width           =   7965
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
      Height          =   1020
      Left            =   30
      TabIndex        =   24
      Top             =   6690
      Width           =   5310
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   345
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   600
         Width           =   1590
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   1020
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2085
         TabIndex        =   7
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   225
         TabIndex        =   25
         Top             =   300
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   720
      Left            =   5415
      Picture         =   "frmCtaCteCliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6915
      Width           =   855
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   720
      Left            =   7170
      Picture         =   "frmCtaCteCliente.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6915
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   720
      Left            =   6285
      Picture         =   "frmCtaCteCliente.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6915
      Width           =   870
   End
   Begin VB.Frame frameBuscar 
      Caption         =   "Buscar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   90
      TabIndex        =   11
      Top             =   0
      Width           =   7845
      Begin VB.CommandButton cmdBuscarCliente 
         Height          =   315
         Left            =   2490
         MaskColor       =   &H000000FF&
         Picture         =   "frmCtaCteCliente.frx":0EDE
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Buscar Cliente"
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   0
         Top             =   300
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
         Left            =   2925
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Descripción"
         Top             =   300
         Width           =   4575
      End
      Begin VB.CommandButton CmdBuscAprox 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
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
         Left            =   6375
         MaskColor       =   &H00000000&
         TabIndex        =   4
         ToolTipText     =   "Buscar "
         Top             =   735
         UseMaskColor    =   -1  'True
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   52232193
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   52232193
         CurrentDate     =   41098
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
         Left            =   870
         TabIndex        =   14
         Top             =   345
         Width           =   525
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   390
         TabIndex        =   13
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3180
         TabIndex        =   12
         Top             =   795
         Width           =   960
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdCtaCte 
      Height          =   3840
      Left            =   75
      TabIndex        =   5
      Top             =   1890
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   6773
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   8388736
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   4890
      Top             =   5805
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   4395
      Top             =   5775
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSaldoActual 
      AutoSize        =   -1  'True
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   2595
      TabIndex        =   29
      Top             =   6225
      Width           =   705
   End
   Begin VB.Label lblSaldoActual1 
      AutoSize        =   -1  'True
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   900
      TabIndex        =   28
      Top             =   6225
      Width           =   705
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
      TabIndex        =   27
      Top             =   5865
      Width           =   750
   End
   Begin VB.Line Line1 
      X1              =   5175
      X2              =   7920
      Y1              =   6300
      Y2              =   6300
   End
   Begin VB.Label lblDebe 
      AutoSize        =   -1  'True
      Caption         =   "Debe"
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
      Left            =   5595
      TabIndex        =   23
      Top             =   5760
      Width           =   585
   End
   Begin VB.Label lblHaber 
      AutoSize        =   -1  'True
      Caption         =   "Haber"
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
      Left            =   5520
      TabIndex        =   22
      Top             =   6030
      Width           =   660
   End
   Begin VB.Label lblSal 
      AutoSize        =   -1  'True
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   5550
      TabIndex        =   21
      Top             =   6315
      Width           =   630
   End
   Begin VB.Label lblTotalDebe 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Debe"
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
      Left            =   7020
      TabIndex        =   19
      Top             =   5775
      Width           =   585
   End
   Begin VB.Label lblTotalHaber 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Haber"
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
      Left            =   6945
      TabIndex        =   20
      Top             =   6045
      Width           =   660
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
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
      Left            =   150
      TabIndex        =   18
      Top             =   1635
      Width           =   660
   End
   Begin VB.Label lblCliente 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
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
      Left            =   150
      TabIndex        =   17
      Top             =   1245
      Width           =   735
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   6975
      TabIndex        =   16
      Top             =   6315
      Width           =   630
   End
End
Attribute VB_Name = "frmCtaCteCliente"
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

Private Sub CmdBuscAprox_Click()
    Dim Debe As Double
    Dim Haber As Double
    Dim Saldo As Double
    Saldo = 0
    Debe = 0
    Haber = 0
    
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT CC.*, TC.TCO_ABREVIA"
    sql = sql & " FROM CTA_CTE_CLIENTE CC, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " CLI_CODIGO=" & XN(txtCliente)
    sql = sql & " AND CC.TCO_CODIGO=TC.TCO_CODIGO"
    If Not IsNull(FechaDesde.Value) Then sql = sql & " AND CTA_CTE_FECHA >=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta.Value) Then sql = sql & " AND CTA_CTE_FECHA <=" & XDQ(FechaHasta)
    
    sql = sql & " ORDER BY CTA_CTE_FECHA"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        GrdCtaCte.Rows = 1
        GrdCtaCte.HighLight = flexHighlightAlways
        Do While rec.EOF = False
            
            If rec!CTA_CTE_DH = "D" Then
                Debe = Debe + CDbl(IIf(IsNull(rec!COM_IMP_DEBE), 0, rec!COM_IMP_DEBE))
                Saldo = Saldo + CDbl(IIf(IsNull(rec!COM_IMP_DEBE), 0, rec!COM_IMP_DEBE))
            Else
                Haber = Haber + CDbl(IIf(IsNull(rec!COM_IMP_HABER), 0, rec!COM_IMP_HABER))
                Saldo = Saldo - CDbl(IIf(IsNull(rec!COM_IMP_HABER), 0, rec!COM_IMP_HABER))
            End If
            
            GrdCtaCte.AddItem Trim(rec!TCO_ABREVIA) & Chr(9) & rec!COM_FECHA _
                              & Chr(9) & Format(rec!COM_SUCURSAL, "0000") & "-" & Format(rec!COM_NUMERO, "00000000") _
                              & Chr(9) & Valido_Importe(Chk0(rec!COM_IMP_DEBE)) _
                              & Chr(9) & Valido_Importe(Chk0(rec!COM_IMP_HABER)) _
                              & Chr(9) & Valido_Importe(CStr(Saldo))
            rec.MoveNext
        Loop
            
        lblTotalDebe.Caption = Valido_Importe(CStr(Debe))
        lblTotalHaber.Caption = Valido_Importe(CStr(Haber))
        lblSaldo.Caption = Valido_Importe(CStr(Debe - Haber))
        lblCliente.Caption = "Cta-Cte del Cliente  " & txtDesCli.Text
        If Not IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
            lblFecha.Caption = "Desde  " & FechaDesde.Value & "  al  " & FechaHasta.Value
        ElseIf Not IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
            lblFecha.Caption = "Desde  " & FechaDesde.Value & "  al  " & Date
        ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value <> "" Then
            lblFecha.Caption = "Al  " & FechaHasta.Value
        ElseIf IsNull(FechaDesde.Value) And FechaHasta.Value = Null Then
            lblFecha.Caption = "Al  " & Date
        End If
        rec.Close
        sql = "SELECT SUM(COM_IMP_DEBE) AS DEBE,SUM(COM_IMP_HABER) AS HABER, (SUM(COM_IMP_DEBE) -SUM(COM_IMP_HABER)) AS SALDO "
        sql = sql & " FROM CTA_CTE_CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(txtCliente)
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            lblSaldoActual1.Visible = True
            lblSaldoActual1.Caption = "Saldo Actual:"
            lblSaldoActual.Visible = True
            lblSaldoActual.Caption = Valido_Importe(rec!Saldo)
        End If
        rec.Close
        GrdCtaCte.SetFocus
    Else
        MsgBox "No se encontraron datos", vbExclamation, TIT_MSGBOX
        txtCliente.SetFocus
    End If
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
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
     'Rep.WindowState = crptMaximized 'crptMinimized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    
    lblEstado.Caption = "Buscando Listado..."
    
    Rep.SelectionFormula = ""
    If txtCliente.Text <> "" Then
        Rep.SelectionFormula = "{CTA_CTE_CLIENTE.CLI_CODIGO}=" & txtCliente.Text
    End If
    If Not IsNull(FechaDesde.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {CTA_CTE_CLIENTE.CTA_CTE_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {CTA_CTE_CLIENTE.CTA_CTE_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        End If
    End If
    If Not IsNull(FechaHasta.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {CTA_CTE_CLIENTE.CTA_CTE_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {CTA_CTE_CLIENTE.CTA_CTE_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
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
        Rep.Formulas(1) = "SALDOACTUAL='" & Valido_Importe(lblSaldoActual) & "'"
        
    Rep.WindowTitle = "CTA-CTE de Cliente..."
    Rep.ReportFileName = DRIVE & DirReport & "rptctacteclientes.rpt"
        
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     lblEstado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
    txtCliente.Text = ""
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    lblCliente.Caption = "Cliente"
    lblFecha.Caption = "Fecha"
    lblSaldo.Caption = "0,00"
    lblTotalDebe.Caption = "0,00"
    lblTotalHaber.Caption = "0,00"
    lblSaldoActual1.Visible = False
    lblSaldoActual.Visible = False
    GrdCtaCte.Rows = 1
    GrdCtaCte.HighLight = flexHighlightNever
    txtCliente.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmCtaCteCliente = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)
    CmdBuscAprox.Enabled = False
    cmdListar.Enabled = False
    GrdCtaCte.FormatString = "Tipo Comp.|^Fecha|^Número|>Debe|>Haber|>Saldo"
    GrdCtaCte.ColWidth(0) = 1200 'TIPO COMPROBANTE
    GrdCtaCte.ColWidth(1) = 1200 'FECHA
    GrdCtaCte.ColWidth(2) = 1500 'NUMERO
    GrdCtaCte.ColWidth(3) = 1200 'DEBE
    GrdCtaCte.ColWidth(4) = 1200 'HABER
    GrdCtaCte.ColWidth(5) = 1200 'SALDO
    GrdCtaCte.Rows = 2
    
    lblSaldo.Caption = "0,00"
    lblTotalDebe.Caption = "0,00"
    lblTotalHaber.Caption = "0,00"
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""
    lblSaldoActual1.Visible = False
    lblSaldoActual.Visible = False
End Sub
Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
        CmdBuscAprox.Enabled = False
        cmdListar.Enabled = False
        GrdCtaCte.Rows = 1
    Else
        CmdBuscAprox.Enabled = True
        cmdListar.Enabled = True
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
        CmdBuscAprox.Enabled = False
        cmdListar.Enabled = False
        GrdCtaCte.Rows = 1
    Else
        CmdBuscAprox.Enabled = True
        cmdListar.Enabled = True
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


