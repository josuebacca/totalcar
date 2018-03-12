VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCtaCteProveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cta-Cte Proveedores"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   60
      TabIndex        =   13
      Top             =   6555
      Width           =   8100
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
      Height          =   825
      Left            =   30
      TabIndex        =   22
      Top             =   6705
      Width           =   5190
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   495
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   225
         Width           =   1335
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   1140
         TabIndex        =   24
         Top             =   405
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2205
         TabIndex        =   23
         Top             =   405
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   345
         TabIndex        =   25
         Top             =   390
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   720
      Left            =   5430
      Picture         =   "frmCtaCteProveedores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6810
      Width           =   855
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   720
      Left            =   7185
      Picture         =   "frmCtaCteProveedores.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6810
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   720
      Left            =   6300
      Picture         =   "frmCtaCteProveedores.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6810
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
      Height          =   1455
      Left            =   105
      TabIndex        =   10
      Top             =   45
      Width           =   8055
      Begin VB.ComboBox cboTipoProveedor 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   285
         Width           =   4215
      End
      Begin VB.TextBox txtCodProveedor 
         Height          =   300
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   1
         Top             =   652
         Width           =   1455
      End
      Begin VB.TextBox txtProvRazSoc 
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
         Left            =   2715
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Descripción"
         Top             =   645
         Width           =   5040
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
         Left            =   6600
         MaskColor       =   &H00000000&
         TabIndex        =   5
         ToolTipText     =   "Buscar "
         Top             =   1005
         UseMaskColor    =   -1  'True
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   16842753
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   3960
         TabIndex        =   4
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   16842753
         CurrentDate     =   41098
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Prov.:"
         Height          =   195
         Left            =   360
         TabIndex        =   29
         Top             =   315
         Width           =   780
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   28
         Top             =   660
         Width           =   540
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   2925
         TabIndex        =   11
         Top             =   1050
         Width           =   960
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdCtaCte 
      Height          =   3480
      Left            =   75
      TabIndex        =   6
      Top             =   2250
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   6138
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
      Left            =   5295
      Top             =   5730
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   4800
      Top             =   5700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   870
      TabIndex        =   31
      Top             =   6225
      Width           =   705
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
      Left            =   2565
      TabIndex        =   30
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
      Left            =   150
      TabIndex        =   27
      Top             =   5805
      Width           =   750
   End
   Begin VB.Line Line1 
      X1              =   5445
      X2              =   8190
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
      Left            =   5865
      TabIndex        =   21
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
      Left            =   5790
      TabIndex        =   20
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
      Left            =   5820
      TabIndex        =   19
      Top             =   6345
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
      Left            =   7290
      TabIndex        =   17
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
      Left            =   7215
      TabIndex        =   18
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
      TabIndex        =   16
      Top             =   1935
      Width           =   660
   End
   Begin VB.Label lblProveedor 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
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
      TabIndex        =   15
      Top             =   1590
      Width           =   1110
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
      Left            =   7245
      TabIndex        =   14
      Top             =   6360
      Width           =   630
   End
End
Attribute VB_Name = "frmCtaCteProveedores"
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
    
    sql = "SELECT CC.*, TC.TCO_DESCRI"
    sql = sql & " FROM CTA_CTE_PROVEEDORES CC, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
    sql = sql & " AND CC.TCO_CODIGO=TC.TCO_CODIGO"
    If Not IsNull(FechaDesde.Value) Then sql = sql & " AND CTA_CTE_FECHA >=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta.Value) Then sql = sql & " AND CTA_CTE_FECHA <=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY COM_FECHA"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        GrdCtaCte.Rows = 1
        GrdCtaCte.HighLight = flexHighlightAlways
        Do While rec.EOF = False
            
            If rec!CTA_CTE_DH = "D" Then
                Debe = Debe + CDbl(rec!COM_IMP_DEBE)
                Saldo = Saldo + CDbl(rec!COM_IMP_DEBE)
            Else
                Haber = Haber + CDbl(rec!COM_IMP_HABER)
                Saldo = Saldo - CDbl(rec!COM_IMP_HABER)
            End If
    
            GrdCtaCte.AddItem rec!TCO_DESCRI & Chr(9) & rec!COM_FECHA & Chr(9) & _
                              rec!COM_SUCURSAL & "-" & rec!COM_NUMERO & Chr(9) & Valido_Importe(rec!COM_IMP_DEBE) _
                              & Chr(9) & Valido_Importe(rec!COM_IMP_HABER) & Chr(9) & Valido_Importe(CStr(Saldo))
            rec.MoveNext
        Loop
        
        lblTotalDebe.Caption = Valido_Importe(CStr(Debe))
        lblTotalHaber.Caption = Valido_Importe(CStr(Haber))
        lblSaldo.Caption = Valido_Importe(CStr(Debe - Haber))
        lblProveedor.Caption = "Cta-Cte del Proveedor  " & txtProvRazSoc.Text
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
        sql = sql & " From CTA_CTE_PROVEEDORES"
        sql = sql & " WHERE TPR_CODIGO=" & XN(cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex))
        sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            lblSaldoActual1.Visible = True
            lblSaldoActual1.Caption = "Saldo Actual:"
            lblSaldoActual.Visible = True
            lblSaldoActual.Caption = Valido_Importe(Chk0(rec!Saldo))
        End If
        rec.Close
        GrdCtaCte.SetFocus
    Else
        MsgBox "No se encontraron datos", vbExclamation, TIT_MSGBOX
        CmdNuevo_Click
        cboTipoProveedor.SetFocus
    End If
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
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
    If txtCodProveedor.Text <> "" Then
        Rep.SelectionFormula = "{CTA_CTE_PROVEEDORES.TPR_CODIGO}=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
        Rep.SelectionFormula = Rep.SelectionFormula & " AND {CTA_CTE_PROVEEDORES.PROV_CODIGO}=" & txtCodProveedor.Text
    End If
    If Not IsNull(FechaDesde.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {CTA_CTE_PROVEEDORES.CTA_CTE_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {CTA_CTE_PROVEEDORES.CTA_CTE_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        End If
    End If
    If Not IsNull(FechaHasta.Value) Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {CTA_CTE_PROVEEDORES.CTA_CTE_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {CTA_CTE_PROVEEDORES.CTA_CTE_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
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

    Rep.WindowTitle = "CTA-CTE de Proveedores..."
    Rep.ReportFileName = DRIVE & DirReport & "rptctacteproveedores.rpt"

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

Private Sub txtCodProveedor_Change()
    If txtCodProveedor.Text = "" Then
        txtProvRazSoc.Text = ""
        CmdBuscAprox.Enabled = False
        cmdListar.Enabled = False
    Else
        CmdBuscAprox.Enabled = True
        cmdListar.Enabled = True
    End If
End Sub

Private Sub txtCodProveedor_GotFocus()
    SelecTexto txtCodProveedor
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodProveedor_LostFocus()
    If txtCodProveedor.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        Rec1.Open BuscoProveedor(txtCodProveedor), DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtProvRazSoc.Text = Rec1!PROV_RAZSOC
            Call BuscaCodigoProxItemData(CInt(Rec1!TPR_CODIGO), cboTipoProveedor)
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtProvRazSoc_Change()
    If txtProvRazSoc.Text = "" Then
        txtCodProveedor.Text = ""
    End If
End Sub

Private Sub txtProvRazSoc_GotFocus()
    SelecTexto txtProvRazSoc
End Sub

Private Sub txtProvRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtProvRazSoc_LostFocus()
    If txtCodProveedor.Text = "" And txtProvRazSoc.Text <> "" Then
        rec.Open BuscoProveedor(txtProvRazSoc), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 5
                frmBuscar.TxtDescriB.Text = txtProvRazSoc.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 1
                    txtCodProveedor.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 2
                    txtProvRazSoc.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 3
                    Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboTipoProveedor)
                    txtCodProveedor_LostFocus
                Else
                    txtCodProveedor.SetFocus
                End If
            Else
                txtCodProveedor.Text = rec!PROV_CODIGO
                txtProvRazSoc.Text = rec!PROV_RAZSOC
                txtCodProveedor_LostFocus
            End If
        Else
            MsgBox "No se encontro el Proveedor", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        rec.Close
    ElseIf txtCodProveedor.Text = "" And txtProvRazSoc.Text = "" Then
        MsgBox "Debe elegir un Proveedor", vbExclamation, TIT_MSGBOX
        txtCodProveedor.SetFocus
    End If
End Sub

Private Sub CmdNuevo_Click()
    txtCodProveedor.Text = ""
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    lblProveedor.Caption = "Proveedor"
    lblFecha.Caption = "Fecha"
    lblSaldo.Caption = "0,00"
    lblTotalDebe.Caption = "0,00"
    lblTotalHaber.Caption = "0,00"
    lblSaldoActual1.Visible = False
    lblSaldoActual.Visible = False
    GrdCtaCte.Rows = 1
    GrdCtaCte.HighLight = flexHighlightNever
    cboTipoProveedor.ListIndex = 0
    cboTipoProveedor.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmCtaCteProveedores = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)
    'Cargo combo tipo Proveedor
    LlenarComboTipoProv
    CmdBuscAprox.Enabled = False
    cmdListar.Enabled = False
    GrdCtaCte.FormatString = "Tipo Comp.|^Fecha|^Número|>Debe|>Haber|>Saldo"
    GrdCtaCte.ColWidth(0) = 2000 'TIPO COMPROBANTE
    GrdCtaCte.ColWidth(1) = 1150 'FECHA
    GrdCtaCte.ColWidth(2) = 1250 'NUMERO
    GrdCtaCte.ColWidth(3) = 1200 'DEBE
    GrdCtaCte.ColWidth(4) = 1200 'HABER
    GrdCtaCte.ColWidth(5) = 1200 'SALDO
    GrdCtaCte.Rows = 2
    
    lblSaldoActual1.Visible = False
    lblSaldoActual.Visible = False
    lblSaldo.Caption = "0,00"
    lblTotalDebe.Caption = "0,00"
    lblTotalHaber.Caption = "0,00"
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""
End Sub

Private Sub LlenarComboTipoProv()
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboTipoProveedor.AddItem "TODOS"
        Do While rec.EOF = False
            cboTipoProveedor.AddItem rec!TPR_DESCRI
            cboTipoProveedor.ItemData(cboTipoProveedor.NewIndex) = rec!TPR_CODIGO
            rec.MoveNext
        Loop
        cboTipoProveedor.ListIndex = 0
    End If
    rec.Close
End Sub

Private Function BuscoProveedor(Pro As String) As String
    sql = "SELECT TPR_CODIGO,PROV_CODIGO, PROV_RAZSOC"
    sql = sql & " FROM PROVEEDOR"
    sql = sql & " WHERE"
    If txtCodProveedor.Text <> "" Then
        sql = sql & " PROV_CODIGO=" & XN(Pro)
    Else
        sql = sql & " PROV_RAZSOC LIKE '" & Pro & "%'"
    End If
    If cboTipoProveedor.List(cboTipoProveedor.ListIndex) <> "TODOS" Then
        'sql = sql & " AND TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    End If
    BuscoProveedor = sql
End Function

