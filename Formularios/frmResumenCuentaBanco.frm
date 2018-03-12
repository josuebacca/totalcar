VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmResumenCuentaBanco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de Cuenta - Banco"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdResumenCuenta 
      Height          =   2985
      Left            =   75
      TabIndex        =   4
      Top             =   1260
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   5265
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   8388736
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   75
      TabIndex        =   13
      Top             =   15
      Width           =   8760
      Begin VB.CommandButton cmdVerResumen 
         Caption         =   "&Ver"
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
         Left            =   6345
         TabIndex        =   3
         Top             =   675
         Width           =   1485
      End
      Begin VB.ComboBox CboCuentas 
         Height          =   315
         Left            =   6345
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1485
      End
      Begin VB.ComboBox CboBancoBoleta 
         Height          =   315
         ItemData        =   "frmResumenCuentaBanco.frx":0000
         Left            =   720
         List            =   "frmResumenCuentaBanco.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   3255
      End
      Begin VB.TextBox TxtBanCodInt 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4860
         TabIndex        =   14
         Top             =   270
         Width           =   465
      End
      Begin MSComCtl2.DTPicker Fecha 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   56688641
         CurrentDate     =   41098
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   690
         Width           =   585
      End
      Begin VB.Label lblPeriodo1 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2250
         TabIndex        =   19
         Top             =   615
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   6
         Left            =   5670
         TabIndex        =   17
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   16
         Top             =   330
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   7
         Left            =   4185
         TabIndex        =   15
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   720
      Left            =   7095
      Picture         =   "frmResumenCuentaBanco.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4980
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   720
      Left            =   7980
      Picture         =   "frmResumenCuentaBanco.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4980
      Width           =   870
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   720
      Left            =   6225
      Picture         =   "frmResumenCuentaBanco.frx":0618
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4980
      Width           =   855
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
      Left            =   75
      TabIndex        =   10
      Top             =   4665
      Width           =   3270
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2085
         TabIndex        =   6
         Top             =   315
         Width           =   1005
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   1035
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   345
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   300
         Width           =   585
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   5550
      Top             =   5250
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   5055
      Top             =   5220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSaldoActual 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   7890
      TabIndex        =   21
      Top             =   4260
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
      Left            =   3420
      TabIndex        =   18
      Top             =   5130
      Width           =   750
   End
End
Attribute VB_Name = "frmResumenCuentaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AplicoImpuesto As Boolean
Dim ValorImpuesto As Double

Private Sub CboBancoBoleta_LostFocus()
     Me.TxtBanCodInt.Text = CStr(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
End Sub

Private Sub CboCuentas_GotFocus()
    If Trim(CboBancoBoleta.Text) <> "" Then
        CboCuentas.Clear
        Call CargoCtaBancaria(CStr(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)))
        CboCuentas.ListIndex = 0
    End If
End Sub

Private Sub CargoCtaBancaria(Banco As String)
    Set rec = New ADODB.Recordset
    sql = "SELECT CTA_NROCTA FROM CTA_BANCARIA"
    sql = sql & " WHERE BAN_CODINT=" & XN(Banco)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
     Do While rec.EOF = False
         CboCuentas.AddItem Trim(rec!CTA_NROCTA)
         rec.MoveNext
     Loop
    End If
    rec.Close
End Sub

Private Sub cmdListar_Click()
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.Formulas(0) = ""
    
    lblEstado.Caption = "Buscando Listado..."
    
    Rep.Formulas(0) = "SALDO='" & lblSaldoActual.Caption & "'"
        
    Rep.WindowTitle = "Resumen Cuenta - Banco..."
    Rep.ReportFileName = DRIVE & DirReport & "ResumenCuentaBanco.rpt"
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     Rep.Formulas(0) = ""
     lblEstado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
    CboBancoBoleta.ListIndex = 0
    CboCuentas.Clear
    TxtBanCodInt.Text = ""
    grdResumenCuenta.Rows = 1
    grdResumenCuenta.Rows = 2
    lblSaldoActual.Caption = "Saldo"
    optPantalla.Value = True
    CboBancoBoleta.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmResumenCuentaBanco = Nothing
    Unload Me
End Sub

Private Sub cmdVerResumen_Click()
    Dim FechaUltimoSaldo As String
    Dim Saldo As Double
    Dim I As Integer
    I = 0
    FechaUltimoSaldo = ""
    Saldo = 0
    grdResumenCuenta.Rows = 1
    lblEstado.Caption = "Buscando Movimientos..."
    'BORRO LA TEMPORAL
    sql = "DELETE FROM TMP_RESUMEN_CUENTA_BANCO"
    DBConn.Execute sql
    
    'BUSCO EL ULTIMO SALDO
    sql = "SELECT MAX(RCB_FECHA) AS FECHA, RCB_SALDO"
    sql = sql & " FROM RESUMEN_CUENTA_BANCO"
    sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    sql = sql & " GROUP BY RCB_SALDO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        FechaUltimoSaldo = rec!Fecha
        Saldo = CDbl(rec!RCB_SALDO)
        I = I + 1
        'INSERTO EN LA TEMPORAL
        sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
        sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
        sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
        sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
        sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
        sql = sql & XS(lblPeriodo1.Caption) & ","
        sql = sql & XDQ("") & ","
        sql = sql & XS("Saldo Anterior - " & Format(rec!Fecha, "mmmm/yyyy")) & ","
        sql = sql & XS("") & ","
        sql = sql & XN("0") & ","
        sql = sql & XN("0") & ","
        sql = sql & XN(CStr(Saldo)) & ","
        sql = sql & XN(CStr(I)) & ")"
        DBConn.Execute sql
'        grdResumenCuenta.AddItem "" & Chr(9) & "Saldo Anterior - " & Format(rec!Fecha, "mmmm/yyyy") & Chr(9) & _
'                                 "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Valido_Importe(rec!RCB_SALDO)
    End If
    rec.Close
    
'---------------BUSCO LAS BOLETAS DE DEPOSITO--------------------------------
    sql = "SELECT BOL_FECHA,BOL_NUMERO,BOL_TOTAL"
    sql = sql & " FROM BOL_DEPOSITO"
    sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    sql = sql & " AND EBO_CODIGO<> 2" 'BOLETAS NO ANULADAS
    sql = sql & " AND MONTH(BOL_FECHA)=" & XN(Month(Fecha.Value))
    sql = sql & " AND YEAR(BOL_FECHA)=" & XN(Year(Fecha.Value))
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            'INSERTO DEPOSITO
            Saldo = Saldo + CDbl(rec!BOL_TOTAL)
            I = I + 1
            sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
            sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
            sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
            sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
            sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
            sql = sql & XS(lblPeriodo1.Caption) & ","
            sql = sql & XDQ(rec!BOL_FECHA) & ","
            sql = sql & XS("DEPOSITO") & ","
            sql = sql & XS(rec!BOL_NUMERO) & ","
            sql = sql & XN("0") & ","
            sql = sql & XN(rec!BOL_TOTAL) & ","
            sql = sql & XN(CStr(Saldo)) & ","
            sql = sql & XN(CStr(I)) & ")"
            DBConn.Execute sql
            'VERIFICO SI LE APLICO EL IMPUESTO AL CHEQUE
            If AplicoImpuesto = True Then
                Saldo = Saldo + (CDbl(rec!BOL_TOTAL) * (ValorImpuesto))
                I = I + 1
                sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
                sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
                sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
                sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
                sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
                sql = sql & XS(lblPeriodo1.Caption) & ","
                sql = sql & XDQ(rec!BOL_FECHA) & ","
                sql = sql & XS("GRAVAMEN LEY 25413 S/CRE") & ","
                sql = sql & XS(rec!BOL_NUMERO) & ","
                sql = sql & XN(CStr(CDbl(rec!BOL_TOTAL) * (ValorImpuesto))) & ","
                sql = sql & XN("0") & ","
                sql = sql & XN(CStr(Saldo)) & ","
                sql = sql & XN(CStr(I)) & ")"
                DBConn.Execute sql
            End If
            rec.MoveNext
        Loop
    End If
    rec.Close
    
'----------BUSCO LOS GASTOS BANCARIOS-------------------------------------
    sql = "SELECT GB.GBA_NUMERO,GB.GBA_FECHA,GB.GBA_IMPORTE,TG.TGB_DESCRI,GB.GBA_IMPUESTO"
    sql = sql & " FROM GASTOS_BANCARIOS GB,TIPO_GASTO_BANCARIO TG"
    sql = sql & " WHERE GB.BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND GB.CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    sql = sql & " AND GB.TGB_CODIGO=TG.TGB_CODIGO"
    sql = sql & " AND MONTH(GB.GBA_FECHA)=" & XN(Month(Fecha.Value))
    sql = sql & " AND YEAR(GB.GBA_FECHA)=" & XN(Year(Fecha.Value))
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            Saldo = Saldo - CDbl(rec!GBA_IMPORTE)
            I = I + 1
            'INSERTO GASTOS BANCARIOS
            sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
            sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
            sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
            sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
            sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
            sql = sql & XS(lblPeriodo1.Caption) & ","
            sql = sql & XDQ(rec!GBA_FECHA) & ","
            sql = sql & XS(rec!TGB_DESCRI) & ","
            sql = sql & XS(rec!GBA_NUMERO) & ","
            sql = sql & XN(rec!GBA_IMPORTE) & ","
            sql = sql & XN("0") & ","
            sql = sql & XN(CStr(Saldo)) & ","
            sql = sql & XN(CStr(I)) & ")"
            DBConn.Execute sql
            'VERIFICO SI LE APLICO EL IMPUESTO AL CHEQUE
            If AplicoImpuesto = True And rec!GBA_IMPUESTO = "S" Then
                Saldo = Saldo - (CDbl(rec!GBA_IMPORTE) * (ValorImpuesto))
                I = I + 1
                sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
                sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
                sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
                sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
                sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
                sql = sql & XS(lblPeriodo1.Caption) & ","
                sql = sql & XDQ(rec!GBA_FECHA) & ","
                sql = sql & XS("GRAVAMEN LEY 25413 S/DEB") & ","
                sql = sql & XS(rec!GBA_NUMERO) & ","
                sql = sql & XN(CStr(CDbl(rec!GBA_IMPORTE) * (ValorImpuesto))) & ","
                sql = sql & XN("0") & ","
                sql = sql & XN(CStr(Saldo)) & ","
                sql = sql & XN(CStr(I)) & ")"
                DBConn.Execute sql
            End If
            rec.MoveNext
        Loop
    End If
    rec.Close
    
'--------------BUSCO LOS CHEQUES LIBRADOS--------------------------------
    sql = "SELECT CHEP_FECVTO,CHEP_NUMERO,CHEP_IMPORT"
    sql = sql & " FROM ChequePropioEstadoVigente"
    sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    sql = sql & " AND ECH_CODIGO IN (7,8)" 'CHEQUES LIBRADOS O RESTITUIDOS
    sql = sql & " AND MONTH(CHEP_FECVTO)=" & XN(Month(Fecha.Value))
    sql = sql & " AND YEAR(CHEP_FECVTO)=" & XN(Year(Fecha.Value))
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            Saldo = Saldo - CDbl(rec!CHEP_IMPORT)
            I = I + 1
            'INSERTO CHEQUES LIBRADOS
            sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
            sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
            sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
            sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
            sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
            sql = sql & XS(lblPeriodo1.Caption) & ","
            sql = sql & XDQ(rec!CHEP_FECVTO) & "," 'FECHA DE PAGO
            sql = sql & XS("CHEQUES LIBRADOS") & ","
            sql = sql & XS(rec!CHEP_NUMERO) & ","
            sql = sql & XN(rec!CHEP_IMPORT) & ","
            sql = sql & XN("0") & ","
            sql = sql & XN(CStr(Saldo)) & ","
            sql = sql & XN(CStr(I)) & ")"
            DBConn.Execute sql
            'VERIFICO SI LE APLICO EL IMPUESTO AL CHEQUE
            If AplicoImpuesto = True Then
                Saldo = Saldo - (CDbl(rec!CHEP_IMPORT) * (ValorImpuesto))
                I = I + 1
                sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
                sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
                sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
                sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
                sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
                sql = sql & XS(lblPeriodo1.Caption) & ","
                sql = sql & XDQ(rec!CHEP_FECVTO) & ","
                sql = sql & XS("GRAVAMEN LEY 25413 S/DEB") & ","
                sql = sql & XS(rec!CHEP_NUMERO) & ","
                sql = sql & XN(CStr(CDbl(rec!CHEP_IMPORT) * (ValorImpuesto))) & ","
                sql = sql & XN("0") & ","
                sql = sql & XN(CStr(Saldo)) & ","
                sql = sql & XN(CStr(I)) & ")"
                DBConn.Execute sql
            End If
            rec.MoveNext
        Loop
    End If
    rec.Close
'------------CARGO GRILLA-------------------------------
    sql = "SELECT * FROM TMP_RESUMEN_CUENTA_BANCO"
    sql = sql & " ORDER BY FECHA,ORDEN"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            grdResumenCuenta.AddItem rec!Fecha & Chr(9) & rec!Descripcion & Chr(9) & _
                                    rec!COMPROBANTE & Chr(9) & Valido_Importe(rec!DEBITO) & Chr(9) & _
                                    Valido_Importe(rec!CREDITO) & Chr(9) & Valido_Importe(CStr(rec!Saldo))
            rec.MoveNext
        Loop
    End If
    rec.Close
    lblSaldoActual.Caption = "Saldo Actual: " & Valido_Importe(CStr(Saldo))
    lblEstado.Caption = ""
End Sub

Private Sub Fecha_Change()
    If Trim(Fecha.Value) <> "" Then
        lblPeriodo1.Caption = UCase(Format(Fecha.Value, "mmmm/yyyy"))
    Else
        lblPeriodo1.Caption = ""
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    'CARGO COMBO BANCO
    CargoBanco
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""
    'CONFIGURO GRILLA
    grdResumenCuenta.FormatString = "^Fecha|Descripción|^Comprob|>Débito|>Crédito|>Saldo"
    grdResumenCuenta.ColWidth(0) = 1100 'FECHA
    grdResumenCuenta.ColWidth(1) = 2800 'DESCRIPCION
    grdResumenCuenta.ColWidth(2) = 1200 'COMPROBANTE
    grdResumenCuenta.ColWidth(3) = 1100 'DEBITO
    grdResumenCuenta.ColWidth(4) = 1100 'CREDITO
    grdResumenCuenta.ColWidth(5) = 1100 'SALDO
    grdResumenCuenta.Cols = 6
    grdResumenCuenta.Rows = 2
    Fecha.Value = Date
    
    sql = "SELECT APLICA_IMPUESTO,VALOR_IMPUESTO FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        AplicoImpuesto = True 'APLICO IMPUESTO
        ValorImpuesto = CDbl(rec!VALOR_IMPUESTO)
    Else
        AplicoImpuesto = False 'NO APLICO IMPUESTO
        ValorImpuesto = 0
    End If
    rec.Close
End Sub

Private Sub CargoBanco()
    sql = "SELECT B.BAN_DESCRI, B.BAN_CODINT"
    sql = sql & " FROM BANCO B, CTA_BANCARIA CB"
    sql = sql & " WHERE B.BAN_CODINT=CB.BAN_CODINT"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            CboBancoBoleta.AddItem Trim(rec!BAN_DESCRI)
            CboBancoBoleta.ItemData(CboBancoBoleta.NewIndex) = Trim(rec!BAN_CODINT)
            rec.MoveNext
        Loop
        CboBancoBoleta.ListIndex = 0
    End If
    rec.Close
End Sub

