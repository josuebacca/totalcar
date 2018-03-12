VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAnulaOrdenPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anular Orden de Pago...."
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameBuscar 
      Caption         =   "Buscar Orden de Pago por..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   90
      TabIndex        =   16
      Top             =   75
      Width           =   9435
      Begin VB.CommandButton CmdBuscAprox 
         Height          =   1215
         Left            =   8670
         MaskColor       =   &H000000FF&
         Picture         =   "frmAnulaOrdenPago.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Buscar  Orden de Pago"
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CheckBox chkFecha 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha"
         Height          =   195
         Left            =   390
         TabIndex        =   2
         Top             =   1080
         Width           =   810
      End
      Begin VB.ComboBox cboBuscaTipoProveedor 
         Height          =   315
         Left            =   2580
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   405
         Width           =   3900
      End
      Begin VB.CheckBox chkTipoProveedor 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Prov"
         Height          =   195
         Left            =   150
         TabIndex        =   0
         Top             =   510
         Width           =   1050
      End
      Begin VB.CheckBox chkProveedor 
         Alignment       =   1  'Right Justify
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   75
         TabIndex        =   1
         Top             =   795
         Width           =   1125
      End
      Begin VB.TextBox txtDesProv 
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
         Left            =   4020
         MaxLength       =   50
         TabIndex        =   18
         Tag             =   "Descripción"
         Top             =   780
         Width           =   4440
      End
      Begin VB.TextBox txtProveedor 
         Height          =   300
         Left            =   2580
         MaxLength       =   40
         TabIndex        =   4
         Top             =   780
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscarProveedor 
         Height          =   300
         Left            =   3585
         MaskColor       =   &H000000FF&
         Picture         =   "frmAnulaOrdenPago.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Buscar Proveedor"
         Top             =   780
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   2580
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   61407233
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   5280
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   61407233
         CurrentDate     =   41098
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   4260
         TabIndex        =   22
         Top             =   1185
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   1470
         TabIndex        =   21
         Top             =   1185
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Prov.:"
         Height          =   195
         Left            =   1695
         TabIndex        =   20
         Top             =   450
         Width           =   780
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1695
         TabIndex        =   19
         Top             =   825
         Width           =   780
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmAnulaOrdenPago.frx":2AAC
      Height          =   720
      Left            =   8625
      Picture         =   "frmAnulaOrdenPago.frx":2DB6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5295
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmAnulaOrdenPago.frx":30C0
      Height          =   720
      Left            =   6855
      Picture         =   "frmAnulaOrdenPago.frx":33CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5295
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmAnulaOrdenPago.frx":36D4
      Height          =   720
      Left            =   7740
      Picture         =   "frmAnulaOrdenPago.frx":39DE
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5295
      Width           =   870
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3480
      Left            =   60
      TabIndex        =   8
      Top             =   1740
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   6138
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorSel    =   8388736
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Left            =   4410
      TabIndex        =   15
      Top             =   5325
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Anulado"
      Height          =   195
      Left            =   5055
      TabIndex        =   14
      Top             =   5760
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Definitivo"
      Height          =   195
      Left            =   5055
      TabIndex        =   13
      Top             =   5550
      Width           =   660
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   4395
      Top             =   5790
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   150
      Left            =   4395
      Top             =   5595
      Width           =   540
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
      Left            =   210
      TabIndex        =   12
      Top             =   5535
      Width           =   750
   End
End
Attribute VB_Name = "frmAnulaOrdenPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim TipoCOMPROBANTE As Integer

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        FechaDesde.Enabled = True
        FechaHasta.Enabled = True
    Else
        FechaDesde.Enabled = False
        FechaHasta.Enabled = False
    End If
End Sub

Private Sub chkProveedor_Click()
    If chkProveedor.Value = Checked Then
        txtProveedor.Enabled = True
        cmdBuscarProveedor.Enabled = True
    Else
        txtProveedor.Text = ""
        txtProveedor.Enabled = False
        cmdBuscarProveedor.Enabled = False
    End If
End Sub

Private Sub chkTipoProveedor_Click()
    If chkTipoProveedor.Value = Checked Then
        cboBuscaTipoProveedor.Enabled = True
        cboBuscaTipoProveedor.ListIndex = 0
    Else
        cboBuscaTipoProveedor.Enabled = False
        cboBuscaTipoProveedor.ListIndex = -1
    End If
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT OP.OPG_NUMERO, OP.OPG_FECHA,TP.TPR_DESCRI, P.PROV_RAZSOC, E.EST_DESCRI,"
    sql = sql & " OP.EST_CODIGO, OP.TPR_CODIGO, OP.PROV_CODIGO, OP.TCO_CODIGO"
    sql = sql & " FROM ORDEN_PAGO OP, TIPO_PROVEEDOR TP, PROVEEDOR P, ESTADO_DOCUMENTO E"
    sql = sql & " WHERE"
    sql = sql & " OP.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND OP.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND P.TPR_CODIGO=TP.TPR_CODIGO"
    sql = sql & " AND OP.EST_CODIGO=E.EST_CODIGO"
    If chkTipoProveedor.Value = Checked Then sql = sql & " AND OP.TPR_CODIGO=" & cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.ListIndex)
    If txtProveedor.Text <> "" Then sql = sql & " AND OP.PROV_CODIGO=" & XN(txtProveedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND OP.OPG_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND OP.OPG_FECHA<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY OP.OPG_NUMERO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        TipoCOMPROBANTE = Rec1!TCO_CODIGO
        Do While Rec1.EOF = False
            GrdModulos.AddItem Format(Rec1!OPG_NUMERO, "00000000") & Chr(9) & Rec1!OPG_FECHA & Chr(9) & _
                               Rec1!TPR_DESCRI & " - " & Rec1!PROV_RAZSOC & Chr(9) & _
                               Rec1!EST_DESCRI & Chr(9) & Rec1!EST_CODIGO & Chr(9) & _
                               Rec1!EST_CODIGO & Chr(9) & Rec1!PROV_CODIGO & Chr(9) & _
                               Rec1!TPR_CODIGO
                               
            If Rec1!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            End If
            Rec1.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        chkTipoProveedor.SetFocus
    End If
    Rec1.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdBuscarProveedor_Click()
    frmBuscar.TipoBusqueda = 5
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 1
        txtProveedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 2
        txtDesProv.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 3
        Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboBuscaTipoProveedor)
        txtProveedor.SetFocus
    Else
        txtProveedor.SetFocus
    End If
End Sub

Private Sub cmdGrabar_Click()
    If MsgBox("¿Confirma Anular?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo SeClavo
    lblEstado.Caption = "Actualizando..."
    Screen.MousePointer = vbHourglass
    DBConn.BeginTrans
    
    ActualizoOrdenPago
    
    DBConn.CommitTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    CmdNuevo_Click
    Exit Sub

SeClavo:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub ActualizoOrdenPago()
    Dim SaldoFactura As String
    SaldoFactura = "0"
    
    For I = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(I, 4) <> GrdModulos.TextMatrix(I, 5) Then
            Set rec = New ADODB.Recordset
            sql = "UPDATE ORDEN_PAGO"
            sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(I, 5))
            sql = sql & " WHERE"
            sql = sql & " OPG_NUMERO=" & XN(GrdModulos.TextMatrix(I, 0))
            sql = sql & " AND OPG_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 1))
            DBConn.Execute sql

            'ACTUALIZO EL SALDO DE LOS FACTURAS
            sql = "SELECT FR.FPR_TCO_CODIGO,FR.FPR_NUMERO,FR.FPR_NROSUC,"
            sql = sql & "FR.OPG_IMPORTE,FC.FPR_SALDO"
            sql = sql & " FROM FACTURAS_ORDEN_PAGO FR, FACTURA_PROVEEDOR FC"
            sql = sql & " WHERE"
            sql = sql & " FR.OPG_NUMERO=" & XN(GrdModulos.TextMatrix(I, 0))
            sql = sql & " AND FR.OPG_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 1))
            sql = sql & " AND FR.FPR_TCO_CODIGO=FC.TCO_CODIGO"
            sql = sql & " AND FR.FPR_NUMERO=FC.FPR_NUMERO"
            sql = sql & " AND FR.FPR_NROSUC=FC.FPR_NROSUC"

            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                Do While rec.EOF = False
                    SaldoFactura = CDbl(rec!OPG_IMPORTE) + CDbl(rec!FPR_SALDO)
                    sql = "UPDATE FACTURA_PROVEEDOR"
                    sql = sql & " SET FPR_SALDO=" & XN(SaldoFactura)
                    sql = sql & " WHERE"
                    sql = sql & " TPR_CODIGO=" & XN(GrdModulos.TextMatrix(I, 7))
                    sql = sql & " AND TCO_CODIGO=" & XN(rec!FPR_TCO_CODIGO)
                    sql = sql & " AND FPR_NROSUC=" & XN(rec!FPR_NROSUC)
                    sql = sql & " AND FPR_NUMERO=" & XN(rec!FPR_NUMERO)
                    DBConn.Execute sql
                    SaldoFactura = "0"
                    rec.MoveNext
                Loop
            End If
            rec.Close
            
            'ACTUALIZO EL DINERO A CUENTA (RECIBO_CLIENTE_SALDO)
            sql = "DELETE FROM ORDEN_PAGO_SALDO"
            sql = sql & " WHERE"
            sql = sql & " OPG_NUMERO=" & XN(GrdModulos.TextMatrix(I, 0))
            sql = sql & " AND OPG_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 1))
            DBConn.Execute sql
            
            sql = "SELECT BAN_CODINT,CHE_NUMERO,CTA_NROCTA"
            sql = sql & " FROM DETALLE_ORDEN_PAGO"
            sql = sql & " WHERE"
            sql = sql & " OPG_NUMERO=" & XN(GrdModulos.TextMatrix(I, 0))
            sql = sql & " AND OPG_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 1))
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                Do While rec.EOF = False
                    If IsNull(rec!CTA_NROCTA) Then
                        'CHEQUES DE TERCEROS
                        'Cambio en Cheque_Estados 1 ES CHEQUES EN CARTERA
                        sql = "INSERT INTO CHEQUE_ESTADOS"
                        sql = sql & "(ECH_CODIGO,BAN_CODINT,CHE_NUMERO,CES_FECHA,CES_DESCRI) "
                        sql = sql & " VALUES ( 1,"
                        sql = sql & XN(rec!BAN_CODINT) & ","
                        sql = sql & XS(rec!CHE_NUMERO) & ","
                        sql = sql & XDQ(Date) & ","
                        sql = sql & "'CHEQUE EN CARTERA')"
                        DBConn.Execute sql
                    
                    Else
                        
                        'CHEQUES PROPIOS
                        'Insert en la Tabla de Estados de Cheques
                        sql = "INSERT INTO CHEQUE_PROPIO_ESTADO (CHEP_NUMERO,BAN_CODINT,ECH_CODIGO,CPES_FECHA,CPES_DESCRI)"
                        sql = sql & " VALUES ("
                        sql = sql & XS(rec!CHE_NUMERO) & ","
                        sql = sql & XN(rec!BAN_CODINT) & "," & XN(5) & ","
                        sql = sql & XDQ(Date) & ",'CHEQUE ANULADO')"
                        DBConn.Execute sql
                        
'                        'ACTUALIZO EL SALDO DE LA CTA-BANCARIA
'                        Set Rec1 = New ADODB.Recordset
'                        sql = "SELECT CHEP_IMPORT FROM CHEQUE_PROPIO"
'                        sql = sql & " WHERE"
'                        sql = sql & " BAN_CODINT=" & XN(rec!BAN_CODINT)
'                        sql = sql & " AND CHEP_NUMERO=" & XS(rec!CHE_NUMERO)
'                        sql = sql & " AND CTA_NROCTA=" & XS(rec!CTA_NROCTA)
'                        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'                        If Rec1.EOF = False Then
'                            sql = "UPDATE CTA_BANCARIA"
'                            sql = sql & " SET CTA_SALACT = CTA_SALACT + " & XN(Rec1!CHEP_IMPORT)
'                            sql = sql & " WHERE"
'                            sql = sql & " CTA_NROCTA=" & XS(rec!CTA_NROCTA)
'                            sql = sql & " AND BAN_CODINT=" & XN(rec!BAN_CODINT)
'                            DBConn.Execute sql
'                        End If
'                        Rec1.Close
                    End If
                    rec.MoveNext
                Loop
            End If
            If rec.State = 1 Then rec.Close
            'ACTUALIZO LA CTA-CTE
            DBConn.Execute QuitoCtaCteProveedores(GrdModulos.TextMatrix(I, 7), GrdModulos.TextMatrix(I, 6), _
                                               CStr(TipoCOMPROBANTE), Sucursal, GrdModulos.TextMatrix(I, 0))
        End If
    Next
End Sub

Private Sub CmdSalir_Click()
    Set frmAnulaOrdenPago = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
     Set rec = New ADODB.Recordset
     Set Rec2 = New ADODB.Recordset
     
    Call Centrar_pantalla(Me)
    ConfiguroGrilla
    LlenarComboTipoProv
    cboBuscaTipoProveedor.Enabled = False
    txtProveedor.Enabled = False
    cmdBuscarProveedor.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    lblEstado.Caption = ""
End Sub

Private Sub LlenarComboTipoProv()
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboBuscaTipoProveedor.AddItem rec!TPR_DESCRI
            cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.NewIndex) = rec!TPR_CODIGO
            rec.MoveNext
        Loop
        cboBuscaTipoProveedor.ListIndex = -1
    End If
    rec.Close
End Sub

Private Sub ConfiguroGrilla()
    GrdModulos.FormatString = "^Número|^Fecha|Proveedor|^Estado|codigo estado|" _
                            & "codigo estado que cambio|COD proveedor|cod tipo Proveedor"
    GrdModulos.ColWidth(0) = 1100 'NUMERO
    GrdModulos.ColWidth(1) = 1200 'FECHA_ORD PAG
    GrdModulos.ColWidth(2) = 5000 'PROVEEDOR
    GrdModulos.ColWidth(3) = 2000 'ESTADO
    GrdModulos.ColWidth(4) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(5) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(6) = 0    'CODIGO PROVEEDOR
    GrdModulos.ColWidth(7) = 0    'CODIGO TIPOMPROVEEDOR
    GrdModulos.Cols = 8
    GrdModulos.Rows = 2
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = 2 Then
            MsgBox "No se puede cambiar el estado a la Orden de Pago" & Chr(13) & _
                   " la misma ya fue Anulada", vbExclamation, TIT_MSGBOX
            
            Exit Sub
        End If
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 3 Then
            GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2
            GrdModulos.TextMatrix(GrdModulos.RowSel, 3) = "ANULADO"
            Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
             
        ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
            GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 3
            GrdModulos.TextMatrix(GrdModulos.RowSel, 3) = "DEFINITIVO"
            Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
        End If
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then GrdModulos_DblClick
End Sub

Private Sub CmdNuevo_Click()
    TipoCOMPROBANTE = 0
    txtProveedor.Text = ""
    txtDesProv.Text = ""
    cboBuscaTipoProveedor.ListIndex = -1
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    GrdModulos.Rows = 1
    GrdModulos.Rows = 2
    txtProveedor.Enabled = False
    cboBuscaTipoProveedor.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    chkProveedor.Value = Unchecked
    chkTipoProveedor.Value = Unchecked
    chkFecha.Value = Unchecked
    chkTipoProveedor.SetFocus
End Sub

Private Sub txtProveedor_Change()
    If txtProveedor.Text = "" Then
        txtDesProv.Text = ""
    End If
End Sub

Private Sub txtProveedor_GotFocus()
    SelecTexto txtProveedor
End Sub

Private Sub txtProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtProveedor_LostFocus()
    If txtProveedor.Text <> "" Then
        sql = "SELECT TPR_CODIGO,PROV_CODIGO,PROV_RAZSOC,"
        sql = sql & " FROM PROVEEDOR"
        sql = sql & " WHERE"
        sql = sql & " PROV_CODIGO=" & XN(txtProveedor)
        If chkTipoProveedor.Value = Checked Then
            sql = sql & " and tpr_codigo=" & XN(cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.ListIndex))
        End If
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtDesProv.Text = Rec1!PROV_RAZSOC
            Call BuscaCodigoProxItemData(CInt(Rec1!TPR_CODIGO), cboBuscaTipoProveedor)
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtProveedor.SetFocus
        End If
        Rec1.Close
    End If
End Sub

