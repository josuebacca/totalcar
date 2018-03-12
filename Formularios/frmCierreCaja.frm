VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCierreCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CIERRE DE CAJA "
   ClientHeight    =   8010
   ClientLeft      =   75
   ClientTop       =   1395
   ClientWidth     =   10740
   Icon            =   "frmCierreCaja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10740
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      DisabledPicture =   "frmCierreCaja.frx":08CA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   1065
      Picture         =   "frmCierreCaja.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1935
      Width           =   990
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   90
      Picture         =   "frmCierreCaja.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1935
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "frmCierreCaja.frx":17A8
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   2055
      Picture         =   "frmCierreCaja.frx":1AB2
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1935
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Height          =   1890
      Left            =   90
      TabIndex        =   3
      Top             =   -45
      Width           =   2985
      Begin MSComCtl2.DTPicker Fecha1 
         Height          =   345
         Left            =   930
         TabIndex        =   10
         Top             =   570
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         _Version        =   393216
         Format          =   16842753
         CurrentDate     =   37629
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar Planilla"
         Height          =   330
         Left            =   945
         TabIndex        =   0
         Top             =   1170
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   615
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   885
      Picture         =   "frmCierreCaja.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7035
      Width           =   1395
   End
   Begin MSFlexGridLib.MSFlexGrid GrdDep 
      Height          =   7935
      Left            =   3195
      TabIndex        =   1
      Top             =   45
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   13996
      _Version        =   393216
      Rows            =   25
      Cols            =   4
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   260
      BackColor       =   16777215
      BackColorFixed  =   14737632
      GridColorFixed  =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   0
      ScrollBars      =   2
      BorderStyle     =   0
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   1710
      Top             =   4590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
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
      Left            =   180
      TabIndex        =   11
      Top             =   5535
      Width           =   750
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "ANULADOS: "
      Height          =   255
      Index           =   3
      Left            =   -2340
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "DESDE: "
      Height          =   300
      Index           =   1
      Left            =   -2340
      TabIndex        =   6
      Top             =   2610
      Width           =   1095
   End
End
Attribute VB_Name = "frmCierreCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Dim I As Integer
Dim Representada As Integer

Private Function Valido_Caja() As Boolean
    sql = "SELECT MAX(CAJA_FECHA) AS FECHA FROM CAJA"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        If Fecha1.Value < rec!Fecha Then
            MsgBox "No puede cerrar la caja, ya que la fecha seleccionada es menor al de la última caja cerada", vbCritical, TIT_MSGBOX
            Valido_Caja = False
        Else
            Valido_Caja = True
        End If
    End If
    rec.Close
End Function

Private Sub CmdCargar_Click()
    Dim FechaBusqueda As String
   
    FechaBusqueda = DateAdd("d", -1, Date)
    
    Screen.MousePointer = vbHourglass
    CmdNuevo_Click
    lblEstado.Caption = "Cargando datos de caja..."
    
    Dim TOTAL As Double

    If GrdDep.Rows > 1 Then
        
        Set rec = New ADODB.Recordset
        sql = "SELECT * FROM CAJA"
        sql = sql & " WHERE CAJA_SALDOIF='I'"
        sql = sql & " AND CAJA_FECHA = " & XDQ(Fecha1.Value)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            cmdGrabar.Enabled = False
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            MsgBox "No puede cerrar la caja, ya que la fecha seleccionada es igual al de la última caja cerada", vbCritical, TIT_MSGBOX
            Exit Sub
        Else
            rec.Close
            If Valido_Caja = False Then
                lblEstado.Caption = ""
                Screen.MousePointer = vbNormal
                Exit Sub
            End If
            cmdGrabar.Enabled = True
            sql = "SELECT CAJA_FECHA, CAJA_SALDO_CHEQUES, CAJA_SALDO_PESOS,"
            sql = sql & " CAJA_SALDO_LN, CAJA_SALDO_LC, CAJA_SALDO_OTROS"
            sql = sql & " FROM CAJA"
            sql = sql & " WHERE CAJA_SALDOIF='F'"
            sql = sql & " AND CAJA_FECHA= (SELECT MAX(CAJA_FECHA) AS FECHA FROM CAJA)"
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        End If
        'SALDO INICIAL
        If rec.EOF = False Then
          FechaBusqueda = rec!CAJA_FECHA
          GrdDep.TextMatrix(3, 1) = Valido_Importe(Chk0(rec!CAJA_SALDO_CHEQUES))
          GrdDep.TextMatrix(5, 1) = Valido_Importe(Chk0(rec!CAJA_SALDO_PESOS))
          GrdDep.TextMatrix(6, 1) = Valido_Importe(Chk0(rec!CAJA_SALDO_LN))
          GrdDep.TextMatrix(7, 1) = Valido_Importe(Chk0(rec!CAJA_SALDO_LC))
          GrdDep.TextMatrix(8, 1) = Valido_Importe(Chk0(rec!CAJA_SALDO_OTROS))
          GrdDep.TextMatrix(9, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(5, 1)) + CDbl(GrdDep.TextMatrix(6, 1)) + CDbl(GrdDep.TextMatrix(7, 1)) + CDbl(GrdDep.TextMatrix(8, 1))))
          GrdDep.TextMatrix(10, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(9, 1)) + CDbl(GrdDep.TextMatrix(3, 1))))
        End If
        rec.Close
        
        Call BUSCARINGRESO_MONEDA(FechaBusqueda)
        Call BUSCARINGRESO_CHEQUES(FechaBusqueda)
        SUMO_INGRESOS
        Call BUSCAREGRESOS_MONEDA(FechaBusqueda)
        Call BUSCAREGRESO_MONEDA_LIQUIDACION(FechaBusqueda)
        Call BUSCAREGRESOS_CHEQUES(FechaBusqueda)
        Call BUSCAREGRESO_CHEQUES_LIQUIDACION(FechaBusqueda)
        SUMO_EGRESOS
        SUMO_SALDOFINAL
    End If
    
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""

End Sub

Private Sub BUSCARINGRESO_MONEDA(FechaCaja As String)
    'BUSCO EN LOS RECIBOS
    sql = "SELECT DR.MON_CODIGO, SUM(DR.DRE_MONIMP) AS IMPORTE"
    sql = sql & " FROM RECIBO_CLIENTE R, DETALLE_RECIBO_CLIENTE DR"
    sql = sql & " WHERE R.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    'sql = sql & " AND R.REP_CODIGO=DR.REP_CODIGO"
    sql = sql & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    'sql = sql & " AND R.REP_CODIGO=" & XN(CStr(Representada))
    sql = sql & " AND R.EST_CODIGO=3"
    sql = sql & " AND R.REC_FECHA >" & XDQ(FechaCaja)
    sql = sql & " AND R.REC_FECHA <=" & XDQ(Fecha1.Value)
    sql = sql & " GROUP BY DR.MON_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            Select Case rec!MON_CODIGO
                Case 1 'PESOS
                    GrdDep.TextMatrix(16, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(16, 1)) + CDbl(rec!Importe)))
                Case 3 'LECOP CORDOBA
                    GrdDep.TextMatrix(18, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(18, 1)) + CDbl(rec!Importe)))
                Case 4 'LECOP NACION
                    GrdDep.TextMatrix(17, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(17, 1)) + CDbl(rec!Importe)))
                Case Else 'OTRAS MONEDAS
                    GrdDep.TextMatrix(19, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(19, 1)) + CDbl(rec!Importe)))
            End Select
            rec.MoveNext
        Loop
    End If
    rec.Close
    
    'BUSCO EN CAJA_INGRESOS
    sql = "SELECT MON_CODIGO, SUM(CIGR_IMPORTE) AS IMPORTE"
    sql = sql & " FROM CAJA_INGRESO"
    sql = sql & " WHERE"
    sql = sql & " CIGR_FECHA >" & XDQ(FechaCaja)
    sql = sql & " AND CIGR_FECHA <=" & XDQ(Fecha1.Value)
    sql = sql & " GROUP BY MON_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            Select Case rec!MON_CODIGO
                Case 1 'PESOS
                    GrdDep.TextMatrix(16, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(16, 1)) + CDbl(rec!Importe)))
                Case 3 'LECOP CORDOBA
                    GrdDep.TextMatrix(18, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(18, 1)) + CDbl(rec!Importe)))
                Case 4 'LECOP NACION
                    GrdDep.TextMatrix(17, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(17, 1)) + CDbl(rec!Importe)))
                Case Else 'OTRAS MONEDAS
                    GrdDep.TextMatrix(19, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(19, 1)) + CDbl(rec!Importe)))
            End Select
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub BUSCARINGRESO_CHEQUES(FechaCaja As String)
    'BUSCO LOS CHEQUES QUE INGRESARON POR LOS RECIBOS
    sql = "SELECT DR.BAN_CODINT, DR.CHE_NUMERO, CH.CHE_IMPORT"
    sql = sql & " FROM RECIBO_CLIENTE R, DETALLE_RECIBO_CLIENTE DR, CHEQUE CH"
    sql = sql & " WHERE R.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    'sql = sql & " AND R.REP_CODIGO=DR.REP_CODIGO"
    sql = sql & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    sql = sql & " AND DR.BAN_CODINT=CH.BAN_CODINT"
    sql = sql & " AND DR.CHE_NUMERO=CH.CHE_NUMERO"
    'sql = sql & " AND R.REP_CODIGO=" & XN(CStr(Representada))
    sql = sql & " AND R.EST_CODIGO=3"
    sql = sql & " AND R.REC_FECHA >" & XDQ(FechaCaja)
    sql = sql & " AND R.REC_FECHA <=" & XDQ(Fecha1.Value)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdDep.TextMatrix(14, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(14, 1)) + CDbl(rec!che_import)))
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub SUMO_INGRESOS()
    'TOTAL MONEDAS
    GrdDep.TextMatrix(20, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(16, 1)) + CDbl(GrdDep.TextMatrix(17, 1)) + CDbl(GrdDep.TextMatrix(18, 1)) + CDbl(GrdDep.TextMatrix(19, 1))))
    'TOTAL INGRESOS
    GrdDep.TextMatrix(20, 4) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(20, 1)) + CDbl(GrdDep.TextMatrix(14, 1))))
End Sub

Private Sub BUSCAREGRESOS_MONEDA(FechaCaja As String)
    'BUSCO EN LAS ORDNES DE PAGO
    sql = "SELECT DO.MON_CODIGO, SUM(DO.DOP_MONIMP) AS IMPORTE"
    sql = sql & " FROM ORDEN_PAGO O, DETALLE_ORDEN_PAGO DO"
    sql = sql & " WHERE O.OPG_NUMERO=DO.OPG_NUMERO"
    sql = sql & " AND O.OPG_FECHA=DO.OPG_FECHA"
    sql = sql & " AND O.TCO_CODIGO=DO.TCO_CODIGO"
    sql = sql & " AND O.EST_CODIGO=3"
    sql = sql & " AND O.OPG_FECHA >" & XDQ(FechaCaja)
    sql = sql & " AND O.OPG_FECHA <=" & XDQ(Fecha1.Value)
    sql = sql & " GROUP BY DO.MON_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            Select Case rec!MON_CODIGO
                Case 1 'PESOS
                    GrdDep.TextMatrix(25, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(25, 1)) + CDbl(rec!Importe)))
                Case 3 'LECOP CORDOBA
                    GrdDep.TextMatrix(27, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(27, 1)) + CDbl(rec!Importe)))
                Case 4 'LECOP NACION
                    GrdDep.TextMatrix(26, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(26, 1)) + CDbl(rec!Importe)))
                Case Else 'OTRAS MONEDAS
                    GrdDep.TextMatrix(28, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(28, 1)) + CDbl(rec!Importe)))
            End Select
            rec.MoveNext
        Loop
    End If
    rec.Close
    
    'BUSCO EN CAJA_EGRESOS
    sql = "SELECT MON_CODIGO, SUM(CEGR_IMPORTE) AS IMPORTE"
    sql = sql & " FROM CAJA_EGRESO"
    sql = sql & " WHERE"
    sql = sql & " CEGR_FECHA >" & XDQ(FechaCaja)
    sql = sql & " AND CEGR_FECHA <=" & XDQ(Fecha1.Value)
    sql = sql & " GROUP BY MON_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            Select Case rec!MON_CODIGO
                Case 1 'PESOS
                    GrdDep.TextMatrix(25, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(25, 1)) + CDbl(rec!Importe)))
                Case 3 'LECOP CORDOBA
                    GrdDep.TextMatrix(27, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(27, 1)) + CDbl(rec!Importe)))
                Case 4 'LECOP NACION
                    GrdDep.TextMatrix(26, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(26, 1)) + CDbl(rec!Importe)))
                Case Else 'OTRAS MONEDAS
                    GrdDep.TextMatrix(28, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(28, 1)) + CDbl(rec!Importe)))
            End Select
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub BUSCAREGRESO_MONEDA_LIQUIDACION(FechaCaja As String)
    'BUSCO EN LOS RECIBOS LA PLATA QUE SALIO POR LA COBRANZA
    'AQUI ME FIJO EN LA FECHA EN LA CUAL REALIZO
    'LA LIQUIDACION A LAS REPRESENTADAS PARA SACAR LOS FONDOS DE LA MISMA
    sql = "SELECT DR.MON_CODIGO, SUM(DR.DRE_MONIMP) AS IMPORTE"
    sql = sql & " FROM RECIBO_CLIENTE R, DETALLE_RECIBO_CLIENTE DR"
    sql = sql & " WHERE R.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    'sql = sql & " AND R.REP_CODIGO=DR.REP_CODIGO"
    sql = sql & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    'sql = sql & " AND R.REP_CODIGO=" & XN(CStr(Representada))
    sql = sql & " AND R.EST_CODIGO=3"
    sql = sql & " AND R.REC_LISTADO IS NOT NULL"
    'sql = sql & " AND R.REC_FECLIQUI IS NOT NULL"
    sql = sql & " AND R.REC_FECLIQUI >" & XDQ(FechaCaja)
    sql = sql & " AND R.REC_FECLIQUI <=" & XDQ(Fecha1.Value)
    sql = sql & " GROUP BY DR.MON_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            If Not IsNull(rec!MON_CODIGO) Then
            Select Case rec!MON_CODIGO
                Case 1 'PESOS
                    GrdDep.TextMatrix(25, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(25, 1)) + CDbl(rec!Importe)))
                Case 3 'LECOP CORDOBA
                    GrdDep.TextMatrix(27, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(27, 1)) + CDbl(rec!Importe)))
                Case 4 'LECOP NACION
                    GrdDep.TextMatrix(26, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(26, 1)) + CDbl(rec!Importe)))
                Case Else 'OTRAS MONEDAS
                    GrdDep.TextMatrix(28, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(28, 1)) + CDbl(rec!Importe)))
            End Select
            End If
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub BUSCAREGRESOS_CHEQUES(FechaCaja As String)
    'BUSCO LOS CHEQUES QUE SALIERON POR LA ORDEN DE PAGO
    sql = "SELECT DO.BAN_CODINT, DO.CHE_NUMERO, CH.CHE_IMPORT"
    sql = sql & " FROM ORDEN_PAGO O, DETALLE_ORDEN_PAGO DO, CHEQUE CH"
    sql = sql & " WHERE O.OPG_NUMERO=DO.OPG_NUMERO"
    sql = sql & " AND O.OPG_FECHA=DO.OPG_FECHA"
    sql = sql & " AND O.TCO_CODIGO=DO.TCO_CODIGO"
    sql = sql & " AND DO.BAN_CODINT=CH.BAN_CODINT"
    sql = sql & " AND DO.CHE_NUMERO=CH.CHE_NUMERO"
    sql = sql & " AND DO.CTA_NROCTA IS NULL"
    sql = sql & " AND O.EST_CODIGO=3"
    sql = sql & " AND O.OPG_FECHA >" & XDQ(FechaCaja)
    sql = sql & " AND O.OPG_FECHA <=" & XDQ(Fecha1.Value)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdDep.TextMatrix(23, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(23, 1)) + CDbl(rec!che_import)))
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub BUSCAREGRESO_CHEQUES_LIQUIDACION(FechaCaja As String)
    'BUSCO EN LOS RECIBOS
    'BUSCO LOS CHEQUES QUE SALIERON POR LA COBRANZA
    'AQUI ME FIJO EN LA FECHA EN LA CUAL REALIZO
    'LA LIQUIDACION A LAS REPRESENTADAS PARA SACAR LOS FONDOS DE LA MISMA
    sql = "SELECT DR.BAN_CODINT, DR.CHE_NUMERO, CH.CHE_IMPORT"
    sql = sql & " FROM RECIBO_CLIENTE R, DETALLE_RECIBO_CLIENTE DR, CHEQUE CH"
    sql = sql & " WHERE R.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
   ' sql = sql & " AND R.REP_CODIGO=DR.REP_CODIGO"
    sql = sql & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    sql = sql & " AND DR.BAN_CODINT=CH.BAN_CODINT"
    sql = sql & " AND DR.CHE_NUMERO=CH.CHE_NUMERO"
    'sql = sql & " AND R.REP_CODIGO=" & XN(CStr(Representada))
    sql = sql & " AND R.EST_CODIGO=3"
    sql = sql & " AND R.REC_LISTADO IS NOT NULL"
    'sql = sql & " AND R.REC_FECLIQUI IS NOT NULL"
    sql = sql & " AND R.REC_FECLIQUI >" & XDQ(FechaCaja)
    sql = sql & " AND R.REC_FECLIQUI <=" & XDQ(Fecha1.Value)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdDep.TextMatrix(23, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(23, 1)) + CDbl(rec!che_import)))
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub


Private Sub SUMO_EGRESOS()
    'TOTAL MONEDAS
    GrdDep.TextMatrix(29, 1) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(25, 1)) + CDbl(GrdDep.TextMatrix(26, 1)) + CDbl(GrdDep.TextMatrix(27, 1)) + CDbl(GrdDep.TextMatrix(28, 1))))
    'TOTAL EGRESOS
    GrdDep.TextMatrix(29, 4) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(29, 1)) + CDbl(GrdDep.TextMatrix(23, 1))))
End Sub

Private Sub SUMO_SALDOFINAL()
    'TOTAL CHEQUES
    GrdDep.TextMatrix(3, 4) = Valido_Importe(CStr((CDbl(GrdDep.TextMatrix(3, 1)) + CDbl(GrdDep.TextMatrix(14, 1))) - CDbl(GrdDep.TextMatrix(23, 1))))
    'TOTAL PESOS
    GrdDep.TextMatrix(5, 4) = Valido_Importe(CStr((CDbl(GrdDep.TextMatrix(5, 1)) + CDbl(GrdDep.TextMatrix(16, 1))) - CDbl(GrdDep.TextMatrix(25, 1))))
    'TOTAL LECOP NACION
    GrdDep.TextMatrix(6, 4) = Valido_Importe(CStr((CDbl(GrdDep.TextMatrix(6, 1)) + CDbl(GrdDep.TextMatrix(17, 1))) - CDbl(GrdDep.TextMatrix(26, 1))))
    'TOTAL LECOP CORDOBA
    GrdDep.TextMatrix(7, 4) = Valido_Importe(CStr((CDbl(GrdDep.TextMatrix(7, 1)) + CDbl(GrdDep.TextMatrix(18, 1))) - CDbl(GrdDep.TextMatrix(27, 1))))
    'TOTAL OTRAS MONEDAS
    GrdDep.TextMatrix(8, 4) = Valido_Importe(CStr((CDbl(GrdDep.TextMatrix(8, 1)) + CDbl(GrdDep.TextMatrix(19, 1))) - CDbl(GrdDep.TextMatrix(28, 1))))
    
    'TOTAL MONEDAS
    GrdDep.TextMatrix(9, 4) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(5, 4)) + CDbl(GrdDep.TextMatrix(6, 4)) + CDbl(GrdDep.TextMatrix(7, 4)) + CDbl(GrdDep.TextMatrix(8, 4))))
    'TOTAL SALDO FINAL
    GrdDep.TextMatrix(10, 4) = Valido_Importe(CStr(CDbl(GrdDep.TextMatrix(3, 4)) + CDbl(GrdDep.TextMatrix(9, 4))))
End Sub

Private Sub cmdGrabar_Click()
    If MsgBox("Confirma el cierre de caja?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    Set rec = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    On Error GoTo CLAVOSE
    DBConn.BeginTrans
    
    sql = "SELECT * FROM CAJA "
    sql = sql & " WHERE CAJA_FECHA = " & XDQ(Fecha1.Value)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
            MsgBox "El cierre de caja ya ha sido cargado!", vbInformation, TIT_MSGBOX
            rec.Close
            Exit Sub
    Else
        'SALDO INICIAL
        sql = "INSERT INTO CAJA (CAJA_FECHA,CAJA_SALDOIF,CAJA_SALDO_CHEQUES,CAJA_SALDO_PESOS,"
        sql = sql & "CAJA_SALDO_LN,CAJA_SALDO_LC,CAJA_SALDO_OTROS)"
        sql = sql & " VALUES ("
        sql = sql & XDQ(Fecha1.Value) & ","
        sql = sql & "'I',"
        sql = sql & XN(GrdDep.TextMatrix(3, 1)) & ","
        sql = sql & XN(GrdDep.TextMatrix(5, 1)) & ","
        sql = sql & XN(GrdDep.TextMatrix(6, 1)) & ","
        sql = sql & XN(GrdDep.TextMatrix(7, 1)) & ","
        sql = sql & XN(GrdDep.TextMatrix(8, 1)) & ")"
        DBConn.Execute sql
        'SALDO FINAL
        sql = "INSERT INTO CAJA (CAJA_FECHA,CAJA_SALDOIF,CAJA_SALDO_CHEQUES,CAJA_SALDO_PESOS,"
        sql = sql & "CAJA_SALDO_LN,CAJA_SALDO_LC,CAJA_SALDO_OTROS)"
        sql = sql & " VALUES ("
        sql = sql & XDQ(Fecha1.Value) & ","
        sql = sql & "'F',"
        sql = sql & XN(GrdDep.TextMatrix(3, 4)) & ","
        sql = sql & XN(GrdDep.TextMatrix(5, 4)) & ","
        sql = sql & XN(GrdDep.TextMatrix(6, 4)) & ","
        sql = sql & XN(GrdDep.TextMatrix(7, 4)) & ","
        sql = sql & XN(GrdDep.TextMatrix(8, 4)) & ")"
        DBConn.Execute sql
    End If
    rec.Close
    DBConn.CommitTrans
    lblEstado.Caption = ""
    MsgBox "La Rendición de Caja se ha grabado con éxito !", vbInformation, TIT_MSGBOX
    cmdGrabar.Enabled = False
    cmdImprimir_Click
    
    Screen.MousePointer = vbNormal
    Exit Sub
    
CLAVOSE:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub cmdImprimir_Click()
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    For I = 0 To 28
        Rep.Formulas(I) = ""
    Next
    'SALDO INICIAL
    Rep.Formulas(0) = "TCH1= '" & GrdDep.TextMatrix(3, 1) & "'"
    Rep.Formulas(1) = "TP1= '" & GrdDep.TextMatrix(5, 1) & "'"
    Rep.Formulas(2) = "TLN1= '" & GrdDep.TextMatrix(6, 1) & "'"
    Rep.Formulas(3) = "TLC1= '" & GrdDep.TextMatrix(7, 1) & "'"
    Rep.Formulas(4) = "TO1= '" & GrdDep.TextMatrix(8, 1) & "'"
    Rep.Formulas(5) = "TM1= '" & GrdDep.TextMatrix(9, 1) & "'"
    Rep.Formulas(6) = "TSI= '" & GrdDep.TextMatrix(10, 1) & "'"
    'SALDO FINAL
    Rep.Formulas(7) = "TCH2= '" & GrdDep.TextMatrix(3, 4) & "'"
    Rep.Formulas(8) = "TP2= '" & GrdDep.TextMatrix(5, 4) & "'"
    Rep.Formulas(9) = "TLN2= '" & GrdDep.TextMatrix(6, 4) & "'"
    Rep.Formulas(10) = "TLC2= '" & GrdDep.TextMatrix(7, 4) & "'"
    Rep.Formulas(11) = "TO2= '" & GrdDep.TextMatrix(8, 4) & "'"
    Rep.Formulas(12) = "TM2= '" & GrdDep.TextMatrix(9, 4) & "'"
    Rep.Formulas(13) = "TSF= '" & GrdDep.TextMatrix(10, 4) & "'"
    'TOTAL INGESOS
    Rep.Formulas(14) = "TCH3= '" & GrdDep.TextMatrix(14, 1) & "'"
    Rep.Formulas(15) = "TP3= '" & GrdDep.TextMatrix(16, 1) & "'"
    Rep.Formulas(16) = "TLN3= '" & GrdDep.TextMatrix(17, 1) & "'"
    Rep.Formulas(17) = "TLC3= '" & GrdDep.TextMatrix(18, 1) & "'"
    Rep.Formulas(18) = "TO3= '" & GrdDep.TextMatrix(19, 1) & "'"
    Rep.Formulas(19) = "TM3= '" & GrdDep.TextMatrix(20, 1) & "'"
    Rep.Formulas(20) = "TI= '" & GrdDep.TextMatrix(20, 4) & "'"
    'TOTAL EGRESOS
    Rep.Formulas(21) = "TCH4= '" & GrdDep.TextMatrix(23, 1) & "'"
    Rep.Formulas(22) = "TP4= '" & GrdDep.TextMatrix(25, 1) & "'"
    Rep.Formulas(23) = "TLN4= '" & GrdDep.TextMatrix(26, 1) & "'"
    Rep.Formulas(24) = "TLC4= '" & GrdDep.TextMatrix(27, 1) & "'"
    Rep.Formulas(25) = "TO4= '" & GrdDep.TextMatrix(28, 1) & "'"
    Rep.Formulas(26) = "TM4= '" & GrdDep.TextMatrix(29, 1) & "'"
    Rep.Formulas(27) = "TE= '" & GrdDep.TextMatrix(29, 4) & "'"
    'FECHA
    'dddd, dd' de 'MMMM' de 'aaaa
    Rep.Formulas(28) = "FECHA= '" & Format(Fecha1.Value, "dddd") & ", " & Format(Fecha1.Value, "dd") & _
                                " de " & Format(Fecha1.Value, "mmmm") & " de " & Format(Fecha1.Value, "yyyy") & "'"
    
    Rep.WindowTitle = "Listado de Caja Diaria"
    Rep.ReportFileName = DRIVE & DirReport & "cierre_caja.rpt"

    
    Rep.Destination = crptToWindow
    Rep.Action = 1
     
    lblEstado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
    cmdGrabar.Enabled = False
    ARMO_GRIYA
    If Me.Visible Then Fecha1.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmCierreCaja = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    'sql = "SELECT REP_CODIGO FROM PARAMETROS"
    'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    'If rec.EOF = False Then
    '    Representada = rec!REP_CODIGO
    'End If
    'rec.Close
    
    Fecha1.Value = Date
    'Call Centrar_pantalla(Me)
    Me.Top = 70
    Me.Left = 400
    GrdDep.RowHeightMin = 250
    GrdDep.RowHeightMin = 250
    lblEstado.Caption = ""
    CmdNuevo_Click
End Sub

Private Sub ARMO_GRIYA()
    GrdDep.FixedRows = 2
    GrdDep.Rows = 1
    GrdDep.FixedCols = 0
    GrdDep.FormatString = "<CIERRE CAJA - TOTALCAR| |||<"
    GrdDep.ColWidth(0) = 2400   'CONCEPTOS
    GrdDep.ColWidth(1) = 1100   'IMPORTES
    GrdDep.ColWidth(2) = 500
    GrdDep.ColWidth(3) = 2400
    GrdDep.ColWidth(4) = 1100
    'GrdDep.AddItem ""
    GrdDep.AddItem " SALDO INICIAL" & Chr(9) & "" & Chr(9) & "" & Chr(9) & " SALDO FINAL"
    GrdDep.AddItem " CHEQUES" & Chr(9) & "" & Chr(9) & "" & Chr(9) & " CHEQUES"
    GrdDep.AddItem "   Total Cheques" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "   Total Cheques"
    GrdDep.AddItem " MONEDAS" & Chr(9) & "" & Chr(9) & "" & Chr(9) & " MONEDAS"
    GrdDep.AddItem "   Total Pesos" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "   Total Pesos"
    GrdDep.AddItem "   Total Lecop Nación" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "   Total Lecop Nación"
    GrdDep.AddItem "   Total Lecop Córdoba" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "   Total Lecop Córdoba"
    GrdDep.AddItem "   Total Otros" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "   Total Otros"
    GrdDep.AddItem " Total Monedas" & Chr(9) & "" & Chr(9) & "" & Chr(9) & " Total Monedas"
    GrdDep.AddItem " TOTAL SALDO INICIAL" & Chr(9) & "" & Chr(9) & "" & Chr(9) & " TOTAL SALDO FINAL"
    
    GrdDep.AddItem " " 'GrdDep.AddItem "CONCEPTOS"
    GrdDep.AddItem "INGRESOS POR COBROS"
    GrdDep.AddItem " CHEQUES"
    GrdDep.AddItem "   Total Cheques"
    GrdDep.AddItem " MONEDAS"
    GrdDep.AddItem "   Total Pesos"
    GrdDep.AddItem "   Total Lecop Nación"
    GrdDep.AddItem "   Total Lecop Córdoba"
    GrdDep.AddItem "   Total Otros"
    GrdDep.AddItem " Total Monedas" & Chr(9) & "" & Chr(9) & "" & Chr(9) & " TOTAL INGRESOS"
    GrdDep.AddItem "EGRESOS POR PAGOS"
    GrdDep.AddItem " CHEQUES"
    GrdDep.AddItem "   Total Cheques"
    'GrdDep.AddItem ""
    GrdDep.AddItem " MONEDAS"
    GrdDep.AddItem "   Total Pesos"
    GrdDep.AddItem "   Total Lecop Nación"
    GrdDep.AddItem "   Total Lecop Córdoba"
    GrdDep.AddItem "   Total Otros"
    GrdDep.AddItem " Total Monedas" & Chr(9) & "" & Chr(9) & "" & Chr(9) & " TOTAL EGRESOS"
    GrdDep.AddItem " "
    
    'PINTA TODA LA COLUMNA 2
    GrdDep.Col = 2
    For a = 0 To GrdDep.Rows - 1
        GrdDep.row = a
        GrdDep.CellBackColor = &HE0E0E0 'GRIS
    Next
    'PINTA LA FILA 1
    'CambiaColorAFilaDeGrilla
    GrdDep.row = 0
    For a = 0 To GrdDep.Cols - 1
        GrdDep.Col = a
        GrdDep.CellBackColor = &H808080   '&HE0E0E0 'GRIS
        GrdDep.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdDep.CellFontBold = True
    Next
    'TITULOS SALDOS
    GrdDep.row = 1
    For a = 0 To GrdDep.Cols - 1
        GrdDep.Col = a
        GrdDep.CellBackColor = &HE0E0E0 'GRIS
        GrdDep.CellFontBold = True
    Next
    'TITULOS CHEQUES
    GrdDep.row = 2
    For a = 0 To GrdDep.Cols - 1
        GrdDep.Col = a
        GrdDep.CellBackColor = &HE0E0E0 'GRIS
        GrdDep.CellFontBold = True
    Next
    'TITULOS MONEDAS
    GrdDep.row = 4
    For a = 0 To GrdDep.Cols - 1
        GrdDep.Col = a
        GrdDep.CellBackColor = &HE0E0E0 'GRIS
        GrdDep.CellFontBold = True
    Next
    'TITULOS TOTAL SALDOS
    GrdDep.row = 10
    For a = 0 To GrdDep.Cols - 1
        GrdDep.Col = a
        GrdDep.CellBackColor = &HE0E0E0 'GRIS
        GrdDep.CellFontBold = True
    Next
    For I = 3 To 4
        GrdDep.Col = I
        For a = 11 To 30
            GrdDep.row = a
            GrdDep.CellBackColor = &HE0E0E0 'GRIS
        Next
    Next
    
    'TITULOS CONCEPTOS
    For I = 11 To 13
        GrdDep.row = I
        For a = 0 To GrdDep.Cols - 1
            GrdDep.Col = a
            GrdDep.CellBackColor = &HE0E0E0 'GRIS
            GrdDep.CellFontBold = True
        Next
    Next
    'Titulo MONEDAS
    GrdDep.row = 15
    For a = 0 To GrdDep.Cols - 1
        GrdDep.Col = a
        GrdDep.CellBackColor = &HE0E0E0 'GRIS
        GrdDep.CellFontBold = True
    Next
    'TITULOS
    For I = 21 To 22
        GrdDep.row = I
        For a = 0 To GrdDep.Cols - 1
            GrdDep.Col = a
            GrdDep.CellBackColor = &HE0E0E0 'GRIS
            GrdDep.CellFontBold = True
        Next
    Next
    'Titulo MONEDAS
    GrdDep.row = 24
    For a = 0 To GrdDep.Cols - 1
        GrdDep.Col = a
        GrdDep.CellBackColor = &HE0E0E0 'GRIS
        GrdDep.CellFontBold = True
    Next
    GrdDep.row = 20
    GrdDep.Col = 3
    GrdDep.CellFontBold = True
    GrdDep.row = 29
    GrdDep.Col = 3
    GrdDep.CellFontBold = True
    'PARA LOS TOTALES EN AMARILLO
    GrdDep.Col = 1
    GrdDep.row = 3
    GrdDep.CellBackColor = &HC0FFFF 'AMARILIO
    GrdDep.TextMatrix(GrdDep.row, GrdDep.Col) = "0,00"
    For a = 5 To 10
        GrdDep.row = a
        GrdDep.CellBackColor = &HC0FFFF 'AMARILIO
        GrdDep.TextMatrix(a, GrdDep.Col) = "0,00"
    Next
    GrdDep.Col = 4
    GrdDep.row = 3
    GrdDep.CellBackColor = &HC0FFFF 'AMARILIO
    GrdDep.TextMatrix(GrdDep.row, GrdDep.Col) = "0,00"
    For a = 5 To 10
        GrdDep.row = a
        GrdDep.CellBackColor = &HC0FFFF 'AMARILIO
        GrdDep.TextMatrix(a, GrdDep.Col) = "0,00"
    Next
    GrdDep.Col = 1
    GrdDep.row = 14
    GrdDep.CellBackColor = &HC0FFFF 'AMARILIO
    GrdDep.TextMatrix(GrdDep.row, GrdDep.Col) = "0,00"
    For a = 16 To 20
        GrdDep.row = a
        GrdDep.CellBackColor = &HC0FFFF 'AMARILIO
        GrdDep.TextMatrix(a, GrdDep.Col) = "0,00"
    Next
    GrdDep.Col = 4
    GrdDep.CellBackColor = &HC0FFFF 'AMARILIO
    GrdDep.CellFontBold = True
    GrdDep.TextMatrix(GrdDep.row, GrdDep.Col) = "0,00"
    GrdDep.Col = 1
    GrdDep.row = 23
    GrdDep.CellBackColor = &HC0FFFF 'AMARILIO
    GrdDep.TextMatrix(GrdDep.row, GrdDep.Col) = "0,00"
    For a = 25 To 29
        GrdDep.row = a
        GrdDep.CellBackColor = &HC0FFFF 'AMARILIO
        GrdDep.TextMatrix(a, GrdDep.Col) = "0,00"
    Next
    GrdDep.Col = 4
    GrdDep.CellBackColor = &HC0FFFF 'AMARILIO
    GrdDep.CellFontBold = True
    GrdDep.TextMatrix(GrdDep.row, GrdDep.Col) = "0,00"
    GrdDep.row = 30
    For a = 0 To GrdDep.Cols - 1
        GrdDep.Col = a
        GrdDep.CellBackColor = &HE0E0E0 'GRIS
    Next
    'PINTO LOS TOTALES DE COLOR MAS FUERTE
    GrdDep.Col = 0
    GrdDep.row = 10
    GrdDep.CellBackColor = &HFF8080
    GrdDep.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
    GrdDep.Col = 3
    GrdDep.CellBackColor = &HFF8080
    GrdDep.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
    GrdDep.row = 20
    GrdDep.CellBackColor = &HFF8080
    GrdDep.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
    GrdDep.row = 29
    GrdDep.CellBackColor = &HFF8080
    GrdDep.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
End Sub

