VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmLiquidacionCobranza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación de Cobranza"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSelec 
      Caption         =   "&Seleccionar todo"
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
      Left            =   105
      TabIndex        =   5
      Top             =   3690
      Width           =   2355
   End
   Begin VB.CommandButton CmdDeselec 
      Caption         =   "&Deseleccionar todo"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   3690
      Width           =   2355
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar Cobranza"
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
      Left            =   90
      TabIndex        =   1
      Top             =   795
      Width           =   4785
   End
   Begin MSFlexGridLib.MSFlexGrid grdLiquidacion 
      Height          =   2430
      Left            =   90
      TabIndex        =   2
      Top             =   1155
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   4286
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorSel    =   8388736
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.ComboBox cboRep 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   390
      Width           =   3300
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   4110
      Picture         =   "frmLiquidacionCobranza.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4140
      Width           =   840
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmLiquidacionCobranza.frx":030A
      Height          =   750
      Left            =   3225
      Picture         =   "frmLiquidacionCobranza.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4140
      Width           =   870
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
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
      Left            =   3555
      TabIndex        =   9
      Top             =   60
      Width           =   540
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Representada:"
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   450
      Width           =   1050
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
      Left            =   90
      TabIndex        =   7
      Top             =   4320
      Width           =   750
   End
End
Attribute VB_Name = "frmLiquidacionCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer

Private Sub CmdAceptar_Click()
    If MsgBox("Confirma Liquidación de Cobranza", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        lblEstado.Caption = "Actualizando..."
        If grdLiquidacion.Rows > 1 Then
            For I = 1 To grdLiquidacion.Rows - 1
                If grdLiquidacion.TextMatrix(I, 3) = "SI" Then
                    sql = "UPDATE RECIBO_CLIENTE"
                    sql = sql & " SET REC_FECLIQUI=" & XDQ(lblFecha.Caption)
                    sql = sql & " WHERE REP_CODIGO=" & XN(grdLiquidacion.TextMatrix(I, 4))
                    sql = sql & " AND REC_SUCURSAL=" & XN(Left(grdLiquidacion.TextMatrix(I, 1), 4))
                    sql = sql & " AND REC_NUMERO=" & XN(Right(grdLiquidacion.TextMatrix(I, 1), 8))
                    sql = sql & " AND REC_LISTADO=" & XN(grdLiquidacion.TextMatrix(I, 0))
                    DBConn.Execute sql
                End If
            Next
        End If
        lblEstado.Caption = ""
    End If
End Sub

Private Sub cmdBuscar_Click()
    Set rec = New ADODB.Recordset
    lblEstado.Caption = "Buscando..."
    sql = "SELECT REC_LISTADO, REC_NUMERO, REC_SUCURSAL, REC_FECHA"
    sql = sql & " FROM RECIBO_CLIENTE"
    sql = sql & " WHERE "
    'REP_CODIGO=" & XN(cboRep.ItemData(cboRep.ListIndex))
    sql = sql & " REC_LISTADO IS NOT NULL"
    sql = sql & " AND REC_FECLIQUI IS NULL"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        grdLiquidacion.Rows = 1
        Do While rec.EOF = False
            grdLiquidacion.AddItem rec!REC_LISTADO & Chr(9) & _
                            Format(rec!REC_SUCURSAL, "0000") & "-" & Format(rec!REC_NUMERO, "00000000") & Chr(9) & _
                            rec!REC_FECHA & Chr(9) & "" & Chr(9) & cboRep.ItemData(cboRep.ListIndex)
            rec.MoveNext
        Loop
        grdLiquidacion.SetFocus
    Else
        lblEstado.Caption = ""
        MsgBox "No hay liquidaciones de Cobranza pendientre por mandar", vbInformation, TIT_MSGBOX
        grdLiquidacion.Rows = 1
        grdLiquidacion.Rows = 2
        cboRep.SetFocus
    End If
    lblEstado.Caption = ""
    rec.Close
End Sub

Private Sub CmdDeselec_Click()
     For I = 1 To grdLiquidacion.Rows - 1
        grdLiquidacion.TextMatrix(I, 3) = ""
    Next
End Sub

Private Sub cmdSalir_Click()
    Set frmLiquidacionCobranza = Nothing
    Unload Me
End Sub

Private Sub CmdSelec_Click()
    For I = 1 To grdLiquidacion.Rows - 1
        grdLiquidacion.TextMatrix(I, 3) = "SI"
    Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    CargoComboRepresentada
    lblEstado.Caption = ""
    lblFecha.Caption = Date
    'CONFIGURO GRILLA
    grdLiquidacion.FormatString = "Cobranza|^Recibo|^Fecha|^Liquida|REPRESENTADA"
    grdLiquidacion.ColWidth(0) = 900
    grdLiquidacion.ColWidth(1) = 1300
    grdLiquidacion.ColWidth(2) = 1100
    grdLiquidacion.ColWidth(3) = 800
    grdLiquidacion.ColWidth(4) = 0
    grdLiquidacion.Rows = 2
End Sub

Private Sub CargoComboRepresentada()
    sql = "SELECT REP_RAZSOC,REP_CODIGO FROM REPRESENTADA ORDER BY REP_RAZSOC"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboRep.AddItem rec!REP_RAZSOC
            cboRep.ItemData(cboRep.NewIndex) = rec!REP_CODIGO
            rec.MoveNext
        Loop
        cboRep.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub grdLiquidacion_DblClick()
    If grdLiquidacion.Rows > 1 Then
        If grdLiquidacion.TextMatrix(grdLiquidacion.RowSel, 0) <> "" Then
            If grdLiquidacion.TextMatrix(grdLiquidacion.RowSel, 3) = "SI" Then
                grdLiquidacion.TextMatrix(grdLiquidacion.RowSel, 3) = ""
            Else
                grdLiquidacion.TextMatrix(grdLiquidacion.RowSel, 3) = "SI"
            End If
        End If
    End If
End Sub

Private Sub grdLiquidacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        grdLiquidacion_DblClick
    End If
End Sub
