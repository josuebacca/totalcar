VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmCargaChequesPropios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Cheques Propios"
   ClientHeight    =   4320
   ClientLeft      =   2535
   ClientTop       =   1005
   ClientWidth     =   7215
   Icon            =   "FrmCargaChequesPropios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "FrmCargaChequesPropios.frx":08CA
      Height          =   735
      Left            =   6180
      Picture         =   "FrmCargaChequesPropios.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3555
      Width           =   900
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      DisabledPicture =   "FrmCargaChequesPropios.frx":0EDE
      Height          =   735
      Left            =   5265
      Picture         =   "FrmCargaChequesPropios.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3555
      Width           =   900
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "FrmCargaChequesPropios.frx":14F2
      Height          =   735
      Left            =   4350
      Picture         =   "FrmCargaChequesPropios.frx":17FC
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3555
      Width           =   900
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "FrmCargaChequesPropios.frx":1B06
      Height          =   735
      Left            =   3435
      Picture         =   "FrmCargaChequesPropios.frx":1E10
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3555
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   3450
      Left            =   120
      TabIndex        =   14
      Top             =   15
      Width           =   6975
      Begin VB.ComboBox cboCtaBancaria 
         Height          =   315
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1110
         Width           =   2100
      End
      Begin VB.ComboBox cboBanco 
         Height          =   315
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   735
         Width           =   5040
      End
      Begin VB.TextBox TxtCheNumero 
         Height          =   315
         Left            =   5115
         MaxLength       =   8
         TabIndex        =   1
         Top             =   270
         Width           =   1380
      End
      Begin VB.TextBox TxtCheMotivo 
         Height          =   315
         Left            =   1455
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1860
         Width           =   5040
      End
      Begin VB.TextBox TxtCheNombre 
         Height          =   315
         Left            =   1455
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1485
         Width           =   5040
      End
      Begin VB.TextBox TxtCheObserv 
         Height          =   315
         Left            =   1455
         MaxLength       =   60
         TabIndex        =   9
         Top             =   2985
         Width           =   5040
      End
      Begin VB.TextBox TxtCheImport 
         Height          =   315
         Left            =   1455
         TabIndex        =   8
         Top             =   2610
         Width           =   1400
      End
      Begin MSComCtl2.DTPicker TxtCheFecEnt 
         Height          =   315
         Left            =   1455
         TabIndex        =   0
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   61014017
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker TxtCheFecEmi 
         Height          =   315
         Left            =   1455
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   61014017
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker TxtCheFecVto 
         Height          =   315
         Left            =   5160
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   61014017
         CurrentDate     =   41098
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cuenta:"
         Height          =   195
         Index           =   4
         Left            =   495
         TabIndex        =   25
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   840
         TabIndex        =   24
         Top             =   765
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Concepto:"
         Height          =   195
         Index           =   10
         Left            =   615
         TabIndex        =   22
         Top             =   1905
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Responsable:"
         Height          =   195
         Index           =   9
         Left            =   375
         TabIndex        =   21
         Top             =   1530
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cheque:"
         Height          =   195
         Index           =   7
         Left            =   4140
         TabIndex        =   20
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   19
         Top             =   3015
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   0
         Left            =   855
         TabIndex        =   18
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   17
         Top             =   2670
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  Emisión:"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   16
         Top             =   2250
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Pago:"
         Height          =   195
         Index           =   3
         Left            =   3930
         TabIndex        =   15
         Top             =   2280
         Width           =   1140
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
      Height          =   240
      Left            =   195
      TabIndex        =   23
      Top             =   3765
      Width           =   750
   End
End
Attribute VB_Name = "FrmCargaChequesPropios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset
Dim Rec1 As ADODB.Recordset
Dim sql As String
Dim ImporteCheque As String

Function Validar() As Boolean
   If Trim(TxtCheNumero.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Número de Cheque.", 16, TIT_MSGBOX
        TxtCheNumero.SetFocus
        Exit Function
        
   ElseIf cboBanco.ListIndex = -1 Then
        Validar = False
        MsgBox "Ingrese el Banco.", 16, TIT_MSGBOX
        cboBanco.SetFocus
        Exit Function
                 
   ElseIf cboCtaBancaria.ListIndex = -1 Then
        Validar = False
        MsgBox "Ingrese la Cta Bancaria.", 16, TIT_MSGBOX
        cboCtaBancaria.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheNombre.Text) = "" Then
        Validar = False
        MsgBox "Debe ingresar la Persona responsable.", 16, TIT_MSGBOX
        TxtCheNombre.SetFocus
        Exit Function
   
   ElseIf Trim(TxtCheMotivo.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Concepto del Cheque.", 16, TIT_MSGBOX
        TxtCheMotivo.SetFocus
        Exit Function
        
   ElseIf IsNull(TxtCheFecEmi.Value) Then
        Validar = False
        MsgBox "Ingrese la Fecha de Emisión.", 16, TIT_MSGBOX
        TxtCheFecEmi.SetFocus
        Exit Function
        
   ElseIf IsNull(TxtCheFecVto.Value) Then
        Validar = False
        MsgBox "Ingrese la Fecha de Vencimiento.", 16, TIT_MSGBOX
        TxtCheFecVto.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheImport.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Importe del Cheque.", 16, TIT_MSGBOX
        TxtCheImport.SetFocus
        Exit Function
        
   End If
   
   Validar = True
End Function


Private Sub CboBanco_LostFocus()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Dim MtrObjetos As Variant
        
    If cboBanco.ListIndex <> -1 Then
    
       'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE_PROPIO " & _
              " WHERE CHEP_NUMERO = " & XS(TxtCheNumero.Text) & _
                " AND BAN_CODINT = " & XN(cboBanco.ItemData(cboBanco.ListIndex))
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then 'EXITE
            Me.TxtCheFecEnt.Value = rec!CHEP_FECENT
            Me.TxtCheNumero.Text = Trim(rec!CHEP_NUMERO)
            
            Me.TxtCheNombre.Text = ChkNull(rec!CHEP_NOMBRE)
            Me.TxtCheMotivo.Text = rec!CHEP_MOTIVO
            Me.TxtCheFecEmi.Value = rec!CHEP_FECEMI
            Me.TxtCheFecVto.Value = rec!CHEP_FECVTO
            Me.TxtCheImport.Text = Valido_Importe(rec!CHEP_IMPORT)
            ImporteCheque = rec!CHEP_IMPORT
            Me.TxtCheObserv.Text = ChkNull(rec!CHEP_OBSERV)
            Call CargoCtaBancaria(CStr(cboBanco.ItemData(cboBanco.ListIndex)))
            Call BuscaProx(Trim(rec!CTA_NROCTA), cboCtaBancaria)
            TxtCheNumero.Enabled = False
            cboBanco.Enabled = False
            MtrObjetos = Array(TxtCheNumero, cboBanco)
            Call CambiarColor(MtrObjetos, 2, &H80000018, "D")
        Else
            
           rec.Close
           Call CargoCtaBancaria(CStr(cboBanco.ItemData(cboBanco.ListIndex)))
           Exit Sub
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub CargoCtaBancaria(Banco As String)
    Set Rec1 = New ADODB.Recordset
    cboCtaBancaria.Clear
    sql = "SELECT CTA_NROCTA FROM CTA_BANCARIA"
    sql = sql & " WHERE BAN_CODINT=" & XN(Banco)
    sql = sql & " AND CTA_FECCIE IS NULL"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
     Do While Rec1.EOF = False
         cboCtaBancaria.AddItem Trim(Rec1!CTA_NROCTA)
         Rec1.MoveNext
     Loop
     cboCtaBancaria.ListIndex = 0
    End If
    Rec1.Close
End Sub

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    
    If Trim(TxtCheNumero.Text) <> "" And Me.cboBanco.ListIndex <> -1 Then
        resp = MsgBox("Seguro desea eliminar el Cheque Nº: " & Trim(Me.TxtCheNumero.Text) & "? ", 36, TIT_MSGBOX)
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Borrando..."
        DBConn.BeginTrans
        
        'ACTUALIZO EL SALDO DE LA CTA-BANCARIA
'        If ImporteCheque <> "" Then
'            sql = "UPDATE CTA_BANCARIA"
'            sql = sql & " SET CTA_SALACT = CTA_SALACT + " & XN(ImporteCheque)
'            sql = sql & " WHERE"
'            sql = sql & " CTA_NROCTA=" & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
'            sql = sql & " AND BAN_CODINT=" & XN(cboBanco.ItemData(cboBanco.ListIndex))
'            DBConn.Execute sql
'        End If
        
        DBConn.Execute "DELETE FROM CHEQUE_PROPIO_ESTADO WHERE CHEP_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.cboBanco.ItemData(cboBanco.ListIndex))
                       
        DBConn.Execute "DELETE FROM CHEQUE_PROPIO WHERE CHEP_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.cboBanco.ItemData(cboBanco.ListIndex))
        
        Screen.MousePointer = vbNormal
        lblEstado.Caption = ""
        DBConn.CommitTrans
        CmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    DBConn.RollbackTrans
    If rec.State = 1 Then rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX

End Sub

Private Sub cmdGrabar_Click()
    
  If Validar = True Then
  
    On Error GoTo CLAVOSE
    
    DBConn.BeginTrans
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    Me.Refresh
    
    sql = "SELECT * FROM CHEQUE_PROPIO WHERE CHEP_NUMERO = " & XS(TxtCheNumero.Text)
    sql = sql & " AND BAN_CODINT = " & XN(cboBanco.ItemData(cboBanco.ListIndex))
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = True Then
         sql = "INSERT INTO CHEQUE_PROPIO(CHEP_NUMERO,BAN_CODINT,CHEP_NOMBRE,CHEP_IMPORT,CHEP_FECEMI,"
         sql = sql & "CHEP_FECVTO,CHEP_FECENT,CHEP_MOTIVO,CHEP_OBSERV,CTA_NROCTA)"
         sql = sql & " VALUES (" & XS(Me.TxtCheNumero.Text) & ","
         sql = sql & XN(cboBanco.ItemData(cboBanco.ListIndex)) & ","
         sql = sql & XS(Me.TxtCheNombre.Text) & ","
         sql = sql & XN(Me.TxtCheImport.Text) & ","
         sql = sql & XDQ(Me.TxtCheFecEmi.Value) & ","
         sql = sql & XDQ(Me.TxtCheFecVto.Value) & ","
         sql = sql & XDQ(Me.TxtCheFecEnt.Value) & ","
         sql = sql & XS(Me.TxtCheMotivo.Text) & ","
         sql = sql & XS(Me.TxtCheObserv.Text) & ","
         sql = sql & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex)) & ")"
         DBConn.Execute sql
    Else
         sql = "UPDATE CHEQUE_PROPIO SET CHEP_NOMBRE = " & XS(Me.TxtCheNombre.Text)
         sql = sql & ",CHEP_IMPORT = " & XN(Me.TxtCheImport.Text)
         sql = sql & ",CHEP_FECEMI =" & XDQ(Me.TxtCheFecEmi.Value)
         sql = sql & ",CHEP_FECVTO =" & XDQ(Me.TxtCheFecVto.Value)
         sql = sql & ",CHEP_FECENT = " & XDQ(Me.TxtCheFecEnt.Value)
         sql = sql & ",CHEP_MOTIVO = " & XS(Me.TxtCheMotivo.Text)
         sql = sql & ",CHEP_OBSERV = " & XS(Me.TxtCheObserv.Text)
         sql = sql & ",CTA_NROCTA= " & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
         sql = sql & " WHERE CHEP_NUMERO = " & XS(Me.TxtCheNumero.Text)
         sql = sql & " AND BAN_CODINT = " & XN(cboBanco.ItemData(cboBanco.ListIndex))
         DBConn.Execute sql
    End If
    rec.Close
     
    'Insert en la Tabla de Estados de Cheques
    sql = "INSERT INTO CHEQUE_PROPIO_ESTADO (CHEP_NUMERO,BAN_CODINT,ECH_CODIGO,CPES_FECHA,CPES_DESCRI)"
    sql = sql & " VALUES ("
    sql = sql & XS(Me.TxtCheNumero.Text) & ","
    sql = sql & XN(cboBanco.ItemData(cboBanco.ListIndex)) & "," & XN(8) & ","
    sql = sql & XDQ(Date) & ",'CHEQUE LIBRADO')"
    DBConn.Execute sql
    
    'ACTUALIZO EL SALDO DE LA CTA-BANCARIA
'    If ImporteCheque <> "" Then
'        sql = "UPDATE CTA_BANCARIA"
'        sql = sql & " SET CTA_SALACT = CTA_SALACT + " & XN(ImporteCheque)
'        sql = sql & " WHERE"
'        sql = sql & " CTA_NROCTA=" & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
'        sql = sql & " AND BAN_CODINT=" & XN(cboBanco.ItemData(cboBanco.ListIndex))
'        DBConn.Execute sql
'    End If
'        sql = "UPDATE CTA_BANCARIA"
'        sql = sql & " SET CTA_SALACT = CTA_SALACT - " & XN(TxtCheImport)
'        sql = sql & " WHERE"
'        sql = sql & " CTA_NROCTA=" & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
'        sql = sql & " AND BAN_CODINT=" & XN(cboBanco.ItemData(cboBanco.ListIndex))
'        DBConn.Execute sql
    
    '************* PREGUNTAR POR SI DESEA IMPRIMIR ***************
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.CommitTrans
    CmdNuevo_Click
 End If
 Exit Sub
      
CLAVOSE:
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub CmdNuevo_Click()
    Me.TxtCheFecEnt.Value = Null
    Me.TxtCheNumero.Enabled = True
    Me.cboBanco.Enabled = True
    cboCtaBancaria.Clear
    Me.TxtCheNombre.Enabled = True
    MtrObjetos = Array(TxtCheNumero, cboBanco)
    Call CambiarColor(MtrObjetos, 2, &H80000005, "E")
    Me.TxtCheNumero.Text = ""
    Me.cboBanco.ListIndex = 0
    Me.TxtCheNombre.Text = ""
    Me.TxtCheMotivo.Text = ""
    Me.TxtCheFecEmi.Value = Null
    Me.TxtCheFecVto.Value = Null
    Me.TxtCheImport.Text = ""
    Me.TxtCheObserv.Text = ""
    ImporteCheque = ""
    Me.TxtCheFecEnt.SetFocus
    'TxtCheNombre.ForeColor = &H80000005
    Me.TxtCheNombre.Text = "HORACIO DANIEL GIRAUDO"
    lblEstado.Caption = ""
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set FrmCargaChequesPropios = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    TxtCheFecEnt.Value = Date
    lblEstado.Caption = ""
    ImporteCheque = ""
    'CARGO LOS BANCON DONDE TIENEN CUENTAS
    CargoBanco
    cboCtaBancaria.Clear
    Me.TxtCheNombre.Text = "HORACIO DANIEL GIRAUDO"
End Sub

Private Sub CargoBanco()
    sql = "SELECT B.BAN_DESCRI, B.BAN_CODINT"
    sql = sql & " FROM BANCO B, CTA_BANCARIA CB"
    sql = sql & " WHERE B.BAN_CODINT=CB.BAN_CODINT"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboBanco.AddItem Trim(rec!BAN_DESCRI)
            cboBanco.ItemData(cboBanco.NewIndex) = Trim(rec!BAN_CODINT)
            rec.MoveNext
        Loop
        cboBanco.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub TxtCheFecVto_LostFocus()
 If Not IsNull(Me.TxtCheFecEmi.Value) And Not IsNull(Me.TxtCheFecVto.Value) Then
 
   If IsDate(TxtCheFecEmi.Value) And IsDate(TxtCheFecVto.Value) Then
    
    If CVDate(TxtCheFecEmi.Value) > CVDate(TxtCheFecVto.Value) Then
        MsgBox "La Fecha de Vencimiento no puede ser anterior a la Fecha de Emisión del Cheque.! ", 16, TIT_MSGBOX
        Me.TxtCheFecVto.Value = Null
        Me.TxtCheFecVto.SetFocus
    Else
       If Me.TxtCheImport.Enabled = False Then 'PAGO EN CUOTAS
            Tasa = Trim(FrmComprobante.txtPmt_Tasa.Text)
            'Saco la Cantidad de Días del Cheque
            Cant_Dias = DateDiff("d", FrmComprobante.TxtFechaComprobante.Text, Me.TxtCheFecVto.Value)
            
            'Cálculo de Interes a Fecha del Cheque
            TxtCheImport.Text = Format(TxtCheImport.Text + (CDbl(TxtCheImport.Text) * CDbl(Chk0(Cant_Dias * Tasa)) / 100), "$ ##,##0.00")
            
        End If
    End If
  End If
 End If
 
End Sub

Private Sub TxtCheImport_GotFocus()
    SelecTexto TxtCheImport
End Sub

Private Sub TxtCheImport_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(TxtCheImport.Text, KeyAscii)
End Sub

Private Sub TxtCheImport_LostFocus()
   If Trim(TxtCheImport.Text) <> "" Then TxtCheImport.Text = Valido_Importe(TxtCheImport)
    
End Sub

Private Sub TxtCheMotivo_GotFocus()
    SelecTexto TxtCheMotivo
End Sub

Private Sub TxtCheNombre_GotFocus()
    SelecTexto TxtCheNombre
End Sub

Private Sub TxtCheNombre_LostFocus()
   If Me.TxtCheNombre.Text <> "" Then
      Me.TxtCheMotivo.SetFocus
   End If
End Sub

Private Sub TxtCheNumero_GotFocus()
    SelecTexto TxtCheNumero
End Sub

Private Sub TxtCheNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCheFecEnt_LostFocus()
    If IsNull(TxtCheFecEnt.Value) Then
        TxtCheFecEnt.Value = Format(Date, "dd/mm/yyyy")
    End If
End Sub

Private Sub TxtCheMotivo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCheNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCheNumero_LostFocus()
   If Len(TxtCheNumero.Text) < 8 Then TxtCheNumero.Text = CompletarConCeros(TxtCheNumero.Text, 8)
End Sub

Private Sub TxtCheObserv_GotFocus()
    SelecTexto TxtCheObserv
End Sub

Private Sub TxtCheObserv_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

