VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmCargaCheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Cheques de Terceros"
   ClientHeight    =   4635
   ClientLeft      =   2535
   ClientTop       =   1005
   ClientWidth     =   7215
   Icon            =   "FrmCargaCheques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "FrmCargaCheques.frx":08CA
      Height          =   735
      Left            =   6180
      Picture         =   "FrmCargaCheques.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   900
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      DisabledPicture =   "FrmCargaCheques.frx":0EDE
      Height          =   735
      Left            =   5265
      Picture         =   "FrmCargaCheques.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3840
      Width           =   900
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "FrmCargaCheques.frx":14F2
      Height          =   735
      Left            =   4350
      Picture         =   "FrmCargaCheques.frx":17FC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3840
      Width           =   900
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "FrmCargaCheques.frx":1B06
      Height          =   735
      Left            =   3435
      Picture         =   "FrmCargaCheques.frx":1E10
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   3750
      Left            =   120
      TabIndex        =   18
      Top             =   15
      Width           =   6975
      Begin VB.CommandButton cmdBuscaCheque 
         Height          =   315
         Left            =   6120
         MaskColor       =   &H000000FF&
         Picture         =   "FrmCargaCheques.frx":211A
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Buscar Cheques"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.TextBox TxtCheNumero 
         Height          =   315
         Left            =   4605
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1380
      End
      Begin VB.Frame Frame2 
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   465
         TabIndex        =   27
         Top             =   615
         Width           =   6120
         Begin VB.CommandButton CmdBanco 
            DisabledPicture =   "FrmCargaCheques.frx":2424
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5655
            Picture         =   "FrmCargaCheques.frx":272E
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   270
            Width           =   375
         End
         Begin VB.TextBox TxtSUCURSAL 
            Height          =   285
            Left            =   3525
            MaxLength       =   3
            TabIndex        =   4
            Top             =   285
            Width           =   450
         End
         Begin VB.TextBox TxtBANCO 
            Height          =   285
            Left            =   780
            MaxLength       =   3
            TabIndex        =   2
            Top             =   285
            Width           =   450
         End
         Begin VB.TextBox TxtLOCALIDAD 
            Height          =   285
            Left            =   2175
            MaxLength       =   3
            TabIndex        =   3
            Top             =   285
            Width           =   450
         End
         Begin VB.TextBox TxtCODIGO 
            Height          =   285
            Left            =   4755
            MaxLength       =   6
            TabIndex        =   5
            Top             =   270
            Width           =   765
         End
         Begin VB.TextBox TxtCodInt 
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   5370
            TabIndex        =   32
            Top             =   660
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox TxtBanDescri 
            BackColor       =   &H00C0C0C0&
            Height          =   330
            Left            =   210
            TabIndex        =   6
            Top             =   675
            Width           =   5820
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Localidad:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   1395
            TabIndex        =   31
            Top             =   315
            Width           =   735
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   10
            Left            =   225
            TabIndex        =   30
            Top             =   315
            Width           =   510
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   2820
            TabIndex        =   29
            Top             =   315
            Width           =   645
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
            Left            =   4125
            TabIndex        =   28
            Top             =   315
            Width           =   540
         End
      End
      Begin VB.TextBox TxtCheMotivo 
         Height          =   315
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   9
         Top             =   2160
         Width           =   5040
      End
      Begin VB.TextBox TxtCheNombre 
         Height          =   315
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1785
         Width           =   5040
      End
      Begin VB.TextBox TxtCheObserv 
         Height          =   315
         Left            =   1470
         MaxLength       =   60
         TabIndex        =   13
         Top             =   3285
         Width           =   5040
      End
      Begin VB.TextBox TxtCheImport 
         Height          =   315
         Left            =   1470
         TabIndex        =   12
         Top             =   2910
         Width           =   1400
      End
      Begin MSComCtl2.DTPicker TxtCheFecEnt 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60555265
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker TxtCheFecEmi 
         Height          =   315
         Left            =   1470
         TabIndex        =   10
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60555265
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker TxtCheFecVto 
         Height          =   315
         Left            =   5160
         TabIndex        =   11
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60555265
         CurrentDate     =   41098
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Concepto:"
         Height          =   195
         Index           =   10
         Left            =   600
         TabIndex        =   26
         Top             =   2205
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Responsable:"
         Height          =   195
         Index           =   9
         Left            =   375
         TabIndex        =   25
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cheque:"
         Height          =   195
         Index           =   7
         Left            =   3555
         TabIndex        =   24
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   23
         Top             =   3315
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   0
         Left            =   885
         TabIndex        =   22
         Top             =   315
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   21
         Top             =   2970
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  Emisión:"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   20
         Top             =   2580
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Vto:"
         Height          =   195
         Index           =   3
         Left            =   4110
         TabIndex        =   19
         Top             =   2580
         Width           =   1005
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
      Left            =   180
      TabIndex        =   33
      Top             =   4065
      Width           =   750
   End
End
Attribute VB_Name = "FrmCargaCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function Validar() As Boolean
   If Trim(TxtCheNumero.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Número de Cheque.", 16, TIT_MSGBOX
        TxtCheNumero.SetFocus
        Exit Function
        
   ElseIf Trim(TxtBanco.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Banco.", 16, TIT_MSGBOX
        TxtBanco.SetFocus
        Exit Function
        
   ElseIf Trim(txtlocalidad.Text) = "" Then
        Validar = False
        MsgBox "Ingrese la Localidad del Banco.", 16, TIT_MSGBOX
        txtlocalidad.SetFocus
        Exit Function
        
   ElseIf Trim(TxtSucursal.Text) = "" Then
        Validar = False
        MsgBox "Ingrese la Sucursal del Banco.", 16, TIT_MSGBOX
        TxtSucursal.SetFocus
        Exit Function
        
   ElseIf Trim(txtcodigo.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Código del Banco.", 16, TIT_MSGBOX
        txtcodigo.SetFocus
        Exit Function
        
   ElseIf Trim(Me.TxtCodInt.Text) = "" Then
        Validar = False
        MsgBox "Verifique el Código de Banco.", 16, TIT_MSGBOX
        txtcodigo.SetFocus
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
        
   ElseIf Trim(TxtCheFecVto.Value) = "" Then
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

Private Sub CmdBanco_Click()
    Viene_Cheque = True
    buscobanco = 1
    ABMBanco.Show vbModal
    Viene_Cheque = False
    buscobanco = 0
End Sub

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    
    If Trim(TxtCheNumero.Text) <> "" And Trim(Me.TxtCodInt.Text) <> "" Then
        resp = MsgBox("Seguro desea eliminar el Cheque Nº: " & Trim(Me.TxtCheNumero.Text) & "? ", 36, TIT_MSGBOX)
        If resp <> 6 Then Exit Sub
        
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Borrando..."
        
        sql = "SELECT BOL_NUMERO "
        sql = sql & " FROM ChequeEstadoVigente "
        sql = sql & " WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text)
        sql = sql & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If Not IsNull(rec!BOL_NUMERO) Then
               MsgBox "No se puede eliminar este Cheque porque fue depositado", vbExclamation, TIT_MSGBOX
               rec.Close
               Screen.MousePointer = vbNormal
               lblEstado.Caption = ""
               Exit Sub
             End If
        End If
        rec.Close

        DBConn.BeginTrans
        DBConn.Execute "DELETE FROM CHEQUE_ESTADOS WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
                       
        DBConn.Execute "DELETE FROM CHEQUE WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
        
        Screen.MousePointer = vbNormal
        lblEstado.Caption = ""
        DBConn.CommitTrans
        CmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub cmdBuscaCheque_Click()
    Dim codint As Integer
    frmBuscar.TipoBusqueda = 6
    frmBuscar.Show vbModal
    'TxtCheNumero.Text = frmBuscar.grdBuscar.Col
    TxtCheNumero.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
    
    
    
    
    sql = "SELECT C.*, E.ECH_CODIGO, E.ECH_DESCRI,B.BAN_DESCRI "
    sql = sql & " FROM CHEQUE C, ESTADO_CHEQUE E, CHEQUE_ESTADOS CE, BANCO B "
    sql = sql & "WHERE "
    sql = sql & " C.CHE_NUMERO = CE.CHE_NUMERO AND"
    sql = sql & " C.BAN_CODINT = CE.BAN_CODINT AND"
    sql = sql & " CE.ECH_CODIGO = E.ECH_CODIGO AND"
    sql = sql & " B.BAN_CODINT = C.BAN_CODINT AND"
    sql = sql & " C.CHE_NUMERO LIKE '" & TxtCheNumero & "'"
    sql = sql & " AND C.BAN_CODINT = " & XN(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4))
    sql = sql & " AND CE.CES_FECHA = (SELECT MAX(CES_FECHA) FROM CHEQUE_ESTADOS"
    sql = sql & " WHERE BAN_CODINT = " & XN(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)) & " "
    sql = sql & " AND CHE_NUMERO LIKE '" & TxtCheNumero & "'" & ")"
    
    'sql = sql & " AND E.ECH_CODIGO <> 1 "
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        
        If rec!ECH_CODIGO = 1 Then 'EN CARTERA
            TxtBanDescri.Text = rec!BAN_DESCRI
            TxtBanco.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 5)
            txtcodigo.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 8)
            txtlocalidad.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 6)
            TxtSucursal.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 7)
            txtcodigo.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 8)
            TxtCheImport.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
            TxtCheFecVto.Value = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2)
            TxtCodInt.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)
            TxtCheNombre.Text = IIf(IsNull(rec!CHE_NOMBRE), "", rec!CHE_NOMBRE)
            TxtCheMotivo.Text = IIf(IsNull(rec!CHE_MOTIVO), "", rec!CHE_MOTIVO)
            TxtCheFecEmi.Value = IIf(IsNull(rec!CHE_FECEMI), "", rec!CHE_FECEMI)
            TxtCheObserv.Text = IIf(IsNull(rec!CHE_OBSERV), "", rec!CHE_OBSERV)
            
        Else
            MsgBox "El Cheque " & TxtCheNumero.Text & " - " & Trim(rec!BAN_DESCRI) & " no puede MODIFICARSE porque está  en estado " & rec!ECH_DESCRI & " ", vbInformation, TIT_MSGBOX
            CmdNuevo_Click
        End If
    End If
    rec.Close
    
    
End Sub

Private Sub cmdGrabar_Click()
    
  If Validar = True Then
  
    On Error GoTo CLAVOSE
    
    DBConn.BeginTrans
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    Me.Refresh
    
    sql = "SELECT * FROM CHEQUE WHERE CHE_NUMERO LIKE '" & TxtCheNumero.Text & "' "
    sql = sql & " AND BAN_CODINT = " & XN(TxtCodInt.Text) & ""
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount = 0 Then
         sql = "INSERT INTO CHEQUE(CHE_NUMERO,BAN_CODINT,CHE_NOMBRE,CHE_IMPORT,CHE_FECEMI,"
         sql = sql & "CHE_FECVTO,CHE_FECENT,CHE_MOTIVO,CHE_OBSERV)"
         sql = sql & " VALUES (" & XS(Me.TxtCheNumero.Text) & ","
         sql = sql & XN(Me.TxtCodInt.Text) & "," & XS(Me.TxtCheNombre.Text) & ","
         sql = sql & XN(Me.TxtCheImport.Text) & "," & XDQ(Me.TxtCheFecEmi.Value) & ","
         sql = sql & XDQ(Me.TxtCheFecVto.Value) & "," & XDQ(Me.TxtCheFecEnt.Value) & ","
         sql = sql & XS(Me.TxtCheMotivo.Text) & "," & XS(Me.TxtCheObserv.Text) & " )"
         DBConn.Execute sql
         
         'Insert en la Tabla de Estados de Cheques
            sql = "INSERT INTO CHEQUE_ESTADOS (CHE_NUMERO,BAN_CODINT,ECH_CODIGO,CES_FECHA,CES_DESCRI)"
            sql = sql & " VALUES ("
            sql = sql & XS(Me.TxtCheNumero.Text) & ","
            sql = sql & XN(Me.TxtCodInt.Text) & "," & XN(1) & ","
            sql = sql & XDQ(Date) & ",'CHEQUE EN CARTERA')"
            DBConn.Execute sql
    Else
         sql = "UPDATE CHEQUE SET CHE_NOMBRE = " & XS(Me.TxtCheNombre.Text)
         sql = sql & ",CHE_IMPORT = " & XN(Me.TxtCheImport.Text)
         sql = sql & ",CHE_FECEMI =" & XDQ(Me.TxtCheFecEmi.Value)
         sql = sql & ",CHE_FECVTO =" & XDQ(Me.TxtCheFecVto.Value)
         sql = sql & ",CHE_FECENT = " & XDQ(Me.TxtCheFecEnt.Value)
         sql = sql & ",CHE_MOTIVO = " & XS(Me.TxtCheMotivo.Text)
         sql = sql & ",CHE_OBSERV = " & XS(Me.TxtCheObserv.Text)
         sql = sql & " WHERE CHE_NUMERO LIKE '" & Me.TxtCheNumero.Text & "'"
         sql = sql & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
         DBConn.Execute sql
    End If
    rec.Close
     
    
    
    '************* PREGUNTAR POR SI DESEA IMPRIMIR ***************
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.CommitTrans
    
    If frmReciboCliente.Visible = True Then
        frmReciboCliente.TxtCheNumero.Text = TxtCheNumero.Text
        Me.Visible = False
    End If
    CmdNuevo_Click
 End If
 Exit Sub
      
CLAVOSE:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub CmdNuevo_Click()
    Me.TxtCheFecEnt.Value = Null
    Me.TxtCheNumero.Enabled = True
    Me.TxtBanco.Enabled = True
    Me.txtlocalidad.Enabled = True
    Me.TxtSucursal.Enabled = True
    Me.txtcodigo.Enabled = True
    Me.TxtCheNombre.Enabled = True
   ' MtrObjetos = Array(TxtCheNumero, TxtBANCO, TxtLOCALIDAD, TxtSUCURSAL, TxtCODIGO, TxtCheNombre)
   ' Call CambiarColor(MtrObjetos, 6, &H80000005, "E")
    Me.TxtCheNumero.Text = ""
    Me.TxtBanco.Text = ""
    Me.txtlocalidad.Text = ""
    Me.TxtSucursal.Text = ""
    TxtBanDescri.Text = ""
    Me.txtcodigo.Text = ""
    Me.TxtCodInt.Text = ""
    Me.TxtCheNombre.Text = ""
    Me.TxtCheMotivo.Text = ""
    Me.TxtCheFecEmi.Value = Null
    Me.TxtCheFecVto.Value = Null
    Me.TxtCheImport.Text = ""
    Me.TxtCheObserv.Text = ""
    'Me.TxtCheFecEnt.SetFocus
    'TxtCheNombre.ForeColor = &H80000005
    lblEstado.Caption = ""
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set FrmCargaCheques = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    
    TxtCheFecEnt.Value = Date
    lblEstado.Caption = ""
End Sub

Private Sub TxtBANCO_LostFocus()
    If Len(TxtBanco.Text) < 3 Then TxtBanco.Text = CompletarConCeros(TxtBanco.Text, 3)
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
    With TxtCheImport
    .SelStart = 0
    .SelLength = Len(.Text)
    End With
End Sub

Private Sub TxtCheImport_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(TxtCheImport.Text, KeyAscii)
End Sub

Private Sub TxtCheImport_LostFocus()
   If Trim(TxtCheImport.Text) <> "" Then TxtCheImport.Text = Valido_Importe(TxtCheImport)
    
End Sub

Private Sub TxtCheNombre_LostFocus()
   If Me.TxtCheNombre.Text <> "" Then
      Me.TxtCheMotivo.SetFocus
   Else
      MsgBox "Debe ingresar la Persona responsable.!", 16, TIT_MSGBOX
   End If
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
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCheNombre_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCheNumero_LostFocus()
   If Len(TxtCheNumero.Text) < 8 Then TxtCheNumero.Text = CompletarConCeros(TxtCheNumero.Text, 8)
      If Len(TxtCheNumero.Text) <> 8 Then
        MsgBox "El Numero de cheque debe tener 8 digitos", vbExclamation, TIT_MSGBOX
        TxtCheNumero.SetFocus
      End If
End Sub

Private Sub TxtCheObserv_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtBanco_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    Dim MtrObjetos As Variant
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
        
    'ChequeRegistrado = False
    
    If Len(txtcodigo.Text) < 6 Then txtcodigo.Text = CompletarConCeros(txtcodigo.Text, 6)
     
    If Trim(Me.TxtCheNumero.Text) <> "" And _
       Trim(Me.TxtBanco.Text) <> "" And _
       Trim(Me.txtlocalidad.Text) <> "" And _
       Trim(Me.TxtSucursal.Text) <> "" And _
       Trim(Me.txtcodigo.Text) <> "" Then
       
       'BUSCO EL CODIGO INTERNO
       sql = "SELECT BAN_CODINT, BAN_DESCRI FROM BANCO WHERE BAN_BANCO = " & _
       XS(TxtBanco.Text) & " AND BAN_LOCALIDAD = " & _
       XS(Me.txtlocalidad.Text) & " AND BAN_SUCURSAL = " & _
       XS(Me.TxtSucursal.Text) & " AND BAN_CODIGO = " & XS(txtcodigo.Text)
       rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
       If rec.RecordCount > 0 Then 'EXITE
          TxtCodInt.Text = rec!BAN_CODINT
          TxtBanDescri.Text = rec!BAN_DESCRI
          rec.Close
       Else
          If Me.ActiveControl.Name <> "CmdSalir" And Me.ActiveControl.Name <> "CmdNuevo" Then
            MsgBox "Banco NO Registrado.", 16, TIT_MSGBOX
            Me.CmdBanco.SetFocus
          End If
          rec.Close
          Exit Sub
       End If
       
       'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE " & _
              " WHERE CHE_NUMERO = " & XS(TxtCheNumero.Text) & _
                " AND BAN_CODINT = " & XN(TxtCodInt.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then 'EXITE
            Me.TxtCheFecEnt.Value = rec!CHE_FECENT
            Me.TxtCheNumero.Text = Trim(rec!CHE_NUMERO)
            
            'BUSCO LOS ATRIBUTOS DE BANCO
            sql = "SELECT BAN_BANCO,BAN_LOCALIDAD,BAN_SUCURSAL,BAN_CODIGO FROM BANCO " & _
                   "WHERE BAN_CODINT = " & XN(Me.TxtCodInt.Text)
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.RecordCount > 0 Then 'EXITE
                Me.TxtBanco.Text = Rec1!BAN_BANCO
                Me.txtlocalidad.Text = Rec1!BAN_LOCALIDAD
                Me.TxtSucursal.Text = Rec1!BAN_SUCURSAL
                Me.txtcodigo.Text = Rec1!BAN_CODIGO
            End If
            Rec1.Close
            Me.TxtCheNombre.Text = ChkNull(rec!CHE_NOMBRE)
            Me.TxtCheMotivo.Text = rec!CHE_MOTIVO
            Me.TxtCheFecEmi.Value = rec!CHE_FECEMI
            Me.TxtCheFecVto.Value = rec!CHE_FECVTO
            Me.TxtCheImport.Text = Valido_Importe(rec!che_import)
            Me.TxtCheObserv.Text = ChkNull(rec!CHE_OBSERV)
            
            TxtCheNumero.Enabled = False
            TxtBanco.Enabled = False
            txtlocalidad.Enabled = False
            TxtSucursal.Enabled = False
            txtcodigo.Enabled = False
            
            MtrObjetos = Array(TxtCheNumero, TxtBanco, txtlocalidad, TxtSucursal, txtcodigo)
            Call CambiarColor(MtrObjetos, 5, &H80000018, "D")
            
        Else
           'TxtCheNombre.ForeColor = &HC0FFFF
           rec.Close
           Exit Sub
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub Txtlocalidad_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtLOCALIDAD_LostFocus()
    If Len(txtlocalidad.Text) < 3 Then txtlocalidad.Text = CompletarConCeros(txtlocalidad.Text, 3)
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtSucursal_LostFocus()
    If Len(TxtSucursal.Text) < 3 Then TxtSucursal.Text = CompletarConCeros(TxtSucursal.Text, 3)
End Sub
