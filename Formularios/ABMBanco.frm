VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ABMBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ABM  de Banco"
   ClientHeight    =   5355
   ClientLeft      =   1500
   ClientTop       =   1815
   ClientWidth     =   7155
   Icon            =   "ABMBanco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMBanco.frx":0442
      Height          =   720
      Left            =   3390
      Picture         =   "ABMBanco.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4575
      Width           =   900
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMBanco.frx":0A56
      Height          =   720
      Left            =   4305
      Picture         =   "ABMBanco.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4575
      Width           =   900
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMBanco.frx":106A
      Height          =   720
      Left            =   5220
      Picture         =   "ABMBanco.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4575
      Width           =   900
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMBanco.frx":167E
      Height          =   720
      Left            =   6135
      Picture         =   "ABMBanco.frx":1988
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4575
      Width           =   900
   End
   Begin TabDlg.SSTab TabTB 
      Height          =   4335
      Left            =   30
      TabIndex        =   10
      Top             =   180
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "ABMBanco.frx":1C92
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMBanco.frx":1CAE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdBancos"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   " Buscar por "
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
         Height          =   1185
         Left            =   -74865
         TabIndex        =   18
         Top             =   405
         Width           =   6750
         Begin VB.TextBox TxtDescri 
            Height          =   285
            Left            =   1020
            TabIndex        =   23
            Top             =   795
            Width           =   4815
         End
         Begin VB.TextBox TxtCODIGO 
            Height          =   285
            Left            =   5025
            MaxLength       =   6
            TabIndex        =   22
            Top             =   330
            Width           =   795
         End
         Begin VB.TextBox TxtLOCALIDAD 
            Height          =   285
            Left            =   2175
            MaxLength       =   3
            TabIndex        =   20
            Top             =   330
            Width           =   540
         End
         Begin VB.TextBox TxtBANCO 
            Height          =   285
            Left            =   660
            MaxLength       =   3
            TabIndex        =   19
            Top             =   330
            Width           =   540
         End
         Begin VB.TextBox TxtSUCURSAL 
            Height          =   285
            Left            =   3615
            MaxLength       =   3
            TabIndex        =   21
            Top             =   330
            Width           =   540
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   330
            Left            =   5985
            MaskColor       =   &H000000FF&
            Picture         =   "ABMBanco.frx":1CCA
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Buscar"
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   435
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   32
            Top             =   825
            Width           =   870
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
            Left            =   4425
            TabIndex        =   29
            Top             =   360
            Width           =   540
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
            Left            =   2910
            TabIndex        =   28
            Top             =   360
            Width           =   645
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
            Left            =   105
            TabIndex        =   27
            Top             =   360
            Width           =   510
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
            Left            =   1380
            TabIndex        =   26
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Datos del Banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   315
         TabIndex        =   11
         Top             =   525
         Width           =   6435
         Begin VB.TextBox TxtBanNomCor 
            Height          =   300
            Left            =   1380
            MaxLength       =   25
            TabIndex        =   5
            Top             =   2745
            Width           =   1695
         End
         Begin VB.TextBox TxtCodInt 
            Height          =   300
            Left            =   2220
            TabIndex        =   30
            Top             =   420
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox TxtBanCodigo 
            Height          =   300
            Left            =   1365
            MaxLength       =   6
            TabIndex        =   3
            Top             =   1785
            Width           =   720
         End
         Begin VB.TextBox TxtBanSucursal 
            Height          =   300
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   2
            Top             =   1335
            Width           =   720
         End
         Begin VB.TextBox TxtBanLocalidad 
            Height          =   300
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   1
            Top             =   885
            Width           =   720
         End
         Begin VB.TextBox TxtBanBanco 
            Height          =   300
            Left            =   1395
            MaxLength       =   3
            TabIndex        =   0
            Top             =   420
            Width           =   720
         End
         Begin VB.TextBox TxtBanDescri 
            Height          =   315
            Left            =   1380
            MaxLength       =   40
            TabIndex        =   4
            Top             =   2280
            Width           =   4875
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Corto:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   31
            Top             =   2790
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal:"
            Height          =   195
            Index           =   6
            Left            =   495
            TabIndex        =   17
            Top             =   1365
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Index           =   5
            Left            =   420
            TabIndex        =   16
            Top             =   900
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Index           =   4
            Left            =   645
            TabIndex        =   15
            Top             =   450
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Index           =   3
            Left            =   615
            TabIndex        =   13
            Top             =   1830
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   12
            Top             =   2340
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdBancos 
         Height          =   2610
         Left            =   -74895
         TabIndex        =   25
         Top             =   1635
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   4604
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorSel    =   8388736
         FocusRect       =   0
         SelectionMode   =   1
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
      Left            =   135
      TabIndex        =   14
      Top             =   4680
      Width           =   750
   End
End
Attribute VB_Name = "ABMBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As ADODB.Recordset
Dim sql As String
Dim resp As Integer

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    
    If Trim(TxtBanCodigo.Text) <> "" Then
    
        sql = " SELECT CH.CHE_NUMERO " & _
              " FROM BANCO B, CHEQUE CH " & _
              " WHERE B.BAN_CODINT = CH.BAN_CODINT " & _
                 " AND B.BAN_BANCO = " & XS(TxtBanBanco.Text) & _
             " AND B.BAN_LOCALIDAD = " & XS(Me.TxtBanLocalidad.Text) & _
              " AND B.BAN_SUCURSAL = " & XS(Me.TxtBanSucursal.Text) & _
                " AND B.BAN_CODIGO = " & XS(Me.TxtBanCodigo.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
            MsgBox "No se puede eliminar este Banco ya que tiene Cheques asociados !", vbExclamation, TIT_MSGBOX
            rec.Close
            Exit Sub
        End If
        rec.Close
        
        sql = " SELECT C.BAN_CODINT " & _
              " FROM BANCO B, CTA_BANCARIA C " & _
              " WHERE B.BAN_CODINT = C.BAN_CODINT " & _
                 " AND B.BAN_BANCO = " & XS(TxtBanBanco.Text) & _
             " AND B.BAN_LOCALIDAD = " & XS(Me.TxtBanLocalidad.Text) & _
              " AND B.BAN_SUCURSAL = " & XS(Me.TxtBanSucursal.Text) & _
                " AND B.BAN_CODIGO = " & XS(Me.TxtBanCodigo.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
            MsgBox "No se puede eliminar este Banco ya que tiene Cuentas Bancarias asociados !", vbExclamation, TIT_MSGBOX
            rec.Close
            Exit Sub
        End If
        rec.Close
        
        resp = MsgBox("Seguro desea eliminar el Banco: " & Trim(Me.TxtBanDescri.Text) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = 11
        lblEstado.Caption = "Eliminando ..."
        
        DBConn.Execute "DELETE FROM BANCO " & _
                       " WHERE BAN_BANCO = " & XS(TxtBanBanco.Text) & _
                     " AND BAN_LOCALIDAD = " & XS(TxtBanLocalidad.Text) & _
                      " AND BAN_SUCURSAL = " & XS(TxtBanSucursal.Text) & _
                        " AND BAN_CODIGO = " & XS(TxtBanCodigo.Text)
        If TxtBanDescri.Enabled Then TxtBanDescri.SetFocus
        lblEstado.Caption = ""
        Screen.MousePointer = 1
        CmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = 1
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    Set rec = New ADODB.Recordset
    GrdBancos.Rows = 1

    Screen.MousePointer = vbHourglass
    Me.Refresh
    sql = "SELECT BAN_BANCO,BAN_LOCALIDAD,BAN_SUCURSAL,BAN_CODIGO,BAN_DESCRI,BAN_NOMCOR,BAN_CODINT " & _
          "FROM BANCO WHERE BAN_DESCRI <> '' "
          
    If Trim(Me.TxtBanco.Text) <> "" Then sql = sql & " AND BAN_BANCO = " & XS(Me.TxtBanco.Text) & ""
    
    If Trim(Me.txtlocalidad.Text) <> "" Then sql = sql & " AND BAN_LOCALIDAD = " & XS(Me.txtlocalidad.Text) & ""
    
    If Trim(Me.TxtSucursal.Text) <> "" Then sql = sql & " AND BAN_SUCURSAL  = " & XS(Me.TxtSucursal.Text) & ""
    
    If Trim(Me.txtcodigo.Text) <> "" Then sql = sql & " AND BAN_CODIGO LIKE '" & Trim(Me.txtcodigo.Text) & "%'"
    
    If Trim(Me.txtdescri.Text) <> "" Then sql = sql & " AND BAN_DESCRI LIKE '" & Trim(Me.txtdescri.Text) & "%'"
    
    sql = sql & " ORDER BY BAN_DESCRI"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        lblEstado.Caption = "Buscando..."
        rec.MoveFirst
        Do While Not rec.EOF
            GrdBancos.AddItem Trim(rec.Fields(0)) & Chr(9) & _
                              Trim(rec.Fields(1)) & Chr(9) & _
                              Trim(rec.Fields(2)) & Chr(9) & _
                              Trim(rec.Fields(3)) & Chr(9) & _
                              Trim(rec.Fields(4)) & Chr(9) & _
                              Trim(rec.Fields(5)) & Chr(9) & _
                              Trim(rec!BAN_CODINT)
            rec.MoveNext
        Loop
        If GrdBancos.Enabled Then GrdBancos.SetFocus
        lblEstado.Caption = ""
    Else
        MsgBox "No se encontraron items con esta descripción !", vbExclamation, TIT_MSGBOX
        TxtBanco.SelStart = 0
        TxtBanco.SelLength = Len(TxtBanco)
        If TxtBanco.Enabled Then TxtBanco.SetFocus
    End If
    rec.Close
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdGrabar_Click()

     On Error GoTo CLAVOSE
     
    If Len(TxtBanCodigo.Text) < 6 Then TxtBanCodigo.Text = CompletarConCeros(TxtBanCodigo.Text, 6)
    
    If Trim(TxtBanBanco.Text) = "" Or _
       Trim(TxtBanLocalidad.Text) = "" Or _
       Trim(TxtBanSucursal.Text) = "" Or _
       Trim(TxtBanCodigo.Text) = "" Or _
       Trim(TxtBanDescri.Text) = "" Or _
       Trim(TxtBanNomCor.Text) = "" Then
    
       If Trim(TxtBanBanco.Text) = "" Then
            MsgBox "No ha ingresado el Banco !", vbExclamation, TIT_MSGBOX
            TxtBanBanco.SetFocus
            Exit Sub
       ElseIf Trim(TxtBanLocalidad.Text) = "" Then
            MsgBox "No ha ingresado la Localidad !", vbExclamation, TIT_MSGBOX
            TxtBanLocalidad.SetFocus
            Exit Sub
       ElseIf Trim(TxtBanSucursal.Text) = "" Then
            MsgBox "No ha ingresado la Sucursal !", vbExclamation, TIT_MSGBOX
            TxtBanSucursal.SetFocus
            Exit Sub
       ElseIf Trim(TxtBanCodigo.Text) = "" Then
            MsgBox "No ha ingresado el Código !", vbExclamation, TIT_MSGBOX
            TxtBanCodigo.SetFocus
            Exit Sub
       ElseIf Trim(TxtBanDescri.Text) = "" Then
            MsgBox "No ha ingresado la Descripción !", vbExclamation, TIT_MSGBOX
            TxtBanDescri.SetFocus
            Exit Sub
       ElseIf Trim(TxtBanNomCor.Text) = "" Then
            MsgBox "No ha ingresado el Nombre Corto del Banco!", vbExclamation, TIT_MSGBOX
            TxtBanNomCor.SetFocus
            Exit Sub
       End If
    End If
    
    Screen.MousePointer = 11
    Set rec = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    lblEstado.Caption = "Guardando ..."
    
    sql = "SELECT BAN_DESCRI FROM BANCO " & _
          "WHERE BAN_BANCO = " & XS(TxtBanBanco.Text) & _
       " AND BAN_LOCALIDAD = " & XS(Me.TxtBanLocalidad.Text) & _
        " AND BAN_SUCURSAL = " & XS(Me.TxtBanSucursal.Text) & _
          " AND BAN_CODIGO = " & XS(Me.TxtBanCodigo.Text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        DBConn.Execute "UPDATE BANCO SET BAN_DESCRI = '" & Trim(Me.TxtBanDescri) & _
                                     "', BAN_NOMCOR = '" & Trim(Me.TxtBanNomCor) & _
                                "' WHERE BAN_BANCO = " & XS(TxtBanBanco.Text) & _
                                " AND BAN_LOCALIDAD = " & XS(Me.TxtBanLocalidad.Text) & _
                                 " AND BAN_SUCURSAL = " & XS(Me.TxtBanSucursal.Text) & _
                                   " AND BAN_CODIGO = " & XS(Me.TxtBanCodigo.Text)
                                   
    Else
        TxtCodInt = "1"
        sql = "SELECT MAX(BAN_CODINT) as maximo FROM BANCO"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(Rec2.Fields!Maximo) Then TxtCodInt = XN(Rec2.Fields!Maximo) + 1
        Rec2.Close
        DBConn.Execute "INSERT INTO BANCO (BAN_CODINT,BAN_BANCO,BAN_LOCALIDAD," & _
                       " BAN_SUCURSAL,BAN_CODIGO,BAN_DESCRI,BAN_NOMCOR) VALUES " & _
                       "(" & XN(TxtCodInt) & "," & XS(TxtBanBanco.Text) & _
                       "," & XS(TxtBanLocalidad.Text) & "," & XS(TxtBanSucursal.Text) & _
                       ", " & XS(Me.TxtBanCodigo.Text) & "," & XS(TxtBanDescri.Text) & _
                       "," & XS(TxtBanNomCor.Text) & ")"
    End If
    rec.Close
    Screen.MousePointer = 1
    
    If buscobanco = 1 Then
        FrmCargaCheques.TxtBanco.Text = TxtBanBanco.Text
        FrmCargaCheques.txtlocalidad.Text = TxtBanLocalidad.Text
        FrmCargaCheques.TxtSucursal.Text = TxtBanSucursal.Text
        FrmCargaCheques.txtcodigo.Text = TxtBanCodigo.Text
        FrmCargaCheques.TxtBanDescri.Text = TxtBanDescri.Text
        FrmCargaCheques.TxtCodInt.Text = TxtCodInt.Text
        Unload Me
    Exit Sub
    End If
    
    CmdNuevo_Click
    Exit Sub
    
'    If Viene_Cheque = True Then
'       FrmCargaCheques.TxtBANCO.Text = Me.TxtBanBanco.Text
'       FrmCargaCheques.TxtLOCALIDAD.Text = Me.TxtBanLocalidad.Text
'       FrmCargaCheques.TxtSUCURSAL.Text = Me.TxtBanSucursal.Text
'       FrmCargaCheques.TxtCODIGO.Text = Me.TxtBanCodigo.Text
'       Screen.MousePointer = 1
'       Unload Me
'       Exit Sub
'    Else
'       Screen.MousePointer = 1
'       cmdNuevo_Click
'       Exit Sub
'    End If
    
CLAVOSE:
    Screen.MousePointer = 1
    Mensaje 1
    
End Sub


Private Sub CmdNuevo_Click()
    TabTB.Tab = 0
    Me.TxtBanBanco.Text = ""
    Me.TxtBanLocalidad.Text = ""
    Me.TxtBanSucursal.Text = ""
    Me.TxtBanCodigo.Text = ""
    Me.TxtBanDescri.Text = ""
    Me.TxtBanNomCor.Text = ""
    lblEstado.Caption = ""
    GrdBancos.Rows = 1
    Me.TxtBanBanco.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set ABMBanco = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then TabTB.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'si presiono ESCAPE salgo del form
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
    If KeyAscii = vbKeyReturn Then 'avanza de campo
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

    Call Centrar_pantalla(Me)

    lblEstado.Caption = ""
    GrdBancos.FormatString = "Banco|Localidad|Sucursal|Código|Descripción|Nombre Corto|Codigo Interno"
    GrdBancos.ColWidth(0) = 800
    GrdBancos.ColWidth(1) = 800
    GrdBancos.ColWidth(2) = 800
    GrdBancos.ColWidth(3) = 800
    GrdBancos.ColWidth(4) = 4500
    GrdBancos.ColWidth(5) = 1500
    GrdBancos.ColWidth(6) = 0
    GrdBancos.Rows = 1
    Screen.MousePointer = 1
    If buscobanco = 1 Then
        TabTB.Tab = 1
    Else
        TabTB.Tab = 0
    End If
End Sub

Private Sub GrdBancos_dblClick()
    If GrdBancos.row > 0 Then
        'paso el item seleccionado al tab 'DATOS'
        
        GrdBancos.Col = 0
        Me.TxtBanBanco.Text = Trim(GrdBancos.Text)
        
        GrdBancos.Col = 1
        Me.TxtBanLocalidad.Text = Trim(GrdBancos.Text)
        
        GrdBancos.Col = 2
        Me.TxtBanSucursal.Text = Trim(GrdBancos.Text)
        
        GrdBancos.Col = 3
        Me.TxtBanCodigo.Text = Trim(GrdBancos.Text)
        
        GrdBancos.Col = 4
        Me.TxtBanDescri.Text = Trim(GrdBancos.Text)
        
        GrdBancos.Col = 5
        Me.TxtBanNomCor.Text = Trim(GrdBancos.Text)
        
        GrdBancos.Col = 6
        Me.TxtCodInt.Text = Trim(GrdBancos.Text)
        
        If TxtBanDescri.Enabled Then TxtBanDescri.SetFocus
        TabTB.Tab = 0
    End If
End Sub

Private Sub GrdBancos_GotFocus()
    GrdBancos.Col = 0
    GrdBancos.ColSel = 1
    GrdBancos.HighLight = flexHighlightAlways
End Sub

Private Sub GrdBancos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then CmdBorrar_Click
    If KeyCode = vbKeyReturn Then GrdBancos_dblClick
End Sub


Private Sub GrdBancos_LostFocus()
    GrdBancos.HighLight = flexHighlightNever
End Sub

Private Sub tabTB_Click(PreviousTab As Integer)
    'Si cambio de 'Pestaña' en el tab
    'pongo el foco en el primer campo de la misma
    If TabTB.Tab = 0 And Me.Visible Then
     Me.TxtBanBanco.SetFocus
     cmdGrabar.Enabled = True
     CmdBorrar.Enabled = True
    End If
    If TabTB.Tab = 1 Then
        TxtBanco.Text = ""
        txtlocalidad.Text = ""
        TxtSucursal.Text = ""
        txtcodigo.Text = ""
        'If TxtBanco.Enabled Then TxtBanco.SetFocus
        cmdGrabar.Enabled = False
        CmdBorrar.Enabled = False
    End If
End Sub

Private Sub TxtBanBanco_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtBanco_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtBanCodigo_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtBanCodigo_LostFocus()
   Set rec = New ADODB.Recordset
   If Trim(TxtBanCodigo.Text) <> "" Then
     TxtBanCodigo.Text = CompletarConCeros(TxtBanCodigo.Text, 4)
     sql = "SELECT BAN_DESCRI,BAN_NOMCOR FROM BANCO " & _
           "WHERE BAN_BANCO = " & XS(Me.TxtBanBanco.Text) & _
        " AND BAN_LOCALIDAD = " & XS(Me.TxtBanLocalidad.Text) & _
         " AND BAN_SUCURSAL = " & XS(Me.TxtBanSucursal.Text) & _
           " AND BAN_CODIGO = " & XS(Me.TxtBanCodigo.Text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Me.TxtBanDescri.Text = Trim(rec!BAN_DESCRI)
        Me.TxtBanNomCor.Text = Trim(rec!BAN_NOMCOR)
    End If
    rec.Close
    Screen.MousePointer = 1
   End If
End Sub
Private Sub TxtBanNomCor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
  If Trim(txtcodigo.Text) <> "" Then
    If Len(txtcodigo.Text) < 6 Then txtcodigo.Text = CompletarConCeros(txtcodigo.Text, 6)
  End If
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub Txtlocalidad_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub
Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtBanDescri_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtBanLocalidad_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtBanSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCodigo_Change()
    If Trim(txtcodigo) = "" And CmdBorrar.Enabled Then
        CmdBorrar.Enabled = False
    ElseIf Trim(txtcodigo) <> "" Then
        CmdBorrar.Enabled = True
    End If
End Sub
