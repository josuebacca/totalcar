VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ABMCuentasBancarias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ABM  de Cuentas Bancarias"
   ClientHeight    =   4815
   ClientLeft      =   1920
   ClientTop       =   1770
   ClientWidth     =   6540
   Icon            =   "ABMCuentasBancarias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "ABMCuentasBancarias.frx":0442
      Height          =   720
      Left            =   3855
      Picture         =   "ABMCuentasBancarias.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4065
      Width           =   855
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMCuentasBancarias.frx":0A56
      Height          =   720
      Left            =   2985
      Picture         =   "ABMCuentasBancarias.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4065
      Width           =   855
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      DisabledPicture =   "ABMCuentasBancarias.frx":106A
      Height          =   720
      Left            =   4725
      Picture         =   "ABMCuentasBancarias.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4065
      Width           =   855
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMCuentasBancarias.frx":167E
      Height          =   720
      Left            =   5595
      Picture         =   "ABMCuentasBancarias.frx":1988
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4065
      Width           =   855
   End
   Begin TabDlg.SSTab TabTB 
      Height          =   3975
      Left            =   60
      TabIndex        =   15
      Top             =   45
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   7011
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
      TabPicture(0)   =   "ABMCuentasBancarias.frx":1C92
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMCuentasBancarias.frx":1CAE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74775
         TabIndex        =   17
         Top             =   375
         Width           =   5910
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   330
            Left            =   5295
            MaskColor       =   &H000000FF&
            Picture         =   "ABMCuentasBancarias.frx":1CCA
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Buscar"
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   435
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   300
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   12
            Top             =   225
            Width           =   4080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. de Cta.:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   18
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Datos de la  Cuenta "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3240
         Left            =   300
         TabIndex        =   16
         Top             =   540
         Width           =   5775
         Begin VB.ComboBox CboTcuCodigo 
            Height          =   315
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1365
            Width           =   3930
         End
         Begin VB.TextBox TxtDescri 
            Height          =   315
            Left            =   1395
            MaxLength       =   20
            TabIndex        =   7
            Top             =   2760
            Width           =   3915
         End
         Begin VB.TextBox TxtSaldoAct 
            Height          =   315
            Left            =   4065
            MaxLength       =   30
            TabIndex        =   6
            Top             =   2325
            Width           =   1245
         End
         Begin VB.TextBox TxtSaldoIni 
            Height          =   315
            Left            =   4065
            MaxLength       =   30
            TabIndex        =   4
            Top             =   1875
            Width           =   1245
         End
         Begin VB.ComboBox CboBancos 
            Height          =   315
            Left            =   1395
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   540
            Width           =   3930
         End
         Begin VB.TextBox TxtCuenta 
            Height          =   315
            Left            =   1395
            MaxLength       =   8
            TabIndex        =   1
            Top             =   945
            Width           =   1515
         End
         Begin MSComCtl2.DTPicker fechaApertura 
            Height          =   315
            Left            =   1395
            TabIndex        =   3
            Top             =   1875
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61145089
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker fechaCierre 
            Height          =   315
            Left            =   1395
            TabIndex        =   5
            Top             =   2325
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61145089
            CurrentDate     =   41098
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cuenta:"
            Height          =   195
            Index           =   2
            Left            =   435
            TabIndex        =   27
            Top             =   975
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Index           =   4
            Left            =   780
            TabIndex        =   26
            Top             =   540
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Apertura:"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   25
            Top             =   1935
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Cierre:"
            Height          =   195
            Index           =   5
            Left            =   345
            TabIndex        =   24
            Top             =   2385
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   8
            Left            =   405
            TabIndex        =   23
            Top             =   2760
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cuenta:"
            Height          =   195
            Index           =   9
            Left            =   150
            TabIndex        =   22
            Top             =   1395
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Actual:"
            Height          =   195
            Index           =   7
            Left            =   3060
            TabIndex        =   21
            Top             =   2385
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Inicial:"
            Height          =   195
            Index           =   6
            Left            =   3105
            TabIndex        =   20
            Top             =   1935
            Width           =   900
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   2715
         Left            =   -74805
         TabIndex        =   14
         Top             =   1155
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   4789
         _Version        =   393216
         Cols            =   9
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
      Left            =   150
      TabIndex        =   19
      Top             =   4230
      Width           =   750
   End
End
Attribute VB_Name = "ABMCuentasBancarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As ADODB.Recordset
Dim sql As String
Dim resp As Integer
Dim cuit  As String
Dim Fecha As String

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(Me.TxtCuenta.Text) <> "" Then
    
        sql = " SELECT CTA_NROCTA " & _
              " FROM CHEQUE_PROPIO" & _
              " WHERE BAN_CODINT = " & XN(CboBancos.ItemData(CboBancos.ListIndex)) & _
                " AND CTA_NROCTA = " & XS(TxtCuenta.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
            MsgBox "No se puede eliminar esta Cuenta Bancaria ya que tiene Cheques asociados !", vbExclamation, TIT_MSGBOX
            rec.Close
            Exit Sub
        End If
        rec.Close
        
        resp = MsgBox("Seguro desea eliminar la Cuenta: " & Trim(Me.TxtCuenta) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = 11
        lblEstado.Caption = "Eliminando ..."
        
        sql = "DELETE FROM CTA_BANCARIA "
        sql = sql & " WHERE BAN_CODINT = " & XN(CboBancos.ItemData(CboBancos.ListIndex))
        sql = sql & " AND CTA_NROCTA = " & XS(TxtCuenta.Text)
        DBConn.Execute sql
        
        If TxtCuenta.Enabled Then TxtCuenta.SetFocus
        lblEstado.Caption = ""
        Screen.MousePointer = 1
        CmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = 1
    Mensaje 2
End Sub

Private Sub CmdBuscAprox_Click()

    GrdModulos.Rows = 1

    Screen.MousePointer = vbHourglass
    Me.Refresh
    
    sql = "SELECT BAN_DESCRI,C.BAN_CODINT,CTA_NROCTA,CTA_FECAPE,TCU_CODIGO,CTA_SALINI,CTA_SALACT,CTA_FECCIE,CTA_DESCRI" & _
          " FROM CTA_BANCARIA C, BANCO B" & _
          " WHERE B.BAN_CODINT = C.BAN_CODINT"
    If Trim(TxtDescriB) <> "" Then sql = sql & " AND CTA_NROCTA LIKE '%" & Trim(TxtDescriB) & "%'"
    sql = sql & " ORDER BY CTA_NROCTA"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            GrdModulos.AddItem Trim(rec!BAN_DESCRI) & Chr(9) & Trim(rec!CTA_NROCTA) & Chr(9) & _
                               Trim(rec!CTA_FECAPE) & Chr(9) & Valido_Importe(Trim(rec!CTA_SALINI)) & Chr(9) & _
                               Valido_Importe(Trim(rec!CTA_SALACT)) & Chr(9) & Trim(rec!CTA_FECCIE) & Chr(9) & _
                               Trim(rec!CTA_DESCRI) & Chr(9) & rec!TCU_CODIGO & Chr(9) & rec!BAN_CODINT
            rec.MoveNext
        Loop
        If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        MsgBox "No se encontraron items con esta descripción !", vbExclamation, TIT_MSGBOX
        TxtDescriB.SelStart = 0
        TxtDescriB.SelLength = Len(TxtDescriB)
        If TxtDescriB.Enabled Then TxtDescriB.SetFocus
    End If
    rec.Close
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdGrabar_Click()
    Dim Sucursal As String
    Dim Localidad As String
    Dim Banco As String
    On Error GoTo CLAVOSE
    
    If Trim(TxtCuenta.Text) = "" Then
        MsgBox "No ha ingresado el Nro. de Cuenta !", vbExclamation, TIT_MSGBOX
        If TxtCuenta.Enabled Then TxtCuenta.SetFocus
        Exit Sub
    ElseIf Trim(Me.fechaApertura.Value) = "" Then
        MsgBox "No ha ingresado la Fecha de Apertura !", vbExclamation, TIT_MSGBOX
        If Me.fechaApertura.Enabled Then fechaApertura.SetFocus
        Exit Sub
    ElseIf Trim(Me.TxtSaldoIni.Text) = "" Then
        MsgBox "No ha ingresado el Saldo Inicial !", vbExclamation, TIT_MSGBOX
        If Me.TxtSaldoIni.Enabled Then TxtSaldoIni.SetFocus
        Exit Sub
    ElseIf Trim(Me.TxtSaldoAct.Text) = "" Then
        MsgBox "No ha ingresado el Saldo Actual !", vbExclamation, TIT_MSGBOX
        If TxtSaldoAct.Enabled Then TxtSaldoAct.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    lblEstado.Caption = "Guardando ..."
    
    'Busco los datos de la cuenta bancaria
    sql = "SELECT CTA_NROCTA FROM CTA_BANCARIA WHERE "
    sql = sql & " BAN_CODINT = " & XN(CboBancos.ItemData(CboBancos.ListIndex))
    sql = sql & " AND CTA_NROCTA = " & XS(TxtCuenta)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount = 0 Then
        sql = "INSERT INTO CTA_BANCARIA "
        sql = sql & " (BAN_CODINT,CTA_NROCTA,CTA_FECAPE,CTA_SALINI,"
        sql = sql & " CTA_SALACT,CTA_FECCIE,TCU_CODIGO,CTA_DESCRI) VALUES "
        sql = sql & "( " & XN(CboBancos.ItemData(CboBancos.ListIndex)) & ","
        sql = sql & XS(TxtCuenta.Text) & ","
        sql = sql & XDQ(fechaApertura.Value) & ","
        sql = sql & XN(CDbl(TxtSaldoIni.Text)) & ","
        sql = sql & XN(CDbl(TxtSaldoAct.Text)) & ","
        sql = sql & XDQ(fechaCierre.Value) & ","
        sql = sql & XN(CboTcuCodigo.ItemData(CboTcuCodigo.ListIndex)) & ","
        sql = sql & XS(TxtDescri.Text) & ")"
    Else
        sql = "UPDATE CTA_BANCARIA SET "
        sql = sql & " CTA_FECAPE =" & XDQ(fechaApertura.Value)
        sql = sql & ", CTA_SALINI =" & XN(CDbl(TxtSaldoIni.Text))
        sql = sql & ", CTA_SALACT =" & XN(CDbl(TxtSaldoAct.Text))
        sql = sql & ", CTA_FECCIE =" & XDQ(Me.fechaCierre.Value)
        sql = sql & ", TCU_CODIGO =" & XN(CboTcuCodigo.ItemData(CboTcuCodigo.ListIndex))
        sql = sql & ", CTA_DESCRI =" & XS(Me.TxtDescri.Text)
        sql = sql & " Where BAN_CODINT = " & XN(CboBancos.ItemData(CboBancos.ListIndex))
        sql = sql & " And CTA_NROCTA = " & XS(TxtCuenta)
    End If
    
    DBConn.Execute sql
    rec.Close
    Screen.MousePointer = 1
    CmdNuevo_Click
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    lblEstado = ""
    Screen.MousePointer = 1
    Mensaje 1
End Sub

Private Sub CmdNuevo_Click()
    TabTB.Tab = 0
    CboBancos.Enabled = True
    TxtCuenta.Enabled = True
        
    fechaApertura.Value = Null
    fechaCierre.Value = Null
    TxtSaldoAct.Text = ""
    TxtSaldoIni.Text = ""
    TxtDescri.Text = ""
    TxtCuenta.Text = ""
    CboBancos.ListIndex = 0
    CboTcuCodigo.ListIndex = 0
    CmdGrabar.Enabled = True
    CmdBorrar.Enabled = True
    lblEstado.Caption = ""
    GrdModulos.Rows = 1
    CboBancos.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set ABMCuentasBancarias = Nothing
End Sub

Private Sub Form_Activate()
    Call Centrar_pantalla(Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then TabTB.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'si presiono ESCAPE salgo del form
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
    If KeyAscii = vbKeyReturn Then   'avanza de campo
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    
    lblEstado.Caption = ""
    GrdModulos.FormatString = "Banco|Nº Cta.|Fecha de Apertura|Saldo Inicial" _
                             & "|Saldo Actual|Fecha de Cierre|Descripción|" _
                             & "tipo cuenta|CODIGO BANCO"
    GrdModulos.ColWidth(0) = 2000
    GrdModulos.ColWidth(1) = 1000
    GrdModulos.ColWidth(2) = 1500
    GrdModulos.ColWidth(3) = 1500
    GrdModulos.ColWidth(4) = 1500
    GrdModulos.ColWidth(5) = 1500
    GrdModulos.ColWidth(6) = 3000
    GrdModulos.ColWidth(7) = 0
    GrdModulos.ColWidth(8) = 0
    GrdModulos.Rows = 1
    
    TabTB.Tab = 0
    Set rec = New ADODB.Recordset
    
    CboBancos.Clear
    sql = "SELECT DISTINCT B.BAN_CODINT,B.BAN_DESCRI FROM BANCO B "
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            CboBancos.AddItem Trim(rec!BAN_DESCRI)
            CboBancos.ItemData(CboBancos.NewIndex) = rec!BAN_CODINT
            rec.MoveNext
        Loop
        CboBancos.ListIndex = 0
    End If
    rec.Close
    
    CboTcuCodigo.Clear
    sql = "SELECT TCU_CODIGO,TCU_DESCRI FROM TIPO_CUENTA "
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            CboTcuCodigo.AddItem Trim(rec!TCU_DESCRI)
            CboTcuCodigo.ItemData(CboTcuCodigo.NewIndex) = rec!TCU_CODIGO
            rec.MoveNext
        Loop
        Me.CboTcuCodigo.ListIndex = 0
    End If
    rec.Close
    Screen.MousePointer = 1
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
        'paso el item seleccionado al tab 'DATOS'
        GrdModulos.Col = 0
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 8)), CboBancos)
        Me.CboBancos.Enabled = False
        
        Me.TxtCuenta.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
        Me.TxtCuenta.Enabled = False
        
        Me.fechaApertura.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
        Me.TxtSaldoIni.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 3))
        Me.TxtSaldoAct.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 4))
        Me.fechaCierre.Value = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 5))
        Me.TxtDescri.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 6))
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 7)), CboTcuCodigo)
        If Me.TxtDescriB.Enabled Then Me.TxtDescriB.SetFocus
        TabTB.Tab = 0
    End If
End Sub

Private Sub GrdModulos_GotFocus()
    GrdModulos.Col = 0
    GrdModulos.ColSel = 1
    GrdModulos.HighLight = flexHighlightAlways
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then CmdBorrar_Click
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub GrdModulos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub GrdModulos_LostFocus()
    GrdModulos.HighLight = flexHighlightNever
End Sub

Private Sub tabTB_Click(PreviousTab As Integer)
    'Si cambio de 'Pestaña' en el tab
    'pongo el foco en el primer campo de la misma
    If TabTB.Tab = 0 And Me.Visible Then
        If CboBancos.Enabled = True Then
            Me.CboBancos.SetFocus
        Else
            Me.CboTcuCodigo.SetFocus
        End If
        CmdGrabar.Enabled = True
        CmdBorrar.Enabled = True
    End If
    If TabTB.Tab = 1 Then
        TxtDescriB.Text = ""
        If TxtDescriB.Enabled Then TxtDescriB.SetFocus
        CmdGrabar.Enabled = False
        CmdBorrar.Enabled = False
    End If
End Sub

Private Sub TxtCuenta_LostFocus()
    If TxtCuenta.Text <> "" Then
        sql = "SELECT * FROM CTA_BANCARIA WHERE " & _
        "BAN_CODINT = " & XN(Me.CboBancos.ItemData(CboBancos.ListIndex)) & " AND CTA_NROCTA = " & XS(TxtCuenta)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
            TxtSaldoIni = Valido_Importe(Trim(rec!CTA_SALINI))
            TxtSaldoAct = Valido_Importe(Trim(rec!CTA_SALACT))
            Call BuscaCodigoProxItemData(CInt(rec!TCU_CODIGO), CboTcuCodigo)
            fechaApertura.Value = rec!CTA_FECAPE
            fechaCierre.Value = ChkNull(rec!CTA_FECCIE)
            Me.TxtDescri.Text = Trim(ChkNull(rec!CTA_DESCRI))
            TxtCuenta.Enabled = False
            CboBancos.Enabled = False
            CboTcuCodigo.SetFocus
        Else
            Me.CboTcuCodigo.ListIndex = 0
            fechaApertura.Value = Null
            fechaCierre.Value = Null
            TxtSaldoIni.Text = ""
            TxtSaldoAct.Text = ""
            TxtDescri.Text = ""
            CboTcuCodigo.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then CmdBuscAprox_Click
End Sub

Private Sub TxtSaldoAct_GotFocus()
    SelecTexto TxtSaldoAct
End Sub

Private Sub TxtSaldoAct_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(TxtSaldoAct.Text, KeyAscii)
End Sub

Private Sub TxtSaldoAct_LostFocus()
    TxtSaldoAct.Text = Valido_Importe(TxtSaldoAct.Text)
End Sub

Private Sub TxtSaldoIni_GotFocus()
    SelecTexto TxtSaldoIni
End Sub

Private Sub TxtSaldoIni_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(TxtSaldoIni.Text, KeyAscii)
End Sub

Private Sub TxtSaldoIni_LostFocus()
  TxtSaldoIni.Text = Valido_Importe(TxtSaldoIni.Text)
End Sub
