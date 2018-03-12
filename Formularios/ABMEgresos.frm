VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ABMEgresos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualización de Egresos"
   ClientHeight    =   4515
   ClientLeft      =   1620
   ClientTop       =   1950
   ClientWidth     =   7890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Cancelar"
      DisabledPicture =   "ABMEgresos.frx":0000
      Height          =   720
      Left            =   5010
      Picture         =   "ABMEgresos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3765
      Width           =   915
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMEgresos.frx":0614
      Height          =   720
      Left            =   4080
      Picture         =   "ABMEgresos.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3765
      Width           =   915
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMEgresos.frx":0C28
      Height          =   720
      Left            =   6870
      Picture         =   "ABMEgresos.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3765
      Width           =   915
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMEgresos.frx":123C
      Height          =   720
      Left            =   5940
      Picture         =   "ABMEgresos.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3765
      Width           =   915
   End
   Begin TabDlg.SSTab TabTB 
      Height          =   3660
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   6456
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
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
      TabPicture(0)   =   "ABMEgresos.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "B&uscar"
      TabPicture(1)   =   "ABMEgresos.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   1050
         Left            =   -74820
         TabIndex        =   18
         Top             =   360
         Width           =   7380
         Begin VB.ComboBox cboBuscaTipoGasto 
            Height          =   315
            Left            =   1305
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   570
            Width           =   4425
         End
         Begin MSComCtl2.DTPicker mFechaD 
            Height          =   330
            Left            =   1305
            TabIndex        =   10
            Top             =   180
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   61276161
            CurrentDate     =   37259
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Consultar"
            Height          =   660
            Left            =   5910
            MaskColor       =   &H000000FF&
            Picture         =   "ABMEgresos.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Buscar"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1305
         End
         Begin MSComCtl2.DTPicker mFechaH 
            Height          =   330
            Left            =   3420
            TabIndex        =   11
            Top             =   180
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   61276161
            CurrentDate     =   37259
         End
         Begin VB.Label Label4 
            Caption         =   "al"
            Height          =   240
            Left            =   3015
            TabIndex        =   25
            Top             =   270
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Gasto:"
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   24
            Top             =   615
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Periodo"
            Height          =   195
            Left            =   135
            TabIndex        =   23
            Top             =   270
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Datos del Egreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2880
         Left            =   180
         TabIndex        =   16
         Top             =   510
         Width           =   7365
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1629
            Width           =   1950
         End
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1215
            TabIndex        =   4
            Top             =   2007
            Width           =   1485
         End
         Begin VB.ComboBox CboGastos 
            Height          =   315
            Left            =   1215
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1251
            Width           =   5160
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   315
            Left            =   1215
            TabIndex        =   0
            Top             =   495
            Width           =   1050
         End
         Begin VB.TextBox TxtDescrip 
            Height          =   315
            Left            =   1215
            MaxLength       =   40
            TabIndex        =   1
            Top             =   873
            Width           =   5940
         End
         Begin MSComCtl2.DTPicker txtcing_fecha 
            Height          =   315
            Left            =   1215
            TabIndex        =   5
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61276161
            CurrentDate     =   41098
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   495
            TabIndex        =   27
            Top             =   1647
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Index           =   4
            Left            =   555
            TabIndex        =   22
            Top             =   2016
            Width           =   570
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   630
            TabIndex        =   21
            Top             =   2385
            Width           =   495
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Gasto:"
            Height          =   195
            Left            =   300
            TabIndex        =   20
            Top             =   1278
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Egreso:"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   19
            Top             =   540
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   17
            Top             =   909
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   1980
         Left            =   -74835
         TabIndex        =   14
         Top             =   1455
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   3493
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
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
      Left            =   120
      TabIndex        =   26
      Top             =   3915
      Width           =   750
   End
End
Attribute VB_Name = "ABMEgresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BuscoDatos()
    Set rec = New ADODB.Recordset
    sql = "SELECT * FROM CAJA_EGRESO"
    sql = sql & " WHERE CEGR_NUMERO = " & XN(TxtCodigo.Text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then ' si existe
        Call BuscaCodigoProxItemData(CInt(rec!TGT_CODIGO), CboGastos)
        Call BuscaCodigoProxItemData(CInt(rec!MON_CODIGO), cboMoneda)
        txtcing_fecha.Value = ChkNull(rec!CEGR_FECHA)
        txtImporte.Text = Valido_Importe(ChkNull(rec!CEGR_IMPORTE))
        TxtDescrip.Text = ChkNull(rec!CEGR_DESCRI)
        TxtDescrip.SetFocus
    Else
        MsgBox "Ingreso Inexistente", vbCritical
        TxtCodigo.Text = ""
        TxtCodigo.SetFocus
        rec.Close
        Exit Sub
    End If
    rec.Close
End Sub



Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(TxtCodigo.Text) <> "" Then
        If MsgBox("Seguro desea eliminar el Egreso '" & Trim(TxtDescrip.Text) & "' ?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Eliminando ..."
            DBConn.BeginTrans
            DBConn.Execute "DELETE FROM CAJA_EGRESO WHERE CEGR_NUMERO = " & XN(TxtCodigo.Text)
            DBConn.CommitTrans
            If TxtDescrip.Enabled Then TxtDescrip.SetFocus
            
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            CmdNuevo_Click
        End If
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub CmdBuscAprox_Click()
    Set rec = New ADODB.Recordset
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    Me.Refresh
    sql = "SELECT C.*, T.TGT_DESCRI"
    sql = sql & " FROM CAJA_EGRESO C, TIPO_GASTO T"
    sql = sql & " WHERE"
    sql = sql & " C.TGT_CODIGO=T.TGT_CODIGO"
    If mFechaD.Value <> "" And mFechaH.Value <> "" Then
        sql = sql & " AND CEGR_FECHA >= " & XDQ(mFechaD.Value) & " AND CEGR_FECHA <= " & XDQ(mFechaH.Value)
    End If
    If cboBuscaTipoGasto.List(cboBuscaTipoGasto.ListIndex) <> "<Todos>" Then
        sql = sql & " AND TGT_CODIGO=" & XN(cboBuscaTipoGasto.ItemData(cboBuscaTipoGasto.ListIndex))
    End If
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        'Número|Descripción|^Fecha|>Importe|Tipo de Ingreso|CODIGO Tipo de Ingreso
        Do While Not rec.EOF
            GrdModulos.AddItem rec!CEGR_NUMERO & Chr(9) & Trim(rec!CEGR_DESCRI) & Chr(9) & _
                        rec!CEGR_FECHA & Chr(9) & Valido_Importe(rec!CEGR_IMPORTE) & Chr(9) & _
                        rec!TGT_DESCRI & Chr(9) & rec!TGT_CODIGO
            rec.MoveNext
        Loop
        If GrdModulos.Enabled Then GrdModulos.SetFocus
        lblEstado.Caption = ""
    Else
        lblEstado.Caption = ""
        MsgBox "No se encontraron items con este Criterio", vbExclamation, TIT_MSGBOX
        If mFechaD.Enabled Then mFechaD.SetFocus
    End If
    lblEstado.Caption = ""
    rec.Close
    Screen.MousePointer = vbNormal
End Sub
Private Sub CmdGrabar_Click()
    On Error GoTo CLAVOSE
    If Trim(TxtDescrip.Text) = "" Then
        MsgBox "No ha ingresado la descripción", vbExclamation, TIT_MSGBOX
        If TxtDescrip.Enabled Then TxtDescrip.SetFocus
        Exit Sub
    End If
    If IsNull(txtcing_fecha.Value) Then
        MsgBox "No ha ingresado la Fecha del Gasto", vbExclamation, TIT_MSGBOX
        If txtcing_fecha.Enabled Then txtcing_fecha.SetFocus
        Exit Sub
    End If
    If txtImporte.Text = "" Then
        MsgBox "No ha ingresado el Importe del Gasto", vbExclamation, TIT_MSGBOX
        If txtImporte.Enabled Then txtImporte.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    
    Set rec = New ADODB.Recordset
    If TxtCodigo.Text = "" Then
        TxtCodigo.Text = "1"
        sql = "SELECT MAX(CEGR_NUMERO) as MAXIMO FROM CAJA_EGRESO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(rec.Fields!Maximo) Then TxtCodigo.Text = Val(Trim(rec.Fields!Maximo)) + 1
        rec.Close
    End If
    DBConn.BeginTrans
    
    sql = "SELECT * FROM CAJA_EGRESO WHERE CEGR_NUMERO = " & XN(TxtCodigo.Text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        sql = "UPDATE CAJA_EGRESO SET CEGR_DESCRI = " & XS(TxtDescrip.Text)
        sql = sql & " ,TGT_CODIGO = " & XN(CboGastos.ItemData(CboGastos.ListIndex))
        sql = sql & " ,CEGR_FECHA = " & XDQ(txtcing_fecha.Value)
        sql = sql & " ,CEGR_IMPORTE = " & XN(txtImporte.Text)
        sql = sql & " ,MON_CODIGO = " & XN(cboMoneda.ItemData(cboMoneda.ListIndex))
        sql = sql & " WHERE CEGR_NUMERO = " & XN(TxtCodigo.Text)
        
        DBConn.Execute sql
    Else
        
        sql = "INSERT INTO CAJA_EGRESO"
        sql = sql & " (CEGR_NUMERO, CEGR_DESCRI, TGT_CODIGO, CEGR_FECHA, CEGR_IMPORTE,MON_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(TxtCodigo.Text) & ","
        sql = sql & XS(TxtDescrip.Text) & ","
        sql = sql & XN(CboGastos.ItemData(CboGastos.ListIndex)) & ","
        sql = sql & XDQ(txtcing_fecha.Value) & ","
        sql = sql & XN(txtImporte.Text) & ","
        sql = sql & XN(cboMoneda.ItemData(cboMoneda.ListIndex)) & ")"
        DBConn.Execute sql
    End If
    rec.Close
    DBConn.CommitTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    CmdNuevo_Click
    Exit Sub
    
CLAVOSE:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub CmdNuevo_Click()
    TxtCodigo.Text = ""
    TxtDescrip.Text = ""
    txtImporte.Text = ""
    lblEstado.Caption = ""
    txtcing_fecha.Value = Null
    GrdModulos.Rows = 1
    cboMoneda.ListIndex = 0
    CboGastos.ListIndex = 0
    If TxtDescrip.Enabled And Me.Visible Then TxtDescrip.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set ABMEgresos = Nothing
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'si preciona f1 voy a la busqueda
    If KeyCode = vbKeyF1 Then TabTB.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'si presiono ESCAPE salgo del form
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
    'If KeyAscii = vbKeyReturn And _
        Me.ActiveControl.Name <> "TxtDescriB" And _
        Me.ActiveControl.Name <> "GrdContactos" Then  'avanza de campo
    If KeyAscii = vbKeyReturn Then
            SendKeys "{TAB}"
            KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    'CARGO COMBO GASTOS
    LlenarComboGastos
    'CARGO COMBO MONEDA
    LLenarComboMoneda
    
    lblEstado.Caption = ""
    cmdGrabar.Enabled = True
    cmdNuevo.Enabled = True
    cmdBorrar.Enabled = False
    
    GrdModulos.FormatString = "Número|Descripción|^Fecha|>Importe|Tipo de Ingreso|CODIGO Tipo de Ingreso"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 3600
    GrdModulos.ColWidth(2) = 1000
    GrdModulos.ColWidth(3) = 1000
    GrdModulos.ColWidth(4) = 2500
    GrdModulos.ColWidth(5) = 0
    GrdModulos.Rows = 1
    
    TabTB.Tab = 0
    
    Screen.MousePointer = vbNormal
    Call Centrar_pantalla(Me)
End Sub

Private Sub LLenarComboMoneda()
    sql = "SELECT * FROM MONEDA ORDER BY MON_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboMoneda.AddItem rec!MON_DESCRI
            cboMoneda.ItemData(cboMoneda.NewIndex) = rec!MON_CODIGO
            rec.MoveNext
        Loop
        cboMoneda.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboGastos()
    sql = "SELECT * FROM TIPO_GASTO ORDER BY TGT_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboBuscaTipoGasto.AddItem "<Todos>"
        Do While rec.EOF = False
            CboGastos.AddItem rec!TGT_DESCRI
            CboGastos.ItemData(CboGastos.NewIndex) = rec!TGT_CODIGO
            cboBuscaTipoGasto.AddItem rec!TGT_DESCRI
            cboBuscaTipoGasto.ItemData(cboBuscaTipoGasto.NewIndex) = rec!TGT_CODIGO
            rec.MoveNext
        Loop
        CboGastos.ListIndex = 0
        cboBuscaTipoGasto.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        'paso el item seleccionado al tab 'DATOS'
        TxtCodigo.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        TxtCodigo_LostFocus
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

Private Sub GrdModulos_LostFocus()
    GrdModulos.HighLight = flexHighlightNever
End Sub

Private Sub mFechaD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub mFechaH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub


Private Sub tabTB_Click(PreviousTab As Integer)
    'Si cambio de 'Pestaña' en el tab
    'pongo el foco en el primer campo de la misma
    If TabTB.Tab = 0 And Me.Visible Then TxtDescrip.SetFocus
    If TabTB.Tab = 1 Then
        mFechaD.Value = Date
        mFechaH.Value = Date
        cboBuscaTipoGasto.ListIndex = 0
        If mFechaD.Enabled Then mFechaD.SetFocus
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCodigo
End Sub

Private Sub TxtCodigo_LostFocus()
    If Trim(TxtCodigo.Text) <> "" Then ' si no viene vacio
        BuscoDatos
    Else
        cmdGrabar.Enabled = True
        cmdNuevo.Enabled = True
        cmdBorrar.Enabled = False
    End If
End Sub



Private Sub TxtDescrip_GotFocus()
    SelecTexto TxtDescrip
End Sub

Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCodigo_Change()
    If Trim(TxtCodigo.Text) = "" And cmdBorrar.Enabled Then
        cmdBorrar.Enabled = False
    ElseIf Trim(TxtCodigo.Text) <> "" Then
        cmdBorrar.Enabled = True
    End If
End Sub

Private Sub TxtDescrip_Change()
    If Trim(TxtDescrip.Text) = "" And cmdGrabar.Enabled Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If
End Sub

Private Sub txtImporte_GotFocus()
    SelecTexto txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImporte, KeyAscii)
End Sub

Private Sub txtImporte_LostFocus()
    If txtImporte.Text = "" Then
        txtImporte.Text = "0,00"
    Else
        txtImporte.Text = Valido_Importe(txtImporte.Text)
    End If
End Sub




