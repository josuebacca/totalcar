VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form ABMComision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM de Comisiones"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMComision.frx":0000
      Height          =   750
      Left            =   3645
      Picture         =   "ABMComision.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3645
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMComision.frx":0614
      Height          =   750
      Left            =   2760
      Picture         =   "ABMComision.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3645
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMComision.frx":0C28
      Height          =   750
      Left            =   5400
      Picture         =   "ABMComision.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3645
      Width           =   870
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMComision.frx":123C
      Height          =   750
      Left            =   4515
      Picture         =   "ABMComision.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3645
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   3510
      Left            =   45
      TabIndex        =   12
      Top             =   75
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   6191
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   529
      ForeColor       =   -2147483630
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
      TabPicture(0)   =   "ABMComision.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMComision.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraDatos 
         Caption         =   "Datos de la Comision del Vendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   285
         TabIndex        =   11
         Top             =   570
         Width           =   5580
         Begin VB.TextBox txtPorCob 
            Height          =   315
            Left            =   1500
            MaxLength       =   6
            TabIndex        =   3
            Tag             =   "Descripción"
            Top             =   2055
            Width           =   885
         End
         Begin VB.ComboBox CboRep 
            Height          =   315
            Left            =   1500
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   990
            Width           =   2805
         End
         Begin VB.ComboBox CboVend 
            Height          =   315
            Left            =   1515
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   465
            Width           =   2805
         End
         Begin VB.TextBox txtPorVta 
            Height          =   315
            Left            =   1500
            MaxLength       =   6
            TabIndex        =   2
            Tag             =   "Descripción"
            Top             =   1530
            Width           =   885
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "% Cobranza:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   495
            TabIndex        =   20
            Top             =   2100
            Width           =   885
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Representada:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   17
            Top             =   1050
            Width           =   1050
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   645
            TabIndex        =   16
            Top             =   510
            Width           =   735
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "% Venta:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   750
            TabIndex        =   15
            Top             =   1575
            Width           =   630
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74850
         TabIndex        =   13
         Top             =   390
         Width           =   5940
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   345
            Left            =   5310
            MaskColor       =   &H000000FF&
            Picture         =   "ABMComision.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Buscar"
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   315
            Left            =   1275
            MaxLength       =   15
            TabIndex        =   8
            Top             =   255
            Width           =   3975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   270
            Width           =   735
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   2205
         Left            =   -74865
         TabIndex        =   10
         Top             =   1155
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   3889
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   18
         Top             =   570
         Width           =   1065
      End
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
      Left            =   135
      TabIndex        =   19
      Top             =   3720
      Width           =   750
   End
End
Attribute VB_Name = "ABMComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim resp As Integer

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
        resp = MsgBox("Seguro desea eliminar esta Comisión?", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Eliminando ..."
        sql = "DELETE FROM COMISION "
        sql = sql & "WHERE "
        sql = sql & " VEN_CODIGO=" & CboVend.ItemData(CboVend.ListIndex)
        sql = sql & " AND REP_CODIGO=" & cboRep.ItemData(cboRep.ListIndex)
         DBConn.Execute sql
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        CmdNuevo_Click
    
    Exit Sub
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    MousePointer = vbHourglass
    
    sql = "SELECT V.VEN_NOMBRE,R.REP_RAZSOC,C.COM_PORVTA,C.COM_PORCOB,V.VEN_CODIGO,R.REP_CODIGO"
    sql = sql & " FROM COMISION C,VENDEDOR V,REPRESENTADA R"
    sql = sql & " WHERE C.VEN_CODIGO = V.VEN_CODIGO "
    sql = sql & " AND C.REP_CODIGO= R.REP_CODIGO AND V.VEN_NOMBRE"
    sql = sql & " LIKE '" & TxtDescriB.Text & "%' ORDER BY V.VEN_NOMBRE"
        
    lblEstado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Do While Not rec.EOF
           GrdModulos.AddItem rec.Fields(0) & Chr(9) & rec.Fields(1) & Chr(9) & _
                              Format(rec.Fields(2), "0.00") & Chr(9) & Format(rec.Fields(3), "0.00") _
                              & Chr(9) & rec.Fields(4) & Chr(9) & rec.Fields(5)
           rec.MoveNext
        Loop
        If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        TxtDescriB.SetFocus
    End If
    rec.Close
    MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub CmdGrabar_Click()
    
    If ValidarComision = False Then Exit Sub
    
    On Error GoTo HayError
    DBConn.BeginTrans
    sql = "SELECT * FROM COMISION"
    sql = sql & " WHERE VEN_CODIGO=" & CboVend.ItemData(CboVend.ListIndex)
    sql = sql & " AND REP_CODIGO=" & cboRep.ItemData(cboRep.ListIndex)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    If rec.EOF = False Then
        sql = "UPDATE COMISION"
        sql = sql & " SET COM_PORVTA=" & XN(txtPorVta)
        sql = sql & " ,COM_PORCOB=" & XN(txtPorCob)
        sql = sql & " WHERE  VEN_CODIGO =" & CboVend.ItemData(CboVend.ListIndex)
        sql = sql & " AND REP_CODIGO=" & cboRep.ItemData(cboRep.ListIndex)
        
        DBConn.Execute sql
        
    Else
        sql = "INSERT INTO COMISION(VEN_CODIGO,REP_CODIGO,COM_PORVTA,COM_PORCOB)"
        sql = sql & " VALUES ("
        sql = sql & CboVend.ItemData(CboVend.ListIndex) & ","
        sql = sql & cboRep.ItemData(cboRep.ListIndex) & ","
        sql = sql & XN(txtPorVta) & ","
        sql = sql & XN(txtPorCob) & ")"
        DBConn.Execute sql
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.CommitTrans
    CmdNuevo_Click
    Exit Sub
    
HayError:
    If rec.State = 1 Then rec.Close
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdNuevo_Click()
    txtPorVta.Text = "0,00"
    txtPorCob.Text = "0,00"
    lblEstado.Caption = ""
    GrdModulos.Rows = 1
    CboVend.Enabled = True
    cboRep.Enabled = True
    CboVend.ListIndex = 0
    cboRep.ListIndex = 0
    tabDatos.Tab = 0
    CboVend.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set ABMComision = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    lblEstado.Caption = ""
    'ARMO LA GRILLA DE DATOS
    preparo_grilla
    tabDatos.Tab = 0
    'CARGO COMBO VENDEDOR
    LLenarComboVendedor
    'CARGO COMBO REPRESENTADA
    llenarComboRepresentada
    txtPorVta.Text = "0,00"
    txtPorCob.Text = "0,00"
End Sub

Function preparo_grilla()
    GrdModulos.FormatString = "Vendedor|Representada|>% Venta|>% Cobranza|CODVEN|CODREP"
    GrdModulos.ColWidth(0) = 2500
    GrdModulos.ColWidth(1) = 2500
    GrdModulos.ColWidth(2) = 1000
    GrdModulos.ColWidth(3) = 1000
    GrdModulos.ColWidth(4) = 0
    GrdModulos.ColWidth(5) = 0
    GrdModulos.Rows = 2
End Function

Function LLenarComboVendedor()
    sql = "SELECT * FROM VENDEDOR ORDER BY VEN_NOMBRE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            CboVend.AddItem rec!VEN_NOMBRE
            CboVend.ItemData(CboVend.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        CboVend.ListIndex = 0
    End If
    rec.Close
End Function

Function llenarComboRepresentada()
    sql = "SELECT * FROM REPRESENTADA ORDER BY REP_RAZSOC"
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
End Function

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
           Call BuscaCodigoProxItemData(GrdModulos.TextMatrix(GrdModulos.RowSel, 4), CboVend)
           Call BuscaCodigoProxItemData(GrdModulos.TextMatrix(GrdModulos.RowSel, 5), cboRep)
           txtPorVta.Text = Format(GrdModulos.TextMatrix(GrdModulos.RowSel, 2), "0.00")
           txtPorCob.Text = Format(GrdModulos.TextMatrix(GrdModulos.RowSel, 3), "0.00")
           CmdGrabar.Enabled = True
           CmdBorrar.Enabled = True
           CboVend.Enabled = False
           cboRep.Enabled = False
           txtPorVta.SetFocus
           tabDatos.Tab = 0
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 0 And Me.Visible Then
        CmdGrabar.Enabled = True
        CmdBorrar.Enabled = True
    End If
    If tabDatos.Tab = 1 Then
        TxtDescriB.Text = ""
        If TxtDescriB.Enabled Then TxtDescriB.SetFocus
        CmdGrabar.Enabled = False
        CmdBorrar.Enabled = False
    End If
End Sub

Private Sub TxtDescriB_GotFocus()
    SelecTexto TxtDescriB
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Function ValidarComision() As Boolean
    If CboVend.ListIndex = -1 Then
        MsgBox "No ha seleccionado un Vendedor", vbExclamation, TIT_MSGBOX
        CboVend.SetFocus
        ValidarComision = False
        Exit Function
    End If
    
    If cboRep.ListIndex = -1 Then
        MsgBox "No ha seleccionado una Representada", vbExclamation, TIT_MSGBOX
        cboRep.SetFocus
        ValidarComision = False
        Exit Function
    End If
    
    If txtPorVta.Text = "" Then
        MsgBox "No ha ingresado el porcentaje de la Comisión de Venta", vbExclamation, TIT_MSGBOX
        txtPorVta.SetFocus
        ValidarComision = False
        Exit Function
    End If
    If txtPorCob.Text = "" Then
        MsgBox "No ha ingresado el porcentaje de la Comisión de Cobranza", vbExclamation, TIT_MSGBOX
        txtPorCob.SetFocus
        ValidarComision = False
        Exit Function
    End If
    
    ValidarComision = True
End Function

Private Sub txtPorCob_GotFocus()
    SelecTexto txtPorCob
End Sub

Private Sub txtPorCob_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorCob, KeyAscii)
End Sub

Private Sub txtPorCob_LostFocus()
   If ValidarPorcentaje(txtPorCob) = False Then
     txtPorCob.Text = "0,00"
   End If
End Sub

Private Sub txtPorVta_GotFocus()
    SelecTexto txtPorVta
End Sub

Private Sub txtPorVta_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorVta, KeyAscii)
End Sub

Private Sub txtPorVta_LostFocus()
   If ValidarPorcentaje(txtPorVta) = False Then
    txtPorVta.Text = "0,00"
   End If
End Sub
