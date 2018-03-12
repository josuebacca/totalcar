VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ABMEstadoCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ABM de Estado de Cheques"
   ClientHeight    =   4170
   ClientLeft      =   1290
   ClientTop       =   1065
   ClientWidth     =   5895
   Icon            =   "ABMEstadoCheques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMEstadoCheques.frx":0442
      Height          =   750
      Left            =   4920
      Picture         =   "ABMEstadoCheques.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3375
      Width           =   885
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      DisabledPicture =   "ABMEstadoCheques.frx":0A56
      Height          =   750
      Left            =   4020
      Picture         =   "ABMEstadoCheques.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3375
      Width           =   885
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMEstadoCheques.frx":106A
      Height          =   750
      Left            =   3120
      Picture         =   "ABMEstadoCheques.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3375
      Width           =   885
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "ABMEstadoCheques.frx":167E
      Height          =   750
      Left            =   2220
      Picture         =   "ABMEstadoCheques.frx":1988
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3375
      Width           =   885
   End
   Begin TabDlg.SSTab TabTB 
      Height          =   3165
      Left            =   60
      TabIndex        =   9
      Top             =   165
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   5583
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
      TabPicture(0)   =   "ABMEstadoCheques.frx":1C92
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMEstadoCheques.frx":1CAE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74865
         TabIndex        =   13
         Top             =   375
         Width           =   5475
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   345
            Left            =   4845
            MaskColor       =   &H000000FF&
            Picture         =   "ABMEstadoCheques.frx":1CCA
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   420
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   315
            Left            =   1125
            MaxLength       =   15
            TabIndex        =   6
            Top             =   225
            Width           =   3660
         End
         Begin VB.TextBox TxtCodigoB 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   10
            Top             =   225
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   14
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Datos del Estado de Cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   360
         TabIndex        =   11
         Top             =   675
         Width           =   4920
         Begin VB.TextBox TxtCodigo 
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            Top             =   465
            Width           =   900
         End
         Begin VB.TextBox TxtDescrip 
            Height          =   315
            Left            =   1260
            MaxLength       =   30
            TabIndex        =   1
            Top             =   1035
            Width           =   3405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Index           =   3
            Left            =   525
            TabIndex        =   15
            Top             =   495
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   12
            Top             =   1080
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   1920
         Left            =   -74910
         TabIndex        =   8
         Top             =   1125
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   3387
         _Version        =   393216
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
      Left            =   90
      TabIndex        =   16
      Top             =   3555
      Width           =   750
   End
End
Attribute VB_Name = "ABMEstadoCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As ADODB.Recordset
Dim sql As String
Dim resp As Integer

Private Sub cmdBorrar_Click()
    On Error GoTo clavose
    If Trim(TxtCodigo) <> "" Then
    
        sql = "SELECT ECH_CODIGO FROM CHEQUE WHERE ECH_CODIGO = " & XN(TxtCodigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
            MsgBox "No se puede eliminar este ESTADO ya que tiene CHEQUES asociados !", vbExclamation, TIT_MSGBOX
            rec.Close
            Exit Sub
        End If
        rec.Close
        
        resp = MsgBox("Seguro desea eliminar el Estado: " & Trim(TxtDescrip) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = 11
        lblEstado.Caption = "Eliminando ..."
        
        DBConn.Execute "DELETE FROM ESTADO_CHEQUE WHERE ECH_CODIGO = " & XN(TxtCodigo)
        If TxtDescrip.Enabled Then TxtDescrip.SetFocus
        lblEstado.Caption = ""
        Screen.MousePointer = 1
        CmdNuevo_Click
    End If
    Exit Sub
    
clavose:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = 1
    Mensaje 2
End Sub

Private Sub CmdBuscAprox_Click()
    Set rec = New ADODB.Recordset
    GrdModulos.Rows = 1

    Screen.MousePointer = 11
    Me.Refresh
    sql = "SELECT ECH_CODIGO,ECH_DESCRI FROM ESTADO_CHEQUE "
    If Trim(TxtDescriB) <> "" Then sql = sql & "WHERE ECH_DESCRI LIKE '%" & Trim(TxtDescriB) & "%'"
    sql = sql & " ORDER BY ECH_DESCRI"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            GrdModulos.AddItem Trim(rec.Fields(0)) & Chr(9) & Trim(rec.Fields(1))
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
    Screen.MousePointer = 1
End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo clavose
    
    If Trim(TxtDescrip) = "" Then
        MsgBox "No ha ingresado la descripción !", vbExclamation, TIT_MSGBOX
        If TxtDescrip.Enabled Then TxtDescrip.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Set rec = New ADODB.Recordset
    lblEstado.Caption = "Guardando ..."

    sql = "SELECT ECH_DESCRI FROM ESTADO_CHEQUE WHERE ECH_CODIGO = " & XN(TxtCodigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        DBConn.Execute "UPDATE ESTADO_CHEQUE SET ECH_DESCRI = '" & Trim(TxtDescrip) & "' " & _
        "WHERE ECH_CODIGO = " & XN(TxtCodigo)
    Else
        Set Rec2 = New ADODB.Recordset
        TxtCodigo = "1"
        sql = "SELECT MAX(ECH_CODIGO) as maximo FROM ESTADO_CHEQUE"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(Rec2.Fields!MAXIMO) Then TxtCodigo = XN(Rec2.Fields!MAXIMO) + 1
        Rec2.Close
        DBConn.Execute "INSERT INTO ESTADO_CHEQUE(ECH_CODIGO,ECH_DESCRI) VALUES " & _
        "(" & XN(TxtCodigo) & "," & XS(TxtDescrip) & ")"
    End If
    rec.Close
    Screen.MousePointer = 1
    CmdNuevo_Click
    Exit Sub
    
clavose:
    Screen.MousePointer = 1
    Mensaje 1
    
End Sub


Private Sub CmdNuevo_Click()
    TabTB.Tab = 0
    TxtCodigo.Text = ""
    TxtDescrip.Text = ""
    lblEstado.Caption = ""
    GrdModulos.Rows = 1
    TxtCodigo.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set ABMEstadoCheques = Nothing
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
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)

    lblEstado.Caption = ""
    GrdModulos.FormatString = "Código|Descripción"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 4500
    GrdModulos.Rows = 1
    TabTB.Tab = 0
End Sub



Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
        'paso el item seleccionado al tab 'DATOS'
        GrdModulos.Col = 0
        TxtCodigo.Text = GrdModulos.Text
        GrdModulos.Col = 1
        TxtDescrip.Text = Trim(GrdModulos.Text)
        
        If TxtDescrip.Enabled Then TxtDescrip.SetFocus
        TabTB.Tab = 0
    End If
End Sub

Private Sub GrdModulos_GotFocus()
    GrdModulos.Col = 0
    GrdModulos.ColSel = 1
    GrdModulos.HighLight = flexHighlightAlways
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmdBorrar_Click
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub GrdModulos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub GrdModulos_LostFocus()
    GrdModulos.HighLight = flexHighlightNever
End Sub

Private Sub TABTB_Click(PreviousTab As Integer)
    'Si cambio de 'Pestaña' en el tab
    'pongo el foco en el primer campo de la misma
    If TabTB.Tab = 0 And Me.Visible Then
     TxtDescrip.SetFocus
     cmdGrabar.Enabled = True
     cmdBorrar.Enabled = True
    End If
    If TabTB.Tab = 1 Then
        TxtCodigoB.Text = ""
        TxtDescriB.Text = ""
        If TxtDescriB.Enabled Then TxtDescriB.SetFocus
        cmdGrabar.Enabled = False
        cmdBorrar.Enabled = False
    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodigo_LostFocus()
    If TxtCodigo.Text <> "" Then
        sql = "SELECT ECH_CODIGO,ECH_DESCRI "
        sql = sql & " FROM ESTADO_CHEQUE "
        sql = sql & " WHERE ECH_CODIGO=" & XN(TxtCodigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
         TxtDescrip.Text = rec!ECH_DESCRI
        Else
         MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
         TxtCodigo.Text = ""
         TxtCodigo.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub TxtCodigoB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then CmdBuscAprox_Click
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then CmdBuscAprox_Click
End Sub

Private Sub TxtDescrip_GotFocus()
    Set rec = New ADODB.Recordset
    TxtDescrip.SelStart = 0
    TxtDescrip.SelLength = Len(TxtDescrip)
End Sub

Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cmdGrabar.Enabled Then cmdGrabar_Click
    KeyAscii = Mayuscula(KeyAscii)
End Sub

Private Sub TxtCodigo_Change()
    If Trim(TxtCodigo) = "" And cmdBorrar.Enabled Then
        cmdBorrar.Enabled = False
    ElseIf Trim(TxtCodigo) <> "" Then
        cmdBorrar.Enabled = True
    End If
End Sub

Private Sub TxtDescrip_Change()
    If Trim(TxtDescrip) = "" And cmdGrabar.Enabled Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If
End Sub
