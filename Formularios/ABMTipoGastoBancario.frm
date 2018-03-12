VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ABMTipoGastoBancario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ABM Tipo de Gasto Bancario"
   ClientHeight    =   3810
   ClientLeft      =   1425
   ClientTop       =   2340
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMTipoGastoBancario.frx":0000
      Height          =   735
      Left            =   4950
      Picture         =   "ABMTipoGastoBancario.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3030
      Width           =   855
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMTipoGastoBancario.frx":0614
      Height          =   735
      Left            =   4080
      Picture         =   "ABMTipoGastoBancario.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3030
      Width           =   855
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMTipoGastoBancario.frx":0C28
      Height          =   735
      Left            =   3210
      Picture         =   "ABMTipoGastoBancario.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3030
      Width           =   855
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMTipoGastoBancario.frx":123C
      Height          =   735
      Left            =   2340
      Picture         =   "ABMTipoGastoBancario.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3030
      Width           =   855
   End
   Begin TabDlg.SSTab TabTB 
      Height          =   2895
      Left            =   60
      TabIndex        =   9
      Top             =   75
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   5106
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
      TabPicture(0)   =   "ABMTipoGastoBancario.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMTipoGastoBancario.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74865
         TabIndex        =   13
         Top             =   375
         Width           =   5475
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   345
            Left            =   4890
            MaskColor       =   &H000000FF&
            Picture         =   "ABMTipoGastoBancario.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   300
            Left            =   1140
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
         Caption         =   " Datos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   135
         TabIndex        =   11
         Top             =   750
         Width           =   5490
         Begin VB.TextBox TxtCodigo 
            Height          =   315
            Left            =   975
            TabIndex        =   0
            Top             =   525
            Width           =   915
         End
         Begin VB.TextBox TxtDescrip 
            Height          =   315
            Left            =   975
            MaxLength       =   50
            TabIndex        =   1
            Top             =   1035
            Width           =   4440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Index           =   3
            Left            =   405
            TabIndex        =   15
            Top             =   570
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   12
            Top             =   1095
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   1635
         Left            =   -74895
         TabIndex        =   8
         Top             =   1125
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   2884
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
      Left            =   150
      TabIndex        =   16
      Top             =   3240
      Width           =   750
   End
End
Attribute VB_Name = "ABMTipoGastoBancario"
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
    If Trim(TxtCodigo) <> "" Then
        resp = MsgBox("Seguro desea eliminar el Tipo de Gasto: " & Trim(TxtDescrip) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = 11
        lblEstado.Caption = "Eliminando ..."
        
        DBConn.Execute "DELETE FROM TIPO_GASTO_BANCARIO WHERE TGB_CODIGO = " & XN(TxtCodigo)
        If TxtDescrip.Enabled Then TxtDescrip.SetFocus
        lblEstado.Caption = ""
        Screen.MousePointer = 1
        cmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    lblEstado.Caption = ""
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = 1
    MsgBox Err.Description, vbCritical, TIT_MSGBOX, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    Set rec = New ADODB.Recordset
    GrdModulos.Rows = 1

    Me.MousePointer = 11
    Me.Refresh
    sql = "SELECT TGB_CODIGO,TGB_DESCRI FROM TIPO_GASTO_BANCARIO "
    If Trim(TxtDescriB) <> "" Then sql = sql & "WHERE TGB_DESCRI LIKE '" & Trim(TxtDescriB) & "%'"
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
    Me.MousePointer = 1
End Sub

Private Sub CmdGrabar_Click()
    On Error GoTo CLAVOSE
    
    If Trim(TxtDescrip) = "" Then
        MsgBox "No ha ingresado el nombre", vbExclamation, TIT_MSGBOX
        If TxtDescrip.Enabled Then TxtDescrip.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Set rec = New ADODB.Recordset
    lblEstado.Caption = "Guardando ..."

    sql = "SELECT TGB_CODIGO FROM TIPO_GASTO_BANCARIO WHERE TGB_CODIGO = " & XN(TxtCodigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        DBConn.Execute "UPDATE TIPO_GASTO_BANCARIO SET TGB_DESCRI = " & XS(TxtDescrip) & _
                       " WHERE TGB_CODIGO = " & XN(TxtCodigo)
    Else
        Set Rec2 = New ADODB.Recordset
        TxtCodigo = "1"
        sql = "SELECT MAX(TGB_CODIGO) as maximo FROM TIPO_GASTO_BANCARIO"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(Rec2.Fields!Maximo) Then TxtCodigo = XN(Rec2.Fields!Maximo) + 1
        Rec2.Close
        DBConn.Execute "INSERT INTO TIPO_GASTO_BANCARIO(TGB_CODIGO,TGB_DESCRI) VALUES(" & XN(TxtCodigo) & "," & XS(TxtDescrip) & ")"
    End If
    rec.Close
    Screen.MousePointer = 1
    cmdNuevo_Click
    Exit Sub
    
CLAVOSE:
    Screen.MousePointer = 1
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub cmdNuevo_Click()
    TxtCodigo.Text = ""
    TxtDescrip.Text = ""
    lblEstado.Caption = ""
    GrdModulos.Rows = 1
    If TxtCodigo.Enabled And Me.Visible Then TxtCodigo.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set ABMTipoGastoBancario = Nothing
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

Private Sub GrdModulos_dblClick()
    If GrdModulos.row > 0 Then
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
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub GrdModulos_LostFocus()
    GrdModulos.HighLight = flexHighlightNever
End Sub

Private Sub tabTB_Click(PreviousTab As Integer)
    'Si cambio de 'Pestaña' en el tab
    'pongo el foco en el primer campo de la misma
    If TabTB.Tab = 0 And Me.Visible Then TxtCodigo.SetFocus
    If TabTB.Tab = 1 Then
        TxtCodigoB.Text = ""
        TxtDescriB.Text = ""
        If TxtDescriB.Enabled Then TxtDescriB.SetFocus
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If TxtCodigo.Text <> "" Then
        sql = "SELECT * FROM TIPO_GASTO_BANCARIO"
        sql = sql & " WHERE TGB_CODIGO=" & XN(TxtCodigo.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            TxtDescrip.Text = rec!TGB_DESCRI
            CmdBorrar.Enabled = True
        Else
            MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
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
    If KeyAscii = vbKeyReturn And CmdGrabar.Enabled Then CmdGrabar_Click
    KeyAscii = Mayuscula(KeyAscii)
End Sub

Private Sub TxtCodigo_Change()
    If Trim(TxtCodigo) = "" And CmdBorrar.Enabled Then
        CmdBorrar.Enabled = False
    ElseIf Trim(TxtCodigo) <> "" Then
        CmdBorrar.Enabled = True
    End If
End Sub

Private Sub TxtDescrip_Change()
    If Trim(TxtDescrip) = "" And CmdGrabar.Enabled Then
        CmdGrabar.Enabled = False
    Else
        CmdGrabar.Enabled = True
    End If
End Sub
