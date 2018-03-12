VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ABMMoneda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ABM de Moneda"
   ClientHeight    =   4080
   ClientLeft      =   1290
   ClientTop       =   2610
   ClientWidth     =   5940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "ABMMoneda.frx":0000
      Height          =   735
      Left            =   2250
      Picture         =   "ABMMoneda.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3285
      Width           =   885
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMMoneda.frx":0614
      Height          =   735
      Left            =   3150
      Picture         =   "ABMMoneda.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3285
      Width           =   885
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMMoneda.frx":0C28
      Height          =   735
      Left            =   4050
      Picture         =   "ABMMoneda.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3285
      Width           =   885
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMMoneda.frx":123C
      Height          =   735
      Left            =   4935
      Picture         =   "ABMMoneda.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3285
      Width           =   885
   End
   Begin TabDlg.SSTab TabTB 
      Height          =   3075
      Left            =   105
      TabIndex        =   9
      Top             =   150
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   5424
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "ABMMoneda.frx":1850
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMMoneda.frx":186C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   195
         TabIndex        =   13
         Top             =   465
         Width           =   5295
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   345
            Left            =   4605
            MaskColor       =   &H000000FF&
            Picture         =   "ABMMoneda.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   315
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   6
            Top             =   225
            Width           =   3375
         End
         Begin VB.TextBox TxtCodigoB 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   10
            Top             =   225
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1665
            TabIndex        =   15
            Top             =   315
            Visible         =   0   'False
            Width           =   540
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
         Caption         =   " Datos de la Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   -74595
         TabIndex        =   11
         Top             =   660
         Width           =   4920
         Begin VB.TextBox TxtCodigo 
            Height          =   315
            Left            =   1140
            TabIndex        =   0
            Top             =   480
            Width           =   870
         End
         Begin VB.TextBox TxtDescrip 
            Height          =   315
            Left            =   1140
            MaxLength       =   20
            TabIndex        =   1
            Top             =   1035
            Width           =   3090
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Index           =   3
            Left            =   525
            TabIndex        =   16
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
         Height          =   1710
         Left            =   180
         TabIndex        =   8
         Top             =   1230
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   3016
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
      Left            =   180
      TabIndex        =   17
      Top             =   3495
      Width           =   750
   End
End
Attribute VB_Name = "ABMMoneda"
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
    If Trim(TxtCODIGO) <> "" Then
    
        sql = "SELECT MON_CODIGO FROM MOVIMIENTO WHERE MON_CODIGO = " & XN(TxtCODIGO)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
            MsgBox "No se puede eliminar esta Moneda ya que tiene Movimientos asociados !", vbExclamation, TIT_MSGBOX
            rec.Close
            Exit Sub
        End If
        rec.Close
        
        resp = MsgBox("Seguro desea eliminar la Moneda: " & Trim(TxtDescrip) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = 11
        lblEstado.Caption = "Eliminando ..."
        
        DBConn.Execute "DELETE FROM MONEDA WHERE MON_CODIGO = " & XN(TxtCODIGO)
        If TxtDescrip.Enabled Then TxtDescrip.SetFocus
        lblEstado.Caption = ""
        Screen.MousePointer = 1
        cmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = 1
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    
    GrdModulos.Rows = 1

    Screen.MousePointer = 11
    Me.Refresh
    sql = "SELECT MON_CODIGO,MON_DESCRI FROM MONEDA "
    If Trim(TxtDescriB) <> "" Then sql = sql & "WHERE MON_DESCRI LIKE '%" & Trim(TxtDescriB) & "%'"
    sql = sql & " ORDER BY MON_DESCRI"
    
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

Private Sub CmdGrabar_Click()
    On Error GoTo CLAVOSE
    
    If Trim(TxtDescrip) = "" Then
        MsgBox "No ha ingresado la descripción !", vbExclamation, TIT_MSGBOX
        If TxtDescrip.Enabled Then TxtDescrip.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    lblEstado.Caption = "Guardando ..."
    
    sql = "SELECT MON_CODIGO FROM MONEDA WHERE MON_CODIGO = " & XN(TxtCODIGO)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        DBConn.Execute "UPDATE MONEDA SET MON_DESCRI = '" & Trim(TxtDescrip) & "' " & _
        "WHERE MON_CODIGO = " & XN(TxtCODIGO)
    Else
        TxtCODIGO = "1"
        sql = "SELECT MAX(MON_CODIGO) as maximo FROM MONEDA"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(Rec2.Fields!Maximo) Then TxtCODIGO = XN(Rec2.Fields!Maximo) + 1
        Rec2.Close
        
        DBConn.Execute "INSERT INTO MONEDA (MON_CODIGO,MON_DESCRI) VALUES " & _
        "(" & XN(TxtCODIGO) & "," & XS(TxtDescrip) & ")"
    End If
    rec.Close
    Screen.MousePointer = 1
    cmdNuevo_Click
    Exit Sub
    
CLAVOSE:
    Screen.MousePointer = 1
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub cmdNuevo_Click()
    TabTB.Tab = 0
    TxtCODIGO.Text = ""
    TxtDescrip.Text = ""
    lblEstado.Caption = ""
    GrdModulos.Rows = 1
    TxtCODIGO.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set ABMMoneda = Nothing
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
    Set Rec2 = New ADODB.Recordset
     
    Call Centrar_pantalla(Me)

    lblEstado.Caption = ""
    GrdModulos.FormatString = "Código|Descripción"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 4000
    GrdModulos.Rows = 1
    TabTB.Tab = 0
End Sub

Private Sub GrdModulos_dblClick()
    If GrdModulos.row > 0 Then
        'paso el item seleccionado al tab 'DATOS'
        GrdModulos.Col = 0
        TxtCODIGO.Text = GrdModulos.Text
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
    If KeyCode = vbKeyDelete Then CmdBorrar_Click
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub GrdModulos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub GrdModulos_LostFocus()
    GrdModulos.HighLight = flexHighlightNever
End Sub

Private Sub tabTB_Click(PreviousTab As Integer)
    'Si cambio de 'Pestaña' en el tab
    'pongo el foco en el primer campo de la misma
    If TabTB.Tab = 0 And Me.Visible Then
     TxtCODIGO.SetFocus
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

Private Sub TxtCodigo_LostFocus()
    If TxtCODIGO.Text <> "" Then
        sql = "SELECT MON_CODIGO,MON_DESCRI FROM MONEDA "
        sql = sql & " WHERE MON_CODIGO=" & XN(TxtCODIGO)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
         TxtDescrip.Text = rec!MON_DESCRI
        Else
         MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
         TxtCODIGO.Text = ""
         TxtCODIGO.SetFocus
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
    'Set rec = New ADODB.Recordset
    TxtDescrip.SelStart = 0
    TxtDescrip.SelLength = Len(TxtDescrip)
End Sub

Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cmdGrabar.Enabled Then CmdGrabar_Click
    KeyAscii = Mayuscula(KeyAscii)
End Sub

Private Sub TxtCodigo_Change()
    If Trim(TxtCODIGO) = "" And cmdBorrar.Enabled Then
        cmdBorrar.Enabled = False
    ElseIf Trim(TxtCODIGO) <> "" Then
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
