VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ABMEstadoDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ABM Estado de Documento"
   ClientHeight    =   4230
   ClientLeft      =   2115
   ClientTop       =   2430
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5895
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMEstadoDocumento.frx":0000
      Height          =   750
      Left            =   4935
      Picture         =   "ABMEstadoDocumento.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3405
      Width           =   855
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      DisabledPicture =   "ABMEstadoDocumento.frx":0614
      Height          =   750
      Left            =   4065
      Picture         =   "ABMEstadoDocumento.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3405
      Width           =   855
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMEstadoDocumento.frx":0C28
      Height          =   750
      Left            =   3195
      Picture         =   "ABMEstadoDocumento.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3405
      Width           =   855
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "ABMEstadoDocumento.frx":123C
      Height          =   750
      Left            =   2325
      Picture         =   "ABMEstadoDocumento.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3405
      Width           =   855
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
      TabPicture(0)   =   "ABMEstadoDocumento.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMEstadoDocumento.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74790
         TabIndex        =   12
         Top             =   525
         Width           =   5280
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   330
            Left            =   4680
            MaskColor       =   &H000000FF&
            Picture         =   "ABMEstadoDocumento.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   300
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   6
            Top             =   240
            Width           =   3465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   13
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Datos del Estado de Recibo"
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
         Left            =   345
         TabIndex        =   10
         Top             =   600
         Width           =   4920
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Left            =   1260
            TabIndex        =   0
            Top             =   495
            Width           =   945
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
            TabIndex        =   14
            Top             =   540
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   11
            Top             =   1080
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   1665
         Left            =   -74820
         TabIndex        =   8
         Top             =   1350
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   2937
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
      TabIndex        =   15
      Top             =   3585
      Width           =   750
   End
End
Attribute VB_Name = "ABMEstadoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim resp As Integer

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(TxtCODIGO) <> "" Then
        resp = MsgBox("Seguro desea eliminar el Estado: " & Trim(TxtDescrip) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        DBConn.BeginTrans
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Eliminando..."
        DBConn.Execute "DELETE FROM ESTADO_DOCUMENTO WHERE EST_CODIGO = " & XN(TxtCODIGO)
        If TxtDescrip.Enabled Then TxtDescrip.SetFocus
        lblEstado.Caption = ""
        DBConn.CommitTrans
        Screen.MousePointer = vbNormal
        cmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    Set rec = New ADODB.Recordset
    GrdModulos.Rows = 1

    Screen.MousePointer = vbHourglass
    Me.Refresh
    sql = "SELECT EST_CODIGO,EST_DESCRI FROM ESTADO_DOCUMENTO "
    If Trim(TxtDescriB) <> "" Then sql = sql & "WHERE EST_DESCRI LIKE '%" & Trim(TxtDescriB) & "%'"
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
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdGrabar_Click()
    On Error GoTo CLAVOSE
    
    If Trim(TxtDescrip) = "" Then
        MsgBox "No ha ingresado la descripción", vbExclamation, TIT_MSGBOX
        TxtDescrip.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    DBConn.BeginTrans
    If TxtCODIGO.Text <> "" Then
        
        sql = "UPDATE ESTADO_DOCUMENTO "
        sql = sql & " SET EST_DESCRI =" & XS(TxtDescrip)
        sql = sql & " WHERE EST_CODIGO = " & XN(TxtCODIGO)
        DBConn.Execute sql
    Else
        TxtCODIGO = "1"
        sql = "SELECT MAX(EST_CODIGO) as maximo FROM ESTADO_DOCUMENTO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(rec.Fields!Maximo) Then TxtCODIGO = XN(rec.Fields!Maximo) + 1
        rec.Close
        
        sql = "INSERT INTO ESTADO_DOCUMENTO(EST_CODIGO,EST_DESCRI)"
        sql = sql & " VALUES ("
        sql = sql & XN(TxtCODIGO)
        sql = sql & "," & XS(TxtDescrip)
        sql = sql & ")"
        DBConn.Execute sql
    End If
    Screen.MousePointer = vbNormal
    DBConn.CommitTrans
    cmdNuevo_Click
    Exit Sub
    
CLAVOSE:
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.RollbackTrans
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
    Set ABMEstadoDocumento = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF1 Then TabTB.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'si presiono ESCAPE salgo del form
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
    'si presiono CTRL + B hago la busqueda aprox
    If KeyAscii = vbKeyReturn Then 'avanza de campo
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

    Call Centrar_pantalla(Me)
    Set rec = New ADODB.Recordset
    
    
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
        TxtCODIGO.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
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
     TxtDescrip.SetFocus
     cmdGrabar.Enabled = True
     cmdBorrar.Enabled = True
    End If
    If TabTB.Tab = 1 Then
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
        sql = "SELECT EST_CODIGO,EST_DESCRI "
        sql = sql & " FROM ESTADO_DOCUMENTO"
        sql = sql & " WHERE EST_CODIGO=" & XN(TxtCODIGO)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
         TxtDescrip.Text = rec!EST_DESCRI
        Else
         MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
         TxtCODIGO.Text = ""
         TxtCODIGO.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub TxtDescrip_GotFocus()
    Set rec = New ADODB.Recordset
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
