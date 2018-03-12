VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ABMLinea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ABM de Lineas"
   ClientHeight    =   4245
   ClientLeft      =   4065
   ClientTop       =   3105
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   720
      Left            =   4950
      Picture         =   "ABMLinea.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   840
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMLinea.frx":030A
      Enabled         =   0   'False
      Height          =   720
      Left            =   4095
      Picture         =   "ABMLinea.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   840
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   720
      Left            =   3240
      Picture         =   "ABMLinea.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   840
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "ABMLinea.frx":0C28
      Enabled         =   0   'False
      Height          =   720
      Left            =   2385
      Picture         =   "ABMLinea.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   840
   End
   Begin TabDlg.SSTab TabTB 
      Height          =   3240
      Left            =   60
      TabIndex        =   9
      Top             =   180
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   5715
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
      TabPicture(0)   =   "ABMLinea.frx":123C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMLinea.frx":1258
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74790
         TabIndex        =   13
         Top             =   360
         Width           =   5340
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   345
            Left            =   4770
            MaskColor       =   &H000000FF&
            Picture         =   "ABMLinea.frx":1274
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar"
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   435
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   330
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   6
            Top             =   225
            Width           =   3570
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
         Caption         =   " Datos de Linea"
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
         Left            =   225
         TabIndex        =   11
         Top             =   615
         Width           =   5295
         Begin VB.TextBox TxtCodigo 
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox TxtDescrip 
            Height          =   315
            Left            =   1260
            MaxLength       =   30
            TabIndex        =   1
            Top             =   1185
            Width           =   3405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Index           =   3
            Left            =   525
            TabIndex        =   16
            Top             =   645
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   12
            Top             =   1230
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   1980
         Left            =   -74820
         TabIndex        =   8
         Top             =   1140
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   3493
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
      Left            =   105
      TabIndex        =   17
      Top             =   3540
      Width           =   750
   End
End
Attribute VB_Name = "ABMLinea"
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
        resp = MsgBox("Seguro desea eliminar la Linea: " & Trim(TxtDescrip) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = 11
        lblEstado.Caption = "Eliminando ..."
        
        DBConn.Execute "DELETE FROM LINEAS WHERE LNA_CODIGO = " & XN(TxtCodigo)
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
    GrdModulos.Rows = 1

    Screen.MousePointer = 11
    Me.Refresh
    sql = "SELECT LNA_CODIGO,LNA_DESCRI FROM LINEAS "
    If Trim(TxtDescriB) <> "" Then sql = sql & "WHERE LNA_DESCRI LIKE '%" & Trim(TxtDescriB) & "%'"
    sql = sql & " ORDER BY LNA_DESCRI"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
       Do While Not rec.EOF
          GrdModulos.AddItem Trim(rec.Fields(0)) & Chr(9) & Trim(rec.Fields(1))
          rec.MoveNext
       Loop
       If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        MsgBox "No se encontraron items con esta descripción !", vbExclamation, TIT_MSGBOX
        If TxtDescriB.Enabled Then TxtDescriB.SetFocus
    End If
    rec.Close
    Screen.MousePointer = 1
End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo CLAVOSE
    
    If Trim(TxtDescrip) = "" Then
        MsgBox "No ha ingresado la descripción !", vbExclamation, TIT_MSGBOX
        If TxtDescrip.Enabled Then TxtDescrip.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    lblEstado.Caption = "Guardando ..."
    
    If TxtCodigo.Text <> "" Then
        sql = "UPDATE LINEAS "
        sql = sql & " SET LNA_DESCRI = " & XS(TxtDescrip)
        sql = sql & " WHERE LNA_CODIGO = " & XN(TxtCodigo)
        DBConn.Execute sql
        
    Else
        TxtCodigo = "1"
        sql = "SELECT MAX(LNA_CODIGO) as maximo FROM LINEAS"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = XN(rec.Fields!Maximo) + 1
        rec.Close
        
        sql = "INSERT INTO LINEAS(LNA_CODIGO,LNA_DESCRI) "
        sql = sql & " VALUES ("
        sql = sql & XN(TxtCodigo) & ","
        sql = sql & XS(TxtDescrip) & ")"
        DBConn.Execute sql
    End If
    Screen.MousePointer = 1
    CmdNuevo_Click
    Exit Sub
    
CLAVOSE:
    lblEstado.Caption = ""
    Screen.MousePointer = 1
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdNuevo_Click()
    TabTB.Tab = 0
    TxtCodigo.Text = ""
    TxtDescrip.Text = ""
    lblEstado.Caption = ""
    GrdModulos.Rows = 1
    TxtCodigo.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set ABMLinea = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then TabTB.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'si presiono ESCAPE salgo del form
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
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
    If KeyCode = vbKeyDelete Then CmdBorrar_Click
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub


Private Sub GrdModulos_LostFocus()
    GrdModulos.HighLight = flexHighlightNever
End Sub

Private Sub tabTB_Click(PreviousTab As Integer)
    'Si cambio de 'Pestaña' en el tab
    'pongo el foco en el primer campo de la misma
    If TabTB.Tab = 0 And Me.Visible Then TxtDescrip.SetFocus
    If TabTB.Tab = 1 Then
        TxtCodigoB.Text = ""
        TxtDescriB.Text = ""
        If TxtDescriB.Enabled Then TxtDescriB.SetFocus
        cmdGrabar.Enabled = False
        cmdBorrar.Enabled = False
    Else
        cmdGrabar.Enabled = True
        cmdBorrar.Enabled = True
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
        sql = "SELECT * FROM LINEAS"
        sql = sql & " WHERE LNA_CODIGO=" & XN(TxtCodigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
         TxtDescrip.Text = rec!LNA_DESCRI
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

Private Sub TxtDescriB_GotFocus()
    SelecTexto TxtDescriB
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then CmdBuscAprox_Click
End Sub

Private Sub TxtDescrip_GotFocus()
    SelecTexto TxtDescrip
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
