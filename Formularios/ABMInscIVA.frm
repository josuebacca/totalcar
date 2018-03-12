VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ABMInscIVA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ABM de Condición IVA"
   ClientHeight    =   4365
   ClientLeft      =   855
   ClientTop       =   2355
   ClientWidth     =   6510
   ForeColor       =   &H00C0C0C0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4365
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMInscIVA.frx":0000
      Height          =   720
      Left            =   3660
      Picture         =   "ABMInscIVA.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   915
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMInscIVA.frx":0614
      Height          =   720
      Left            =   2730
      Picture         =   "ABMInscIVA.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   915
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMInscIVA.frx":0C28
      Height          =   720
      Left            =   5520
      Picture         =   "ABMInscIVA.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   915
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMInscIVA.frx":123C
      Height          =   720
      Left            =   4590
      Picture         =   "ABMInscIVA.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   915
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   3420
      Left            =   60
      TabIndex        =   11
      Top             =   135
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   6033
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
      TabPicture(0)   =   "ABMInscIVA.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMInscIVA.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74865
         TabIndex        =   16
         Top             =   405
         Width           =   6135
         Begin VB.TextBox TxtDescriB 
            Height          =   300
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   7
            Top             =   225
            Width           =   4215
         End
         Begin VB.TextBox TxtCodigoB 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   17
            Top             =   225
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   345
            Left            =   5550
            MaskColor       =   &H000000FF&
            Picture         =   "ABMInscIVA.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   19
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1665
            TabIndex        =   18
            Top             =   315
            Visible         =   0   'False
            Width           =   540
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   " Datos de la Condición IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Left            =   315
         TabIndex        =   10
         Top             =   570
         Width           =   5625
         Begin VB.TextBox txtPorcentaje 
            Height          =   300
            Left            =   1365
            MaxLength       =   6
            TabIndex        =   2
            Top             =   1785
            Width           =   1005
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Left            =   1365
            TabIndex        =   0
            Top             =   735
            Width           =   1005
         End
         Begin VB.TextBox txtdescri 
            Height          =   300
            Left            =   1365
            MaxLength       =   30
            TabIndex        =   1
            Tag             =   "Descripción"
            Top             =   1260
            Width           =   3375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje:"
            Height          =   195
            Left            =   405
            TabIndex        =   20
            Top             =   1800
            Width           =   810
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   675
            TabIndex        =   15
            Top             =   765
            Width           =   540
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   330
            TabIndex        =   14
            Top             =   1305
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   2070
         Left            =   -74910
         TabIndex        =   9
         Top             =   1215
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   3651
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   12
         Top             =   570
         Width           =   1065
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
      TabIndex        =   13
      Top             =   3750
      Width           =   750
   End
End
Attribute VB_Name = "ABMInscIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Actualizar()
    
    On Error GoTo ErrorTrans
    sql = "UPDATE CONDICION_IVA "
    sql = sql & " SET IVA_DESCRI = " & XS(txtdescri.Text)
    sql = sql & ",IVA_PORCEN=" & XN(txtPorcentaje.Text)
    sql = sql & " WHERE IVA_CODIGO = " & XN(TxtCodigo.Text)
    DBConn.BeginTrans
    DBConn.Execute sql, dbExecDirect
    DBConn.CommitTrans
    Exit Sub
    
ErrorTrans:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    Exit Sub
End Sub

Private Sub Insertar()

    On Error GoTo ErrorTrans
    If TxtCodigo.Text = "" Then ' Si está VACIO
        TxtCodigo.Text = "1"
        sql = "select max(IVA_codigo) as maximo from CONDICION_IVA "
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = Val(Trim(rec.Fields!Maximo)) + 1
        rec.Close
    End If
    DBConn.BeginTrans
    sql = "INSERT INTO CONDICION_IVA (IVA_CODIGO, IVA_DESCRI,IVA_PORCEN) "
    sql = sql & " VALUES ( " & XN(TxtCodigo) & " ," & XS(txtdescri.Text) & ","
    sql = sql & XN(txtPorcentaje.Text) & ")"
    DBConn.Execute sql, dbExecDirect
    DBConn.CommitTrans
    Exit Sub
ErrorTrans:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    Exit Sub
End Sub


Function LimpiarControles()
    
    TxtCodigo = ""
    txtdescri.Text = ""
    txtPorcentaje.Text = ""
    txtdescri.Enabled = True
    TxtCodigo.SetFocus
    cmdGrabar.Enabled = False
    cmdBorrar.Enabled = False
    cmdNuevo.Enabled = True
    lblEstado.Caption = ""
End Function

Function ValidarIngreso()
Dim MensCampos As String
ValidarIngreso = True
For Each Control In ABMInscIVA.Controls ' revisar los controles del form
    If TypeOf Control Is TextBox _
        Or TypeOf Control Is ListBox _
        Or TypeOf Control Is ComboBox Then ' si el control es de carga o selección de datos
            If Trim(Control.Tag) <> "" Then  'si el tag no está vacio, es un campo necesario
                If Trim(Control.Text) = "" Then ' dejaron vacio un campo necesario
                    MensCampos = MensCampos & Chr(13) & Control.Tag
                    ValidarIngreso = False
                End If
            End If
    End If
Next Control

If MensCampos <> "" Then ' si hay mensaje es que hay campos incompletos
    Beep
    MsgBox "Debe completar los siguientes campos:" & MensCampos, vbOKOnly + vbInformation, TIT_MSGBOX
    txtdescri.SetFocus
End If
End Function

Private Sub CmdBorrar_Click()
    If TxtCodigo.Text <> "" Then
        On Error GoTo CLAVOSE
        resp = MsgBox("Seguro desea eliminar la Condición de IVA: " & Trim(txtdescri.Text) & " ?", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        Screen.MousePointer = 11
        lblEstado.Caption = "Eliminando ..."
        
        sql = "DELETE FROM CONDICION_IVA WHERE IVA_CODIGO = " & XN(TxtCodigo.Text)
        DBConn.BeginTrans
        DBConn.Execute sql
        DBConn.CommitTrans
        lblEstado.Caption = ""
        Screen.MousePointer = 1
        LimpiarControles
    End If
    Exit Sub
CLAVOSE:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    Screen.MousePointer = 1
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1

    Screen.MousePointer = 11
    Me.Refresh
    sql = "SELECT * FROM CONDICION_IVA"
    If Trim(TxtDescriB) <> "" Then sql = sql & "WHERE IVA_descri LIKE '" & Trim(TxtDescriB) & "%'"
    sql = sql & " ORDER BY IVA_DESCRI"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            GrdModulos.AddItem rec.Fields(0) & Chr(9) & rec.Fields(1) & Chr(9) & rec.Fields(2)
            rec.MoveNext
        Loop
        If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        MsgBox "No se encontraron items con esta descripcion !", vbExclamation, TIT_MSGBOX
        TxtDescriB.SelStart = 0
        TxtDescriB.SelLength = Len(TxtDescriB)
        If TxtDescriB.Enabled Then TxtDescriB.SetFocus
    End If
    rec.Close
    Screen.MousePointer = 1

End Sub

Private Sub CmdNuevo_Click()
    tabDatos.Tab = 0
    LimpiarControles
    GrdModulos.Rows = 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set ABMInscIVA = Nothing
End Sub

Private Sub cmdGrabar_Click()
    lblEstado.Caption = "Grabando..."
    If txtdescri.Text = "" Then
        MsgBox "Debe ingresar la descripción", vbExclamation, TIT_MSGBOX
        txtdescri.SetFocus
        lblEstado.Caption = ""
        Exit Sub
    End If
    If TxtCodigo.Text <> "" Then
        Actualizar
    Else
        Insertar
    End If
    LimpiarControles
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    lblEstado.Caption = ""
    Set rec = New ADODB.Recordset
    GrdModulos.FormatString = "Código|Descripción|Porcent."
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 4900
    GrdModulos.ColWidth(2) = 0
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    Centrar_pantalla Me
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
        'paso el item seleccionado al tab 'DATOS'
        TxtCodigo.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0))
        txtdescri.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 1))
        txtPorcentaje.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 2))
        cmdBorrar.Enabled = True
        cmdGrabar.Enabled = True
        If txtdescri.Enabled Then txtdescri.SetFocus
        tabDatos.Tab = 0
    End If
End Sub

Private Sub GrdModulos_GotFocus()
    GrdModulos.Col = 0
    GrdModulos.ColSel = 1
    GrdModulos.HighLight = flexHighlightAlways
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub GrdModulos_LostFocus()
    GrdModulos.HighLight = flexHighlightNever
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 0 And Me.Visible Then TxtCodigo.SetFocus
    If tabDatos.Tab = 1 Then
       GrdModulos.Rows = 1
       GrdModulos.Refresh
       cmdBorrar.Enabled = False
       cmdGrabar.Enabled = False
       TxtDescriB.SetFocus
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If Me.TxtCodigo.Text <> "" Then
        sql = "select * from condicion_iva where iva_codigo =  " & XN(TxtCodigo.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
           txtdescri.Text = Trim(rec!IVA_DESCRI)
           txtPorcentaje.Text = Format((rec!IVA_PORCEN), "0.00")
           cmdBorrar.Enabled = True
           cmdGrabar.Enabled = True
        Else
           MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
           TxtCodigo.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtDescri_Change()
    If Trim(txtdescri) = "" Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If
End Sub

Private Sub txtdescri_GotFocus()
   SelecTexto txtdescri
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
   KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtPorcentaje_GotFocus()
    SelecTexto txtPorcentaje
End Sub

Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorcentaje, KeyAscii)
End Sub

Private Sub txtPorcentaje_LostFocus()
  If txtPorcentaje.Text <> "" Then
    If ValidarPorcentaje(txtPorcentaje) = False Then
     txtPorcentaje.Text = ""
     txtPorcentaje.SetFocus
    End If
  End If
End Sub
