VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ABMServicios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ABM de Srevicios"
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
      DisabledPicture =   "ABMServicios.frx":0000
      Height          =   720
      Left            =   3660
      Picture         =   "ABMServicios.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   915
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMServicios.frx":0614
      Height          =   720
      Left            =   2730
      Picture         =   "ABMServicios.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   915
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMServicios.frx":0C28
      Height          =   720
      Left            =   5520
      Picture         =   "ABMServicios.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   915
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMServicios.frx":123C
      Height          =   720
      Left            =   4590
      Picture         =   "ABMServicios.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   915
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   3420
      Left            =   60
      TabIndex        =   10
      Top             =   120
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   6033
      _Version        =   327681
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
      TabPicture(0)   =   "ABMServicios.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMServicios.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74865
         TabIndex        =   15
         Top             =   405
         Width           =   6135
         Begin VB.TextBox TxtDescriB 
            Height          =   315
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   7
            Top             =   240
            Width           =   4215
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   345
            Left            =   5490
            MaskColor       =   &H000000FF&
            Picture         =   "ABMServicios.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Buscar"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   16
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   " Datos del Servicio"
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
         TabIndex        =   9
         Top             =   570
         Width           =   5625
         Begin VB.TextBox txtPrecio 
            Height          =   300
            Left            =   1365
            MaxLength       =   10
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
            MaxLength       =   40
            TabIndex        =   1
            Tag             =   "Descripción"
            Top             =   1260
            Width           =   3375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Precio:"
            Height          =   195
            Left            =   720
            TabIndex        =   17
            Top             =   1800
            Width           =   495
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   1305
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   2085
         Left            =   -74880
         TabIndex        =   18
         Top             =   1170
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3678
         _Version        =   65541
         Cols            =   3
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
         TabIndex        =   11
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
      Left            =   165
      TabIndex        =   12
      Top             =   3750
      Width           =   750
   End
End
Attribute VB_Name = "ABMServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Actualizar()
    
    On Error GoTo ErrorTrans
    sql = "UPDATE SERVICIOS "
    sql = sql & " SET SER_DESCRI = " & XS(txtdescri.Text)
    sql = sql & ",SER_PRECIO=" & XN(txtPrecio.Text)
    sql = sql & " WHERE SER_CODIGO = " & XN(txtcodigo.Text)
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
    If txtcodigo.Text = "" Then ' Si está VACIO
        txtcodigo.Text = "1"
        sql = "select max(SER_CODIGO) as maximo from SERVICIOS"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(rec.Fields!Maximo) Then txtcodigo = Val(Trim(rec.Fields!Maximo)) + 1
        rec.Close
    End If
    DBConn.BeginTrans
    sql = "INSERT INTO SERVICIOS (SER_CODIGO, SER_DESCRI,SER_PRECIO) "
    sql = sql & " VALUES ( " & XN(txtcodigo) & " ," & XS(txtdescri.Text) & ","
    sql = sql & XN(txtPrecio.Text) & ")"
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
    txtcodigo = ""
    txtdescri.Text = ""
    txtPrecio.Text = ""
    txtdescri.Enabled = True
    txtcodigo.SetFocus
    cmdGrabar.Enabled = False
    CmdBorrar.Enabled = False
    cmdNuevo.Enabled = True
    lblEstado.Caption = ""
End Function

Function ValidarIngreso()
Dim MensCampos As String
ValidarIngreso = True
For Each Control In ABMServicios.Controls ' revisar los controles del form
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
    If txtcodigo.Text <> "" Then
        On Error GoTo CLAVOSE
        resp = MsgBox("Seguro desea eliminar el Servicio: " & Trim(txtdescri.Text) & " ?", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        Screen.MousePointer = 11
        lblEstado.Caption = "Eliminando ..."
        
        sql = "DELETE FROM SERVICIOS WHERE SER_CODIGO = " & XN(txtcodigo.Text)
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
    sql = "SELECT * FROM SERVICIOS"
    If Trim(TxtDescriB) <> "" Then sql = sql & "WHERE SER_DESCRI LIKE '" & Trim(TxtDescriB) & "%'"
    sql = sql & " ORDER BY SER_DESCRI"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            GrdModulos.AddItem rec.Fields(0) & Chr(9) & rec.Fields(1) & Chr(9) & Valido_Importe(rec.Fields(2))
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
    Set ABMServicios = Nothing
End Sub

Private Sub CmdGrabar_Click()
    If txtdescri.Text = "" Then
        MsgBox "La descripción es requerida", vbExclamation, TIT_MSGBOX
        txtdescri.SetFocus
        Exit Sub
    End If
    If txtPrecio.Text = "" Then
        MsgBox "El Precio es requerido", vbExclamation, TIT_MSGBOX
        txtPrecio.SetFocus
        Exit Sub
    End If
    lblEstado.Caption = "Grabando..."
    If txtcodigo.Text <> "" Then
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
    GrdModulos.FormatString = "Código|Descripción|Costo"
    GrdModulos.ColWidth(0) = 900
    GrdModulos.ColWidth(1) = 4000
    GrdModulos.ColWidth(2) = 1000
    GrdModulos.Rows = 1
    GrdModulos.Cols = 3
    tabDatos.Tab = 0
    Centrar_pantalla Me
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
        'paso el item seleccionado al tab 'DATOS'
        txtcodigo.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0))
        txtdescri.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 1))
        txtPrecio.Text = Valido_Importe(Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 2)))
        CmdBorrar.Enabled = True
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
    If tabDatos.Tab = 0 And Me.Visible Then txtcodigo.SetFocus
    If tabDatos.Tab = 1 Then
       GrdModulos.Rows = 1
       GrdModulos.Refresh
       CmdBorrar.Enabled = False
       cmdGrabar.Enabled = False
       TxtDescriB.SetFocus
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto txtcodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If Me.txtcodigo.Text <> "" Then
        sql = "select * from SERVICIOS where SER_CODIGO =  " & XN(txtcodigo.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
           txtdescri.Text = Trim(rec!SER_DESCRI)
           txtPrecio.Text = Valido_Importe(rec!SER_PRECIO)
           CmdBorrar.Enabled = True
           cmdGrabar.Enabled = True
        Else
           MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
           txtcodigo.SetFocus
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

Private Sub txtPrecio_GotFocus()
    SelecTexto txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPrecio, KeyAscii)
End Sub

Private Sub txtPrecio_LostFocus()
  If txtPrecio.Text <> "" Then
     txtPrecio.Text = Valido_Importe(txtPrecio)
  End If
End Sub
