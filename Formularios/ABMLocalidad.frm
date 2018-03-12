VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ABMLocalidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localidad"
   ClientHeight    =   4410
   ClientLeft      =   1770
   ClientTop       =   1905
   ClientWidth     =   6345
   ForeColor       =   &H00C0C0C0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4410
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMLocalidad.frx":0000
      Height          =   720
      Index           =   1
      Left            =   4530
      Picture         =   "ABMLocalidad.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3645
      Width           =   855
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMLocalidad.frx":0614
      Height          =   720
      Left            =   5400
      Picture         =   "ABMLocalidad.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3645
      Width           =   855
   End
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMLocalidad.frx":0C28
      Height          =   720
      Index           =   2
      Left            =   3660
      Picture         =   "ABMLocalidad.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3645
      Width           =   855
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   3510
      Left            =   60
      TabIndex        =   10
      Top             =   90
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
      TabPicture(0)   =   "ABMLocalidad.frx":123C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMLocalidad.frx":1258
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74850
         TabIndex        =   17
         Top             =   390
         Width           =   5940
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   345
            Left            =   5310
            MaskColor       =   &H000000FF&
            Picture         =   "ABMLocalidad.frx":1274
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   315
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   19
            Top             =   225
            Width           =   3975
         End
         Begin VB.TextBox TxtCodigoB 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   18
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
            TabIndex        =   22
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos de la Localidad"
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
         Begin VB.TextBox txtcodpostal 
            Height          =   300
            Left            =   4410
            MaxLength       =   6
            TabIndex        =   3
            Tag             =   "Identificación"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox txtcodigo 
            Height          =   300
            Left            =   1515
            MaxLength       =   3
            TabIndex        =   2
            Tag             =   "Identificación"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox txtdescri 
            Height          =   315
            Left            =   1515
            MaxLength       =   20
            TabIndex        =   4
            Tag             =   "Descripción"
            Top             =   2130
            Width           =   3690
         End
         Begin VB.ComboBox CboPais 
            Height          =   315
            Left            =   1515
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   465
            Width           =   2805
         End
         Begin VB.ComboBox CboProvincia 
            Height          =   315
            Left            =   1515
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1050
            Width           =   2805
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código Postal:"
            Height          =   195
            Left            =   3240
            TabIndex        =   23
            Top             =   1680
            Width           =   1020
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
            Left            =   570
            TabIndex        =   15
            Top             =   2160
            Width           =   885
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
            Left            =   915
            TabIndex        =   14
            Top             =   1680
            Width           =   540
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "País:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   13
            Top             =   540
            Width           =   375
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   750
            TabIndex        =   12
            Top             =   1065
            Width           =   705
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   2205
         Left            =   -74865
         TabIndex        =   21
         Top             =   1185
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   3889
         _Version        =   393216
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
         TabIndex        =   16
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Guardar"
      DisabledPicture =   "ABMLocalidad.frx":157E
      Height          =   720
      Index           =   0
      Left            =   2790
      Picture         =   "ABMLocalidad.frx":1888
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3645
      Width           =   855
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
      Left            =   90
      TabIndex        =   9
      Top             =   3750
      Width           =   750
   End
End
Attribute VB_Name = "ABMLocalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function LimpiarControles()
    tabDatos.Tab = 0
    CboProvincia.Clear
    txtcodigo.Text = ""
    txtdescri.Text = ""
    cmdBotonDatos(0).Enabled = True
    cmdBotonDatos(1).Enabled = False
    cmdBotonDatos(2).Enabled = True
    txtcodpostal.Text = ""
    CboPais.SetFocus
End Function

Private Sub Actualizar()
        On Error GoTo ErrorTrans
        DBConn.BeginTrans
        sql = "UPDATE localidad "
        sql = sql & " SET   loc_descri = " & XS(txtdescri.Text)
        sql = sql & " , LOC_CODPOS = " & XS(txtcodpostal.Text)
        sql = sql & " WHERE loc_codigo = " & XN(txtcodigo.Text)
        sql = sql & " AND   pro_codigo = " & CboProvincia.ItemData(CboProvincia.ListIndex)
        sql = sql & " AND   pai_codigo = " & CboPais.ItemData(CboPais.ListIndex)

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

Private Sub Borrar()
    On Error GoTo ErrorTrans
    DBConn.BeginTrans
    
    sql = " DELETE FROM LOCALIDAD "
    sql = sql & " WHERE loc_codigo = " & XN(txtcodigo.Text)
    sql = sql & " AND   pro_codigo = " & CboProvincia.ItemData(CboProvincia.ListIndex)
    sql = sql & " AND   pai_codigo = " & CboPais.ItemData(CboPais.ListIndex)
    DBConn.Execute sql, dbExecDirect
    DBConn.CommitTrans
    Exit Sub
    
ErrorTrans:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    Exit Sub
End Sub

Private Sub Insertar()
    sql = "select * from localidad "
    sql = sql & " where pai_codigo = " & CboPais.ItemData(CboPais.ListIndex)
    sql = sql & " and pro_codigo = " & CboProvincia.ItemData(CboProvincia.ListIndex)
    sql = sql & " and loc_descri = " & XS(Me.txtdescri)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        MsgBox "La localidad ya existe !", vbExclamation, TIT_MSGBOX
        Rec1.Close
        Exit Sub
    End If
    Rec1.Close
    
    sql = " SELECT max(loc_codigo) as maximo FROM localidad"
    sql = sql & " where pro_codigo = " & CboProvincia.ItemData(CboProvincia.ListIndex)
    sql = sql & " AND pai_codigo = " & CboPais.ItemData(CboPais.ListIndex)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If IsNull(Rec1!Maximo) Then
         Maximo = 0
         txtcodigo.Text = Maximo + 1
    Else
         txtcodigo.Text = Rec1!Maximo + 1
    End If
    Rec1.Close
    
    On Error GoTo ErrorTrans
    
    DBConn.BeginTrans
    sql = "INSERT INTO localidad (loc_codigo, loc_descri,pai_codigo,pro_codigo,loc_codpos) "
    sql = sql & "VALUES ( " & XN(txtcodigo.Text) & ", " & XS(txtdescri.Text)
    sql = sql & ", " & CboPais.ItemData(CboPais.ListIndex) & ", "
    sql = sql & CboProvincia.ItemData(CboProvincia.ListIndex) & " ,"
    sql = sql & XS(txtcodpostal.Text) & ")"
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

Private Sub CboPais_LostFocus()
       
    'Set rec = New ADODB.Recordset
      'cargo el combo de Provincias
      CboProvincia.Clear
      sql = "SELECT PRO_CODIGO,PRO_DESCRI"
      sql = sql & " FROM PROVINCIA "
      sql = sql & " WHERE PAI_CODIGO=" & CboPais.ItemData(CboPais.ListIndex)
      sql = sql & " ORDER BY PRO_DESCRI"
      
      rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
      If (rec.BOF And rec.EOF) = 0 Then
         Do While rec.EOF = False
            CboProvincia.AddItem Trim(rec!PRO_DESCRI)
            CboProvincia.ItemData(CboProvincia.NewIndex) = rec!PRO_CODIGO
            rec.MoveNext
         Loop
         CboProvincia.ListIndex = 0
      Else
         MsgBox "No hay cargado Provincia para ese País.", vbOKOnly + vbCritical, TIT_MSGBOX
      End If
      rec.Close
End Sub

Private Sub cmdBotonDatos_Click(Index As Integer)

    If tabDatos.Tab <> 0 Then
        Exit Sub
    End If
    Select Case Index
         Case 0 ' aceptar
            If txtdescri.Text = "" Then
             MsgBox "Debe ingresar la descipción", vbExclamation, TIT_MSGBOX
             txtdescri.SetFocus
             Exit Sub
            End If
            lblEstado.Caption = "Grabando..."
            
            If txtcodigo = "" Then
                Insertar
            Else
                Actualizar
            End If
            lblEstado.Caption = ""
            LimpiarControles
        
        Case 1 ' eliminar
           If txtcodigo.Text <> "" Then
                resp = MsgBox("Seguro desea eliminar la Localidad: " & Trim(txtdescri.Text) & " ?", 36, "Eliminar:")
                If resp <> 6 Then Exit Sub
                
                Screen.MousePointer = 11
                lblEstado.Caption = "Eliminando ..."
                Borrar
                LimpiarControles
                lblEstado.Caption = ""
                Screen.MousePointer = 1
                LimpiarControles
          End If
        Case 2 ' cancelar
            LimpiarControles
    End Select
End Sub

Private Sub CmdBuscAprox_Click()
   
    GrdModulos.Rows = 1
    'Set rec = New ADODB.Recordset
           
If Not (CboPais.ListIndex = -1) And _
   Not (CboProvincia.ListIndex = -1) Then
    
    MousePointer = vbHourglass
    
    sql = "SELECT LOC_CODIGO,LOC_DESCRI"
    sql = sql & " FROM  LOCALIDAD"
    sql = sql & " WHERE PRO_CODIGO=" & CboProvincia.ItemData(CboProvincia.ListIndex)
    sql = sql & " AND PAI_CODIGO=" & CboPais.ItemData(CboPais.ListIndex)
    If TxtDescriB <> "" Then
     sql = sql & " AND LOC_DESCRI LIKE '" & TxtDescriB.Text & "%' ORDER BY LOC_DESCRI"
    End If

    lblEstado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Do While Not rec.EOF
           GrdModulos.AddItem rec.Fields(0) & Chr(9) & rec.Fields(1)
           rec.MoveNext
        Loop
        If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        LimpiarControles
    End If
    rec.Close
    MousePointer = vbDefault
    lblEstado.Caption = ""
Else
    lblEstado.Caption = ""
    MsgBox "Seleccione la opción que falta.", vbOKOnly + vbCritical, TIT_MSGBOX
    tabDatos.Tab = 0
    CboPais.SetFocus
    Exit Sub
End If
End Sub

Private Sub CmdCerrar_Click()
        Unload Me
        Set frmLocalidad = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdCerrar_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset

    GrdModulos.FormatString = "Código|Descripción"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 4600
    GrdModulos.row = 0
    
    KeyPreview = True
    Call Centrar_pantalla(Me)
    
    txtcodigo.MaxLength = 3
    txtdescri.MaxLength = 20
    
    lblEstado.Visible = True
    lblEstado.Caption = ""
    
      
    tabDatos.Tab = 0
    'cargo el combo de Pais
    CboPais.Clear
    
    sql = "SELECT PAI_CODIGO,PAI_DESCRI"
    sql = sql & " FROM PAIS ORDER BY PAI_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          CboPais.AddItem Trim(rec!PAI_DESCRI)
          CboPais.ItemData(CboPais.NewIndex) = rec!PAI_CODIGO
          rec.MoveNext
       Loop
       CboPais.ListIndex = 0
    Else
       MsgBox "No hay cargado País.", vbOKOnly + vbCritical, TIT_MSGBOX
    End If
    rec.Close
    Screen.MousePointer = 1

End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
        GrdModulos.Col = 0
        txtcodigo = GrdModulos.Text
        GrdModulos.Col = 1
        txtdescri.Text = Trim(GrdModulos.Text)
        If txtdescri.Enabled Then TxtCodigo_LostFocus
        cmdBotonDatos(0).Enabled = True
        cmdBotonDatos(1).Enabled = True
        tabDatos.Tab = 0
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 0 And Me.Visible Then txtcodigo.SetFocus
    If tabDatos.Tab = 1 Then
     TxtDescriB.SetFocus
     TxtDescriB.Text = ""
     GrdModulos.Rows = 1
     cmdBotonDatos(0).Enabled = False
     cmdBotonDatos(1).Enabled = False
    End If
End Sub
Private Sub TxtCodigo_LostFocus()
    If txtcodigo.Text <> "" Then
        sql = "SELECT LOC_DESCRI,LOC_CODPOS"
        sql = sql & " FROM LOCALIDAD"
        sql = sql & " WHERE PRO_CODIGO=" & CboProvincia.ItemData(CboProvincia.ListIndex)
        sql = sql & " AND PAI_CODIGO=" & CboPais.ItemData(CboPais.ListIndex)
        sql = sql & " AND LOC_CODIGO=" & XN(txtcodigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
         txtdescri.Text = rec!loc_descri
         txtcodpostal.Text = IIf(IsNull(rec!LOC_CODPOS), "", rec!LOC_CODPOS)
         cmdBotonDatos(1).Enabled = True
        Else
         MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
         txtcodigo.Text = ""
         txtcodigo.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtcodpostal_GotFocus()
    SelecTexto txtcodpostal
End Sub

Private Sub txtcodpostal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_Change()
    If txtdescri.Text = "" Then
      cmdBotonDatos(0).Enabled = False
    Else
      cmdBotonDatos(0).Enabled = True
    End If
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtdescri_GotFocus()
   SelecTexto txtdescri
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
   KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

