VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ABMProvincia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provincia"
   ClientHeight    =   4395
   ClientLeft      =   975
   ClientTop       =   1500
   ClientWidth     =   5955
   ForeColor       =   &H00C0C0C0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4395
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMProvincia.frx":0000
      Height          =   735
      Index           =   2
      Left            =   3225
      Picture         =   "ABMProvincia.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3615
      Width           =   870
   End
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMProvincia.frx":0614
      Height          =   735
      Index           =   0
      Left            =   2355
      Picture         =   "ABMProvincia.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3615
      Width           =   870
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMProvincia.frx":0C28
      Height          =   735
      Left            =   4995
      Picture         =   "ABMProvincia.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3615
      Width           =   870
   End
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMProvincia.frx":123C
      Height          =   735
      Index           =   1
      Left            =   4110
      Picture         =   "ABMProvincia.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3615
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   3435
      Left            =   105
      TabIndex        =   11
      Top             =   135
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   6059
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
      TabPicture(0)   =   "ABMProvincia.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMProvincia.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74775
         TabIndex        =   16
         Top             =   480
         Width           =   5235
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   330
            Left            =   4650
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProvincia.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   435
         End
         Begin VB.TextBox TxtCodigoB 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   17
            Top             =   225
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   300
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   7
            Top             =   225
            Width           =   3390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos de la Provincia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   405
         TabIndex        =   10
         Top             =   690
         Width           =   4740
         Begin VB.ComboBox CboPais 
            Height          =   315
            Left            =   1440
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   465
            Width           =   2130
         End
         Begin VB.TextBox txtdescri 
            Height          =   315
            Left            =   1440
            MaxLength       =   30
            TabIndex        =   2
            Tag             =   "Descripción"
            Top             =   1635
            Width           =   2865
         End
         Begin VB.TextBox txtcodigo 
            Height          =   300
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   1
            Top             =   1065
            Width           =   795
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "País:"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   780
            TabIndex        =   15
            Top             =   525
            Width           =   420
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
            Left            =   660
            TabIndex        =   14
            Top             =   1080
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
            Left            =   315
            TabIndex        =   13
            Top             =   1695
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   1875
         Left            =   -74835
         TabIndex        =   9
         Top             =   1335
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   3307
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
      Left            =   165
      TabIndex        =   19
      Top             =   3675
      Width           =   750
   End
End
Attribute VB_Name = "ABMProvincia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBuscAprox_Click()

    GrdModulos.Rows = 1
    MousePointer = vbHourglass
    
    sql = "SELECT * FROM PROVINCIA"
    sql = sql & " WHERE PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
    sql = sql & " And PRO_DESCRI"
    sql = sql & " LIKE '" & Me.TxtDescriB.Text & "%' ORDER BY PRO_DESCRI"
    
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
        Me.TxtDescriB.SetFocus
    End If
    rec.Close
    MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Function Agregar_Combo()
    OtroForm = True
    CrtlDescri = txtdescri.Text
    LimpiarControles
    Unload Me
    Set frmProvincia = Nothing
End Function

Function LimpiarControles()
    TxtCodigo.Text = ""
    txtdescri.Text = ""
    txtdescri.Enabled = True
    cboPais.SetFocus
    cmdBotonDatos(0).Enabled = True
    cmdBotonDatos(1).Enabled = False
    cmdBotonDatos(2).Enabled = True
End Function

Private Sub CboPais_LostFocus()
    TxtCodigo.Text = ""
End Sub

Private Sub cmdBotonDatos_Click(Index As Integer)
    
    If tabDatos.Tab <> 0 Then
        Exit Sub
    End If
    On Error GoTo ErrorTrans
    Select Case Index
         Case 0 ' Grabar
            If txtdescri.Text = "" Then
                MsgBox "Debe ingresar la Descripción", vbExclamation, TIT_MSGBOX
                txtdescri.SetFocus
                Exit Sub
            End If
            
            lblEstado.Caption = "Grabando..."
            If TxtCodigo.Text = "" Then
                TxtCodigo = "1"
                sql = "SELECT max(PRO_CODIGO) as maximo "
                sql = sql & " FROM PROVINCIA "
                sql = sql & " WHERE PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
                
                rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                
                If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = Val(Trim(rec.Fields!Maximo)) + 1
                rec.Close
                
                DBConn.BeginTrans
                sql = "INSERT INTO Provincia (pai_codigo,pro_codigo,pro_descri) "
                sql = sql & " VALUES ( " & cboPais.ItemData(cboPais.ListIndex) & " ,"
                sql = sql & Trim(TxtCodigo.Text) & " ," & XS(txtdescri.Text) & ")"
                DBConn.Execute sql, dbExecDirect
                DBConn.CommitTrans
            Else
                DBConn.BeginTrans
                sql = " UPDATE Provincia "
                sql = sql & " SET pro_descri = " & XS(txtdescri.Text)
                sql = sql & " WHERE Pro_codigo = " & XN(TxtCodigo.Text)
                sql = sql & " AND pai_codigo = " & cboPais.ItemData(cboPais.ListIndex)
                DBConn.Execute sql, dbExecDirect
                DBConn.CommitTrans
            End If
            lblEstado.Caption = ""
             LimpiarControles
             
        Case 1 ' eliminar
            If TxtCodigo.Text <> "" Then
                sql = "select * from localidad "
                sql = sql & " where PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
                sql = sql & " and PRO_CODIGO = " & Trim(TxtCodigo.Text)
                rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If rec.RecordCount > 0 Then
                    MsgBox "No se puede eliminar esta PROVINCIA porque tiene LOCALIDAD asociadas!", vbExclamation, "Mensaje:"
                    rec.Close
                    Exit Sub
                End If
                rec.Close
                
                resp = MsgBox("Seguro desea eliminar la Provincia: " & Trim(txtdescri.Text) & " ?", 36, "Eliminar:")
                If resp <> 6 Then Exit Sub
                
                Screen.MousePointer = 11
                lblEstado.Caption = "Eliminando ..."
                
                DBConn.BeginTrans
                sql = "DELETE FROM provincia "
                sql = sql & " WHERE pai_codigo = " & cboPais.ItemData(cboPais.ListIndex)
                sql = sql & " AND pro_codigo = " & XN(TxtCodigo.Text)
                DBConn.Execute sql
                DBConn.CommitTrans
                
                LimpiarControles
                lblEstado.Caption = ""
                Screen.MousePointer = 1
            End If
        Case 2 ' cancelar
            LimpiarControles
    End Select
    Exit Sub
    
ErrorTrans:
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    LimpiarControles
End Sub

Private Sub CmdCerrar_Click()
        Unload Me
        Set frmProvincia = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then CmdCerrar_Click
    If KeyAscii = vbKeyReturn Then 'avanza de campo
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    GrdModulos.FormatString = "Código|Descripción"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 4030
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    
    Call Centrar_pantalla(Me)
    Set rec = New ADODB.Recordset
    
    TxtCodigo.MaxLength = 3
    txtdescri.MaxLength = 15
    lblEstado.Visible = True
    lblEstado.Caption = ""
    
    cboPais.Clear    'cargo el combo de Pais
    sql = "SELECT * FROM PAIS"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboPais.AddItem Trim(rec!PAI_DESCRI)
          cboPais.ItemData(cboPais.NewIndex) = rec!PAI_CODIGO
          rec.MoveNext
       Loop
       cboPais.ListIndex = 0
    Else
        MsgBox "No hay País cargados", vbOKOnly + vbCritical, TIT_MSGBOX
        Exit Sub
    End If
    rec.Close
    Screen.MousePointer = 1
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
        GrdModulos.Col = 0
        TxtCodigo = GrdModulos.Text
        GrdModulos.Col = 1
        txtdescri.Text = Trim(GrdModulos.Text)
        iEdita = True
        If txtdescri.Enabled Then txtdescri.SetFocus
        cmdBotonDatos(1).Enabled = True
        tabDatos.Tab = 0
    End If

End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 0 And Me.Visible Then TxtCodigo.SetFocus
    If tabDatos.Tab = 1 Then
        TxtDescriB.Text = ""
        TxtDescriB.SetFocus
        GrdModulos.Rows = 1
        cmdBotonDatos(0).Enabled = False
        cmdBotonDatos(1).Enabled = False
    Else
        cmdBotonDatos(0).Enabled = True
        cmdBotonDatos(1).Enabled = True
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCodigo
End Sub

Private Sub TxtCodigo_LostFocus()
    If TxtCodigo.Text <> "" Then
     sql = "SELECT * FROM PROVINCIA"
     sql = sql & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
     sql = sql & " AND PRO_CODIGO=" & XN(TxtCodigo.Text)
     
     rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
     If rec.EOF = False Then
      txtdescri.Text = rec!PRO_DESCRI
     Else
      MsgBox "El Código no Existe", vbExclamation, TIT_MSGBOX
      TxtCodigo.SetFocus
     End If
     rec.Close
    End If
End Sub

Private Sub txtDescri_Change()
    If Trim(txtdescri) = "" Then
        cmdBotonDatos(0).Enabled = False
    Else
        cmdBotonDatos(0).Enabled = True
    End If
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


