VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ABMPais 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "País"
   ClientHeight    =   4035
   ClientLeft      =   1680
   ClientTop       =   2250
   ClientWidth     =   6540
   ForeColor       =   &H00C0C0C0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMPais.frx":0000
      Height          =   735
      Index           =   1
      Left            =   4545
      Picture         =   "ABMPais.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3255
      Width           =   930
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMPais.frx":0614
      Height          =   735
      Left            =   5505
      Picture         =   "ABMPais.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3255
      Width           =   930
   End
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMPais.frx":0C28
      Height          =   735
      Index           =   0
      Left            =   2655
      Picture         =   "ABMPais.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3255
      Width           =   930
   End
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMPais.frx":123C
      Height          =   735
      Index           =   2
      Left            =   3600
      Picture         =   "ABMPais.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3255
      Width           =   930
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   3045
      Left            =   105
      TabIndex        =   10
      Top             =   135
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   5371
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "ABMPais.frx":1850
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMPais.frx":186C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   270
         TabIndex        =   14
         Top             =   360
         Width           =   5775
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   345
            Left            =   5130
            MaskColor       =   &H000000FF&
            Picture         =   "ABMPais.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   315
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   6
            Top             =   225
            Width           =   3900
         End
         Begin VB.TextBox TxtCodigoB 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   15
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos de la País"
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
         Left            =   -74580
         TabIndex        =   9
         Top             =   600
         Width           =   5430
         Begin VB.TextBox txtdescri 
            Height          =   315
            Left            =   1110
            MaxLength       =   30
            TabIndex        =   1
            Tag             =   "Descripción"
            Top             =   1125
            Width           =   3465
         End
         Begin VB.TextBox txtcodigo 
            Height          =   300
            Left            =   1110
            MaxLength       =   3
            TabIndex        =   0
            Top             =   555
            Width           =   795
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
            Left            =   525
            TabIndex        =   13
            Top             =   585
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
            Left            =   180
            TabIndex        =   12
            Top             =   1155
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   1755
         Left            =   255
         TabIndex        =   8
         Top             =   1125
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   3096
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
      Left            =   135
      TabIndex        =   18
      Top             =   3435
      Width           =   750
   End
End
Attribute VB_Name = "ABMPais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodPais As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Function LimpiarControles()
    TxtCodigo.Text = ""
    txtDescri.Text = ""
    txtDescri.Enabled = True
    TxtCodigo.SetFocus
    cmdBotonDatos(0).Enabled = True
    cmdBotonDatos(1).Enabled = False
    cmdBotonDatos(2).Enabled = True
End Function


Private Sub cmdBotonDatos_Click(Index As Integer)
    
    If tabDatos.Tab <> 0 Then
        Exit Sub
    End If
    
    Select Case Index
         Case 0 ' GRABAR
            On Error GoTo ErrorTrans
            If txtDescri.Text = "" Then
             MsgBox "Debe ingresar la descripción", vbExclamation, TIT_MSGBOX
             txtDescri.SetFocus
             Exit Sub
            End If
            lblEstado.Caption = "Grabando..."
            If TxtCodigo.Text = "" Then
                TxtCodigo = "1"
                sql = "SELECT max(PAI_CODIGO) as maximo FROM PAIS "
                rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = Val(Trim(rec.Fields!Maximo)) + 1
                rec.Close
                
                DBConn.BeginTrans
                sql = "INSERT INTO PAIS (PAI_CODIGO,PAI_DESCRI) " & _
                       "VALUES ( " & XN(TxtCodigo.Text) & " ," & XS(txtDescri.Text) & ")"
                DBConn.Execute sql, dbExecDirect
                DBConn.CommitTrans
            Else
                DBConn.BeginTrans
                sql = " UPDATE PAIS "
                sql = sql & " SET   PAI_DESCRI = " & XS(txtDescri.Text)
                sql = sql & " WHERE PAI_CODIGO = " & XN(TxtCodigo.Text)
              
                DBConn.Execute sql, dbExecDirect
                DBConn.CommitTrans
            End If
            lblEstado.Caption = ""
            LimpiarControles
        Case 1 ' ELIMINAR
           If Me.TxtCodigo.Text <> "" Then
                On Error GoTo ErrorTrans
                sql = "SELECT * from PROVINCIA where PAI_CODIGO = " & Trim(TxtCodigo.Text)
                rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If rec.RecordCount > 0 Then
                    MsgBox "No se puede eliminar este PAIS porque tiene PROVINCIAS asociadas!", vbExclamation, "Mensaje:"
                    rec.Close
                    Exit Sub
                End If
                rec.Close
                
                resp = MsgBox("Seguro desea eliminar el Pais: " & Trim(txtDescri.Text) & " ?", 36, "Eliminar:")
                If resp <> 6 Then Exit Sub
                
                Screen.MousePointer = 11
                lblEstado.Caption = "Eliminando ..."
                
                DBConn.Execute "DELETE FROM PAIS WHERE PAI_CODIGO = " & XN(TxtCodigo.Text)
                If txtDescri.Enabled Then
                    LimpiarControles
                End If
                lblEstado.Caption = ""
                Screen.MousePointer = 1
          End If
        Case 2 ' cancelar
            LimpiarControles
            GrdModulos.Rows = 1
    End Select
    Exit Sub

ErrorTrans:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub CmdBuscAprox_Click()
    
    GrdModulos.Rows = 1

    sql = "SELECT PAI_CODIGO,PAI_DESCRI FROM PAIS "
    If Trim(TxtDescriB) <> "" Then sql = sql & "WHERE PAI_DESCRI LIKE '" & Trim(Me.TxtDescriB) & "%'"
    sql = sql & " ORDER BY PAI_DESCRI"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            GrdModulos.AddItem rec.Fields(0) & Chr(9) & rec.Fields(1)
            rec.MoveNext
        Loop
        
        If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        MsgBox "No se encontraron items con esta descripcion !", vbExclamation, TIT_MSGBOX
    End If
    rec.Close

End Sub

Private Sub CmdCerrar_Click()
    Unload Me
    Set frmPais = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then CmdCerrar_Click
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    Call Centrar_pantalla(Me)
    Set rec = New ADODB.Recordset
    
    lblEstado.Caption = ""

    GrdModulos.FormatString = "Código|Descripción"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 4450
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    Screen.MousePointer = 1
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
        GrdModulos.Col = 0
        TxtCodigo = GrdModulos.Text
        GrdModulos.Col = 1
        txtDescri.Text = Trim(GrdModulos.Text)
        If txtDescri.Enabled Then txtDescri.SetFocus
        cmdBotonDatos(1).Enabled = True
        cmdBotonDatos(0).Enabled = True
        tabDatos.Tab = 0
    End If
End Sub
Private Sub GrdModulos_GotFocus()
    GrdModulos.Col = 0
    GrdModulos.ColSel = GrdModulos.Cols - 1
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
        TxtDescriB.Text = ""
        TxtDescriB.SetFocus
        GrdModulos.Rows = 1
        cmdBotonDatos(0).Enabled = False
        cmdBotonDatos(1).Enabled = False
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
        sql = "SELECT * FROM PAIS"
        sql = sql & " WHERE PAI_CODIGO=" & XN(TxtCodigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
         txtDescri.Text = rec!PAI_DESCRI
         cmdBotonDatos(1).Enabled = True
        Else
         MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
         TxtCodigo.Text = ""
         TxtCodigo.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtDescri_Change()
    If Trim(txtDescri) = "" Then
        cmdBotonDatos(0).Enabled = False
    Else
        cmdBotonDatos(0).Enabled = True
    End If
End Sub

Private Sub txtdescri_GotFocus()
   SelecTexto txtDescri
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
   KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Function ValidarIngreso()
Dim MensCampos As String
ValidarIngreso = True
For Each Control In frmPais.Controls ' revisar los controles del form
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
    txtDescri.SetFocus
End If
End Function
