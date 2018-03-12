VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ABMFormaPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ABM de Forma de Pago"
   ClientHeight    =   4395
   ClientLeft      =   1245
   ClientTop       =   2355
   ClientWidth     =   6495
   ForeColor       =   &H00C0C0C0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4395
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMFormaPago.frx":0000
      Height          =   735
      Left            =   3780
      Picture         =   "ABMFormaPago.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3615
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMFormaPago.frx":0614
      Height          =   735
      Left            =   2895
      Picture         =   "ABMFormaPago.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3615
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMFormaPago.frx":0C28
      Height          =   735
      Left            =   5550
      Picture         =   "ABMFormaPago.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3615
      Width           =   870
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMFormaPago.frx":123C
      Height          =   735
      Left            =   4665
      Picture         =   "ABMFormaPago.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3615
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   3420
      Left            =   45
      TabIndex        =   10
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
      TabPicture(0)   =   "ABMFormaPago.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMFormaPago.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74865
         TabIndex        =   15
         Top             =   405
         Width           =   6135
         Begin VB.TextBox TxtDescriB 
            Height          =   330
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   6
            Top             =   225
            Width           =   4215
         End
         Begin VB.TextBox TxtCodigoB 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   16
            Top             =   225
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   360
            Left            =   5535
            MaskColor       =   &H000000FF&
            Picture         =   "ABMFormaPago.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   315
            Visible         =   0   'False
            Width           =   540
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   " Datos de Forma de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   345
         TabIndex        =   9
         Top             =   720
         Width           =   5625
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Left            =   1350
            MaxLength       =   40
            TabIndex        =   0
            Top             =   720
            Width           =   945
         End
         Begin VB.TextBox txtdescri 
            Height          =   330
            Left            =   1365
            MaxLength       =   40
            TabIndex        =   1
            Tag             =   "Descripción"
            Top             =   1440
            Width           =   4005
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
            Top             =   1470
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   2070
         Left            =   -74910
         TabIndex        =   8
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
      Left            =   150
      TabIndex        =   12
      Top             =   3645
      Width           =   750
   End
End
Attribute VB_Name = "ABMFormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Function Actualizar() As Integer
    If ValidarIngreso() = False Then
        Actualizar = False
        Exit Function
    Else
        On Error GoTo ErrorTrans
        sql = "UPDATE TIPO_FACTURA " & _
                 " SET tfa_descri = " & XS(txtdescri.Text) & _
               " WHERE tfa_codigo = " & XN(TxtCodigo)
        DBConn.BeginTrans
        DBConn.Execute sql, dbExecDirect
        DBConn.CommitTrans
        Actualizar = True
    End If
    Exit Function
    
ErrorTrans:
    Actualizar = False
    DBConn.RollbackTrans
    Beep
    MsgBox "No se pudo actualizar el Tipo de Factura.", 16, TIT_MSGBOX
    Exit Function
End Function

Function Insertar() As Integer

    If ValidarIngreso() = False Then
        Insertar = False
        Exit Function
    Else
        On Error GoTo ErrorTrans
        
        Set rec = New ADODB.Recordset
        If TxtCodigo.Text = "" Then ' Si está VACIO
            sql = "select max(tfa_codigo) as maximo from TIPO_FACTURA "
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = Val(Trim(rec.Fields!Maximo)) + 1
            rec.Close
        End If
        DBConn.BeginTrans
        sql = "INSERT INTO TIPO_FACTURA (tfa_codigo, tfa_descri) " & _
               "VALUES ( " & XN(Me.TxtCodigo.Text) & " ," & XS(txtdescri.Text) & ")"
        
        DBConn.Execute sql, dbExecDirect
        DBConn.CommitTrans
        Insertar = True
    End If
    Exit Function
    
ErrorTrans:
    Insertar = False
    DBConn.RollbackTrans
    Beep
    MsgBox "No se pudo grabar el Nuevo Tipo de Factura.", 16, TIT_MSGBOX
    Exit Function
End Function


Function LimpiarControles()
    TxtCodigo.Text = ""
    TxtCodigo.Enabled = True
    txtdescri.Text = ""
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
For Each Control In ABMFormaPago.Controls ' revisar los controles del form
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
            On Error GoTo HayError
            resp = MsgBox("Seguro desea eliminar la Forma de Pago: " & Trim(txtdescri.Text) & " ?", 36, "Eliminar:")
            If resp <> 6 Then Exit Sub
            
            Screen.MousePointer = 11
            lblEstado.Caption = "Eliminando ..."
            
            sql = "DELETE FROM FORMA_PAGO WHERE FPG_CODIGO=" & XN(TxtCodigo)
            DBConn.Execute sql
            lblEstado.Caption = ""
            Screen.MousePointer = 1
            LimpiarControles
       End If
       Exit Sub
HayError:
       Screen.MousePointer = vbNormal
       lblEstado.Caption = ""
       MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    Set rec = New ADODB.Recordset
    GrdModulos.Rows = 1

    Screen.MousePointer = 11
    Me.Refresh
    sql = "SELECT * FROM FORMA_PAGO"
    If Trim(TxtDescriB) <> "" Then sql = sql & " WHERE FPG_DESCRI LIKE '" & Trim(TxtDescriB) & "%'"
    sql = sql & " ORDER BY FPG_DESCRI"
    
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
        Set ABMFormaPago = Nothing
End Sub

Private Sub cmdGrabar_Click()
    If txtdescri.Text = "" Then
        MsgBox "Debe ingresar la descripción", vbExclamation, TIT_MSGBOX
        txtdescri.SetFocus
        Exit Sub
    End If
    On Error GoTo ErrorTrans
    
    lblEstado.Caption = "Grabando..."
    Screen.MousePointer = 11
    If TxtCodigo.Text = "" Then
        TxtCodigo.Text = "1"
        sql = "select max(FPG_CODIGO) as maximo FROM FORMA_PAGO "
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = Val(Trim(rec.Fields!Maximo)) + 1
        rec.Close
    
        DBConn.BeginTrans
        sql = "INSERT INTO FORMA_PAGO (FPG_CODIGO, FPG_DESCRI) " & _
              "VALUES ( " & XN(Me.TxtCodigo.Text) & " ," & XS(txtdescri.Text) & ")"
    
        DBConn.Execute sql, dbExecDirect
        DBConn.CommitTrans
    Else
        sql = "UPDATE FORMA_PAGO "
        sql = sql & " SET FPG_DESCRI = " & XS(txtdescri.Text)
        sql = sql & " WHERE FPG_CODIGO = " & XN(TxtCodigo)
        
        DBConn.BeginTrans
        DBConn.Execute sql, dbExecDirect
        DBConn.CommitTrans
    End If
        lblEstado.Caption = "Grabando..."
        Screen.MousePointer = 1
        LimpiarControles
    Exit Sub
    
ErrorTrans:
     lblEstado.Caption = ""
     Screen.MousePointer = 1
     DBConn.RollbackTrans
     MsgBox Err.Description, vbCritical, TIT_MSGBOX
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
    GrdModulos.FormatString = "Código|Descripción"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 4900
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    Centrar_pantalla Me
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
        'paso el item seleccionado al tab 'DATOS'
        GrdModulos.Col = 0
        TxtCodigo.Text = GrdModulos.Text
        GrdModulos.Col = 1
        txtdescri.Text = Trim(GrdModulos.Text)
        cmdBorrar.Enabled = True
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
    If tabDatos.Tab = 0 And Me.Visible Then txtdescri.SetFocus
    If tabDatos.Tab = 1 Then
       GrdModulos.Rows = 1
       GrdModulos.Refresh
       TxtDescriB.SetFocus
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
    If Me.TxtCodigo.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT * FROM FORMA_PAGO WHERE FPG_CODIGO=" & XN(TxtCodigo.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
           txtdescri.Text = Trim(rec!FPG_DESCRI)
           cmdBorrar.Enabled = True
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

Private Sub TxtDescriB_GotFocus()
    SelecTexto TxtDescriB
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
