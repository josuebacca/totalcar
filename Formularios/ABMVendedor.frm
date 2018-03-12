VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form ABMVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM de Vendedores"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMVendedor.frx":0000
      Height          =   720
      Left            =   4740
      Picture         =   "ABMVendedor.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5130
      Width           =   855
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMVendedor.frx":0614
      Height          =   720
      Left            =   5610
      Picture         =   "ABMVendedor.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5130
      Width           =   855
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMVendedor.frx":0C28
      Height          =   720
      Left            =   3000
      Picture         =   "ABMVendedor.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5130
      Width           =   855
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMVendedor.frx":123C
      Height          =   720
      Left            =   3870
      Picture         =   "ABMVendedor.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5130
      Width           =   855
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   4995
      Left            =   75
      TabIndex        =   16
      Top             =   75
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   8811
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
      TabPicture(0)   =   "ABMVendedor.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMVendedor.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraDatos 
         Caption         =   " Datos del Vendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4245
         Left            =   150
         TabIndex        =   20
         Top             =   555
         Width           =   6105
         Begin VB.CommandButton cmdNuevoPais 
            Height          =   315
            Left            =   4050
            MaskColor       =   &H000000FF&
            Picture         =   "ABMVendedor.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Agregar País"
            Top             =   1635
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaLocalidad 
            Height          =   315
            Left            =   4545
            MaskColor       =   &H000000FF&
            Picture         =   "ABMVendedor.frx":1C12
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Agregar Localidad"
            Top             =   2385
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaProvincia 
            Height          =   315
            Left            =   4050
            MaskColor       =   &H000000FF&
            Picture         =   "ABMVendedor.frx":1F9C
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Agregar Provincia"
            Top             =   2010
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboProv 
            Height          =   315
            Left            =   1125
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   2000
            Width           =   2880
         End
         Begin VB.ComboBox cboPais 
            Height          =   315
            Left            =   1125
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1620
            Width           =   2880
         End
         Begin VB.ComboBox cboLocal 
            Height          =   315
            Left            =   1125
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2380
            Width           =   3375
         End
         Begin VB.TextBox txtmail 
            Height          =   300
            Left            =   1125
            LinkTimeout     =   0
            MaxLength       =   40
            TabIndex        =   8
            Top             =   3810
            Width           =   4455
         End
         Begin VB.TextBox txtfax 
            Height          =   300
            Left            =   3735
            MaxLength       =   25
            TabIndex        =   7
            Top             =   3435
            Width           =   1815
         End
         Begin VB.TextBox txtdomici 
            Height          =   300
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   5
            Top             =   2760
            Width           =   4335
         End
         Begin VB.TextBox txtnombre 
            Height          =   300
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "Descripción"
            Top             =   885
            Width           =   4335
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Left            =   1125
            MaxLength       =   40
            TabIndex        =   0
            Top             =   465
            Width           =   975
         End
         Begin VB.TextBox txttelefono 
            Height          =   300
            Left            =   1125
            MaxLength       =   25
            TabIndex        =   6
            Text            =   " "
            Top             =   3435
            Width           =   1575
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Vias de Comunicación"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   300
            TabIndex        =   32
            Top             =   3180
            Width           =   1575
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   300
            TabIndex        =   33
            Top             =   1320
            Width           =   630
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   360
            X2              =   5580
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            X1              =   360
            X2              =   5580
            Y1              =   3330
            Y2              =   3330
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "País:"
            Height          =   195
            Left            =   555
            TabIndex        =   31
            Top             =   1650
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail:"
            Height          =   195
            Left            =   390
            TabIndex        =   30
            Top             =   3855
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Left            =   3225
            TabIndex        =   29
            Top             =   3480
            Width           =   300
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   195
            TabIndex        =   28
            Top             =   3465
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   195
            TabIndex        =   27
            Top             =   2040
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   195
            TabIndex        =   26
            Top             =   2430
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   255
            TabIndex        =   25
            Top             =   2805
            Width           =   675
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   22
            Top             =   930
            Width           =   600
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
            Left            =   450
            TabIndex        =   21
            Top             =   510
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74865
         TabIndex        =   17
         Top             =   600
         Width           =   6135
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   360
            Left            =   5580
            MaskColor       =   &H000000FF&
            Picture         =   "ABMVendedor.frx":2326
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Buscar"
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   420
         End
         Begin VB.TextBox TxtNombreb 
            Height          =   330
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   13
            Top             =   225
            Width           =   4215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1665
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   270
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   3060
         Left            =   -74895
         TabIndex        =   15
         Top             =   1530
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   5398
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
         TabIndex        =   23
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
      Left            =   180
      TabIndex        =   24
      Top             =   5265
      Width           =   750
   End
End
Attribute VB_Name = "ABMVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As ADODB.Recordset
Dim sql As String
Dim resp As Integer
Dim Consulta As Boolean
Dim Pais As String
Dim Provincia As String

Private Sub cboPais_Click()
    cboProv.Enabled = True
End Sub

Private Sub CboPais_LostFocus()
    
    If Consulta = True And cboPais.Text = Pais Then
        Exit Sub
    Else
        Pais = ""
    End If
    'cargo combo PROVINCIA
    sql = "SELECT * FROM PROVINCIA"
    sql = sql & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
    sql = sql & " ORDER BY PRO_DESCRI "
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboProv.AddItem rec!PRO_DESCRI
            cboProv.ItemData(cboProv.NewIndex) = rec!PRO_CODIGO
            rec.MoveNext
        Loop
        cboProv.ListIndex = 0
        cboProv.Enabled = True
    End If
    rec.Close
    cboProv.SetFocus
End Sub

Private Sub cboProv_Click()
    cboLocal.Enabled = True
End Sub

Private Sub cboProv_LostFocus()
    If cboProv.ListIndex <> -1 Then
        If Consulta = True And cboProv.Text = Provincia Then
            Exit Sub
        Else
            Provincia = ""
        End If
        'cargo combo Localidad
        sql = "SELECT * FROM LOCALIDAD"
        sql = sql & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
        sql = sql & " AND PRO_CODIGO=" & cboProv.ItemData(cboProv.ListIndex)
        sql = sql & " ORDER BY LOC_DESCRI "
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                cboLocal.AddItem rec!LOC_DESCRI
                cboLocal.ItemData(cboLocal.NewIndex) = rec!LOC_CODIGO
                rec.MoveNext
            Loop
            cboLocal.ListIndex = 0
            cboLocal.Enabled = True
        End If
        rec.Close
        cboLocal.SetFocus
    End If
End Sub

Private Sub CmdBorrar_Click()
On Error GoTo CLAVOSE
    If Trim(TxtCodigo) <> "" Then
        If Trim(TxtCodigo) = 1 Then
            MsgBox "No se puede eliminar el Vendedor: " & Trim(txtnombre) & " "
            Exit Sub
        End If
        resp = MsgBox("Seguro desea eliminar el Vendedor: " & Trim(txtnombre) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = 11
        lblEstado.Caption = "Eliminando ..."
        
        DBConn.Execute "DELETE FROM VENDEDOR WHERE VEN_CODIGO = " & XN(TxtCodigo)
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
    MousePointer = vbHourglass
    
    sql = "SELECT * FROM VENDEDOR"
    sql = sql & " WHERE VEN_NOMBRE"
    sql = sql & " LIKE '" & Me.TxtNombreb.Text & "%' ORDER BY VEN_NOMBRE"
        
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
        Me.TxtNombreb.SetFocus
    End If
    rec.Close
    MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub CmdGrabar_Click()
    
    If ValidarVendedor = False Then Exit Sub
    
    On Error GoTo HayError
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    DBConn.BeginTrans
    If TxtCodigo.Text <> "" Then
        sql = "UPDATE VENDEDOR "
        sql = sql & " SET VEN_NOMBRE= " & XS(txtnombre)
        sql = sql & " , VEN_DOMICI= " & XS(txtDomici)
        sql = sql & " , PAI_CODIGO= " & cboPais.ItemData(cboPais.ListIndex)
        sql = sql & " , PRO_CODIGO= " & cboProv.ItemData(cboProv.ListIndex)
        sql = sql & " , LOC_CODIGO= " & cboLocal.ItemData(cboLocal.ListIndex)
        sql = sql & " , VEN_TELEFONO= " & XS(txtTelefono)
        sql = sql & " , VEN_FAX= " & XS(txtFax)
        sql = sql & " , VEN_MAIL= " & XS(txtMail)
        sql = sql & " WHERE VEN_CODIGO = " & XN(TxtCodigo)
        DBConn.Execute sql
        
    Else
        TxtCodigo = "1"
        sql = "SELECT MAX(VEN_CODIGO) as maximo FROM VENDEDOR"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = XN(rec.Fields!Maximo) + 1
        rec.Close
        
        sql = "INSERT INTO VENDEDOR(VEN_CODIGO,VEN_NOMBRE,VEN_DOMICI,"
        sql = sql & "PAI_CODIGO,PRO_CODIGO,LOC_CODIGO,VEN_TELEFONO,VEN_FAX,VEN_MAIL) "
        sql = sql & " VALUES ("
        sql = sql & XN(TxtCodigo) & ","
        sql = sql & XS(txtnombre) & ","
        sql = sql & XS(txtDomici) & ","
        sql = sql & cboPais.ItemData(cboPais.ListIndex) & ","
        sql = sql & cboProv.ItemData(cboProv.ListIndex) & ","
        sql = sql & cboLocal.ItemData(cboLocal.ListIndex) & ","
        sql = sql & XS(txtTelefono) & ","
        sql = sql & XS(txtFax) & ","
        sql = sql & XS(txtMail) & ")"
        DBConn.Execute sql
    End If
    Screen.MousePointer = vbNormal
    DBConn.CommitTrans
    CmdNuevo_Click
    Exit Sub
    
HayError:
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX

End Sub

Private Sub cmdNuevaLocalidad_Click()
    ABMLocalidad.Show vbModal
    cboLocal.Clear
    cboProv_LostFocus
End Sub

Private Sub cmdNuevaProvincia_Click()
    ABMProvincia.Show vbModal
    cboProv.Clear
    Provincia = ""
    CboPais_LostFocus
End Sub

Private Sub CmdNuevo_Click()
    tabDatos.Tab = 0
    TxtCodigo.Text = ""
    txtnombre.Text = ""
    txtDomici.Text = ""
    txtTelefono.Text = ""
    txtFax.Text = ""
    txtMail.Text = ""
    lblEstado.Caption = ""
    'para la consulta
    Consulta = False 'no consulta true consulta
    Pais = ""
    Provincia = ""
    
    GrdModulos.Rows = 1
    cboProv.Clear
    cboLocal.Clear
    cboProv.Enabled = False
    cboLocal.Enabled = False
    cboPais.ListIndex = 0
    TxtCodigo.SetFocus
End Sub

Private Sub cmdNuevoPais_Click()
    ABMPais.Show vbModal
    cboPais.Clear
    Pais = ""
    LlenarComboPais
    cboPais.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set ABMVendedor = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub LlenarComboPais()
    sql = "SELECT * FROM PAIS ORDER BY PAI_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboPais.AddItem rec!PAI_DESCRI
            cboPais.ItemData(cboPais.NewIndex) = rec!PAI_CODIGO
            rec.MoveNext
        Loop
        cboPais.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Call Centrar_pantalla(Me)

    lblEstado.Caption = ""
    GrdModulos.FormatString = "Código|Nombre"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 4000
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    'cargo combo pais
    LlenarComboPais
    'para la consulta
    Consulta = False 'no consulta true consulta
    Pais = ""
    Provincia = ""
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
        TxtCodigo = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        cmdGrabar.Enabled = True
        cmdBorrar.Enabled = True
        TxtCodigo_LostFocus
        tabDatos.Tab = 0
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 0 And Me.Visible Then
        txtnombre.SetFocus
        cmdGrabar.Enabled = True
        cmdBorrar.Enabled = True
    End If
    If tabDatos.Tab = 1 Then
        TxtNombreb.Text = ""
        TxtNombreb.SetFocus
        cmdGrabar.Enabled = False
        cmdBorrar.Enabled = False
    End If
End Sub

Private Sub TxtCodigo_Change()
    If Trim(TxtCodigo) = "" And cmdBorrar.Enabled Then
        cmdBorrar.Enabled = False
    ElseIf Trim(TxtCodigo) <> "" Then
        cmdBorrar.Enabled = True
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
        KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtdomici_GotFocus()
    SelecTexto txtDomici
End Sub

Private Sub txtdomici_KeyPress(KeyAscii As Integer)
   KeyAscii = Mayuscula(KeyAscii)
End Sub

Private Sub txtfax_GotFocus()
    SelecTexto txtFax
End Sub

Private Sub txtfax_KeyPress(KeyAscii As Integer)
   KeyAscii = Mayuscula(KeyAscii)
End Sub

Private Sub txtmail_GotFocus()
    SelecTexto txtMail
End Sub

Private Sub txtnombre_Change()
If Trim(txtnombre) = "" And cmdGrabar.Enabled Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If
End Sub

Private Function ValidarVendedor() As Boolean
    
    If txtnombre.Text = "" Then
        MsgBox "No ha ingresado el nombre", vbExclamation, TIT_MSGBOX
        txtnombre.SetFocus
        ValidarVendedor = False
        Exit Function
    End If
     
    
    If cboPais.ListIndex = -1 Then
        MsgBox "No ha seleccionado Pais", vbExclamation, TIT_MSGBOX
        cboPais.SetFocus
        ValidarVendedor = False
        Exit Function
    End If
    If cboProv.ListIndex = -1 Then
        MsgBox "No ha seleccionado Provincia", vbExclamation, TIT_MSGBOX
        cboProv.SetFocus
        ValidarVendedor = False
        Exit Function
    End If
    If cboLocal.ListIndex = -1 Then
        MsgBox "No ha seleccionado Localidad", vbExclamation, TIT_MSGBOX
        cboLocal.SetFocus
        ValidarVendedor = False
        Exit Function
    End If
    If txtDomici.Text = "" Then
        MsgBox "No ha ingresado el Domicilio", vbExclamation, TIT_MSGBOX
        txtDomici.SetFocus
        ValidarVendedor = False
        Exit Function
    End If
    ValidarVendedor = True
End Function

Private Sub TxtCodigo_LostFocus()
    If TxtCodigo.Text <> "" Then
        sql = "SELECT * FROM VENDEDOR"
        sql = sql & " WHERE VEN_CODIGO=" & XN(TxtCodigo)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Consulta = True
            txtnombre.Text = Rec1!VEN_NOMBRE
            txtDomici.Text = Rec1!VEN_DOMICI
            Call BuscaCodigoProxItemData(Rec1!PAI_CODIGO, cboPais)
            CboPais_LostFocus
            Pais = cboPais.Text
            
            Call BuscaCodigoProxItemData(Rec1!PRO_CODIGO, cboProv)
            cboProv_LostFocus
            Provincia = cboProv.Text
            
            Call BuscaCodigoProxItemData(Rec1!LOC_CODIGO, cboLocal)
            
            txtTelefono.Text = IIf(IsNull(Rec1!VEN_TELEFONO), "", Rec1!VEN_TELEFONO)
            txtFax.Text = IIf(IsNull(Rec1!VEN_FAX), "", Rec1!VEN_FAX)
            txtMail.Text = IIf(IsNull(Rec1!VEN_MAIL), "", Rec1!VEN_MAIL)
            txtnombre.SetFocus
        Else
            MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
            TxtCodigo.SetFocus
            Consulta = False
            Pais = ""
            Provincia = ""
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtnombre_GotFocus()
    SelecTexto txtnombre
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txttelefono_GotFocus()
    SelecTexto txtTelefono
End Sub
