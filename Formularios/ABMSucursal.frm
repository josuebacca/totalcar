VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form ABMSucursal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM de Sucursales"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMSucursal.frx":0000
      Height          =   735
      Left            =   4815
      Picture         =   "ABMSucursal.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMSucursal.frx":0614
      Height          =   735
      Left            =   5700
      Picture         =   "ABMSucursal.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6480
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMSucursal.frx":0C28
      Height          =   735
      Left            =   3045
      Picture         =   "ABMSucursal.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6480
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMSucursal.frx":123C
      Height          =   735
      Left            =   3930
      Picture         =   "ABMSucursal.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6360
      Left            =   75
      TabIndex        =   19
      Top             =   75
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   11218
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
      TabPicture(0)   =   "ABMSucursal.frx":1850
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMSucursal.frx":186C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   135
         TabIndex        =   20
         Top             =   600
         Width           =   6135
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   345
            Left            =   5580
            MaskColor       =   &H000000FF&
            Picture         =   "ABMSucursal.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Buscar"
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   420
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   330
            Left            =   1275
            MaxLength       =   15
            TabIndex        =   16
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
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   270
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4695
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame fraDatos 
         Caption         =   " Datos de la Sucursal "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5685
         Left            =   -74850
         TabIndex        =   23
         Top             =   540
         Width           =   6225
         Begin VB.ComboBox cboVendedor 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1875
            Width           =   3375
         End
         Begin VB.TextBox txtClienteDescri 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   465
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   43
            Top             =   660
            Width           =   4320
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   2385
            MaskColor       =   &H000000FF&
            Picture         =   "ABMSucursal.frx":1B92
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Buscar"
            Top             =   315
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Left            =   1350
            MaxLength       =   40
            TabIndex        =   1
            Top             =   1185
            Width           =   975
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Height          =   315
            Left            =   2835
            MaskColor       =   &H000000FF&
            Picture         =   "ABMSucursal.frx":1E9C
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Buscar"
            Top             =   315
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaProvincia 
            Height          =   315
            Left            =   4185
            MaskColor       =   &H000000FF&
            Picture         =   "ABMSucursal.frx":2226
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Buscar"
            Top             =   2955
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaLocalidad 
            Height          =   315
            Left            =   4785
            MaskColor       =   &H000000FF&
            Picture         =   "ABMSucursal.frx":25B0
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Buscar"
            Top             =   3345
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoPais 
            Height          =   315
            Left            =   4185
            MaskColor       =   &H000000FF&
            Picture         =   "ABMSucursal.frx":293A
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Buscar"
            Top             =   2580
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboProv 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2970
            Width           =   2790
         End
         Begin VB.ComboBox cboPais 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2595
            Width           =   2790
         End
         Begin VB.ComboBox cboLocal 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   3345
            Width           =   3375
         End
         Begin VB.TextBox txtMail 
            Height          =   285
            Left            =   1350
            LinkTimeout     =   0
            MaxLength       =   40
            TabIndex        =   10
            Top             =   4860
            Width           =   4455
         End
         Begin VB.TextBox txtFax 
            Height          =   285
            Left            =   3975
            MaxLength       =   25
            TabIndex        =   9
            Top             =   4515
            Width           =   1815
         End
         Begin VB.TextBox txtDomici 
            Height          =   300
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   7
            Top             =   3720
            Width           =   4335
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   300
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   2
            Tag             =   "Descripción"
            Top             =   1530
            Width           =   4335
         End
         Begin VB.TextBox txtCodigoCli 
            Height          =   300
            Left            =   1350
            MaxLength       =   40
            TabIndex        =   0
            Top             =   315
            Width           =   975
         End
         Begin VB.TextBox txtTelefono 
            Height          =   285
            Left            =   1350
            MaxLength       =   25
            TabIndex        =   8
            Top             =   4515
            Width           =   1815
         End
         Begin VB.TextBox txtContacto 
            Height          =   285
            Left            =   1350
            LinkTimeout     =   0
            MaxLength       =   40
            TabIndex        =   11
            Top             =   5205
            Width           =   4455
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Contacto:"
            Height          =   195
            Left            =   540
            TabIndex        =   45
            Top             =   5265
            Width           =   690
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   495
            TabIndex        =   44
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código Suc.:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   41
            Top             =   1230
            Width           =   915
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Vias de Comunicación"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   480
            TabIndex        =   36
            Top             =   4140
            Width           =   1575
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            X1              =   540
            X2              =   5760
            Y1              =   4290
            Y2              =   4290
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   480
            TabIndex        =   35
            Top             =   2280
            Width           =   630
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   540
            X2              =   5760
            Y1              =   2430
            Y2              =   2430
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "País:"
            Height          =   195
            Left            =   855
            TabIndex        =   34
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail:"
            Height          =   195
            Left            =   750
            TabIndex        =   33
            Top             =   4920
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Left            =   3555
            TabIndex        =   32
            Top             =   4560
            Width           =   300
         End
         Begin VB.Label Label7 
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   555
            TabIndex        =   31
            Top             =   4575
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   525
            TabIndex        =   30
            Top             =   3015
            Width           =   705
         End
         Begin VB.Label Label4 
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   495
            TabIndex        =   29
            Top             =   3390
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   555
            TabIndex        =   28
            Top             =   3750
            Width           =   675
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrición:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   405
            TabIndex        =   25
            Top             =   1545
            Width           =   795
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código Cli.:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   24
            Top             =   375
            Width           =   795
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   26
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
      TabIndex        =   27
      Top             =   6615
      Width           =   750
   End
End
Attribute VB_Name = "ABMSucursal"
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

Private Sub CboPais_LostFocus()
   If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Then Exit Sub
    If Consulta = True And cboPais.Text = Pais Then
        Exit Sub
    Else
        Pais = ""
    End If
    'cargo combo PROVINCIA
    cboProv.Clear
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
    End If
    rec.Close
End Sub

Private Sub cboProv_LostFocus()
    If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Then Exit Sub
    If cboProv.ListIndex <> -1 Then
        If Consulta = True And cboProv.Text = Provincia Then
            Exit Sub
        Else
            Provincia = ""
        End If
        'cargo combo Localidad
        cboLocal.Clear
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
        End If
        rec.Close
    End If
End Sub

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(TxtCodigo) <> "" And Trim(TxtCodigoCli) <> "" Then
        resp = MsgBox("Seguro desea eliminar la Sucursal: " & Trim(txtDescripcion.Text) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Eliminando ..."

        sql = "DELETE FROM SUCURSAL "
        sql = sql & " WHERE CLI_CODIGO = " & XN(TxtCodigoCli)
        sql = sql & " AND SUC_CODIGO=" & XN(TxtCodigo)
        DBConn.Execute sql
        
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        CmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    MousePointer = vbHourglass
    
    sql = "SELECT SUC_CODIGO,CLI_CODIGO,SUC_DESCRI "
    sql = sql & " FROM SUCURSAL"
    sql = sql & " WHERE SUC_DESCRI"
    sql = sql & " LIKE '" & TxtDescriB.Text & "%' ORDER BY SUC_DESCRI"
        
    lblEstado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Do While Not rec.EOF
           GrdModulos.AddItem rec.Fields(0) & Chr(9) & rec.Fields(2) & Chr(9) & rec.Fields(1)
           rec.MoveNext
        Loop
        If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        TxtDescriB.SetFocus
    End If
    rec.Close
    MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Rows > 1 Then
        frmBuscar.grdBuscar.Col = 0
        TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtClienteDescri.Text = frmBuscar.grdBuscar.Text
        TxtCodigo.SetFocus
    Else
        TxtCodigoCli.SetFocus
    End If
End Sub

Private Sub CmdGrabar_Click()

    If ValidarCliente = False Then Exit Sub
    
    On Error GoTo HayError
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    DBConn.BeginTrans
    If TxtCodigo.Text <> "" Then
        sql = "UPDATE SUCURSAL "
        sql = sql & " SET SUC_DESCRI=" & XS(txtDescripcion)
        sql = sql & " , SUC_DOMICI=" & XS(txtDomici)
        sql = sql & " , PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
        sql = sql & " , PRO_CODIGO=" & cboProv.ItemData(cboProv.ListIndex)
        sql = sql & " , LOC_CODIGO=" & cboLocal.ItemData(cboLocal.ListIndex)
        sql = sql & " , SUC_TELEFONO=" & XS(txtTelefono)
        sql = sql & " , SUC_FAX=" & XS(txtFax)
        sql = sql & " , SUC_MAIL=" & XS(txtMail)
        sql = sql & " , VEN_CODIGO=" & cboVendedor.ItemData(cboVendedor.ListIndex)
        sql = sql & " ,SUC_CONTACTO=" & XS(txtContacto.Text)
        sql = sql & " WHERE CLI_CODIGO=" & XN(TxtCodigoCli)
        sql = sql & " AND SUC_CODIGO=" & XN(TxtCodigo)
        DBConn.Execute sql
        
    Else
        TxtCodigo = "1"
        sql = "SELECT MAX(SUC_CODIGO) as maximo FROM SUCURSAL"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = XN(rec.Fields!Maximo) + 1
        rec.Close
        
        sql = "INSERT INTO SUCURSAL(SUC_CODIGO,CLI_CODIGO,SUC_DESCRI,SUC_DOMICI,"
        sql = sql & "SUC_TELEFONO,SUC_FAX,SUC_MAIL,SUC_CONTACTO,"
        sql = sql & "VEN_CODIGO,PAI_CODIGO,PRO_CODIGO,LOC_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(TxtCodigo) & ","
        sql = sql & XN(TxtCodigoCli) & ","
        sql = sql & XS(txtDescripcion) & ","
        sql = sql & XS(txtDomici) & ","
        sql = sql & XS(txtTelefono) & ","
        sql = sql & XS(txtFax) & ","
        sql = sql & XS(txtMail) & ","
        sql = sql & XS(txtContacto) & ","
        sql = sql & cboVendedor.ItemData(cboVendedor.ListIndex) & ","
        sql = sql & cboPais.ItemData(cboPais.ListIndex) & ","
        sql = sql & cboProv.ItemData(cboProv.ListIndex) & ","
        sql = sql & cboLocal.ItemData(cboLocal.ListIndex) & ")"

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

Private Function ValidarCliente() As Boolean
    If TxtCodigoCli.Text = "" Then
        MsgBox "No ha ingresado el Cliente", vbExclamation, TIT_MSGBOX
        TxtCodigoCli.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    If txtDescripcion.Text = "" Then
        MsgBox "No ha ingresado la Razón Social", vbExclamation, TIT_MSGBOX
        txtDescripcion.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    If cboVendedor.ListIndex = -1 Then
        MsgBox "No ha seleccionado Vendedor", vbExclamation, TIT_MSGBOX
        cboVendedor.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    If cboPais.ListIndex = -1 Then
        MsgBox "No ha seleccionado Pais", vbExclamation, TIT_MSGBOX
        cboPais.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    If cboProv.ListIndex = -1 Then
        MsgBox "No ha seleccionado Provincia", vbExclamation, TIT_MSGBOX
        cboProv.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    If cboLocal.ListIndex = -1 Then
        MsgBox "No ha seleccionado Localidad", vbExclamation, TIT_MSGBOX
        cboLocal.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    If txtDomici.Text = "" Then
        MsgBox "No ha ingresado el Domicilio", vbExclamation, TIT_MSGBOX
        txtDomici.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    ValidarCliente = True
End Function

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
    TxtCodigoCli.Text = ""
    txtClienteDescri.Text = ""
    txtDescripcion.Text = ""
    txtTelefono.Text = ""
    txtFax.Text = ""
    txtMail.Text = ""
    lblEstado.Caption = ""
    txtDomici.Text = ""
    txtContacto.Text = ""
    Consulta = False
    Pais = ""
    Provincia = ""
    GrdModulos.Rows = 1
    cboProv.Clear
    cboLocal.Clear
    cboPais.ListIndex = 0
    cboVendedor.ListIndex = 0
    TxtCodigoCli.SetFocus
End Sub

Private Sub cmdNuevoCliente_Click()
    ABMCliente.Show vbModal
    TxtCodigoCli.SetFocus
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
    Set ABMSucursal = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    lblEstado.Caption = ""
    GrdModulos.FormatString = "Código|Razón Social|Cliente"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 5000
    GrdModulos.ColWidth(2) = 0
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    
    'cargo combo pais
    LlenarComboPais
    'CARGO COMBO VENDEDOR
    LLenarComboVendedor
    'para la consulta
    Consulta = False 'no consulta true consulta
    Pais = ""
    Provincia = ""
    
End Sub

Private Sub LLenarComboVendedor()
    sql = "SELECT VEN_CODIGO,VEN_NOMBRE FROM VENDEDOR ORDER BY VEN_NOMBRE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboVendedor.AddItem rec!VEN_NOMBRE
            cboVendedor.ItemData(cboVendedor.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        cboVendedor.ListIndex = 0
    End If
    rec.Close
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

Private Sub GrdModulos_DblClick()
 If GrdModulos.row > 0 Then
        TxtCodigo = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        TxtCodigoCli.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
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
        If TxtCodigo <> "" And TxtCodigoCli <> "" Then
            txtDescripcion.SetFocus
        Else
            TxtCodigoCli.SetFocus
        End If
        cmdGrabar.Enabled = True
        cmdBorrar.Enabled = True
    End If
    If tabDatos.Tab = 1 Then
        TxtDescriB.Text = ""
        TxtDescriB.SetFocus
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

Private Sub TxtCodigoCli_GotFocus()
    SelecTexto TxtCodigoCli
End Sub

Private Sub txtCodigoCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigoCli_LostFocus()
    If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Or ActiveControl.Name = "CmdSalir" Then Exit Sub
    If TxtCodigoCli.Text <> "" Then
        sql = "SELECT * FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(TxtCodigoCli)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtClienteDescri.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
            TxtCodigoCli.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtContacto_GotFocus()
    SelecTexto txtContacto
End Sub

Private Sub txtContacto_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtdomici_GotFocus()
    SelecTexto txtDomici
End Sub

Private Sub txtdomici_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtfax_GotFocus()
    SelecTexto txtFax
End Sub

Private Sub txtDescripcion_Change()
If Trim(txtDescripcion) = "" And cmdGrabar.Enabled Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If
End Sub

Private Sub TxtCodigo_LostFocus()
    If TxtCodigo.Text <> "" Then
        sql = "SELECT * FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(TxtCodigo)
        If TxtCodigoCli.Text <> "" Then
         sql = sql & " AND CLI_CODIGO=" & XN(TxtCodigoCli)
        End If
        
        Set Rec1 = New ADODB.Recordset
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            If Rec1.RecordCount > 1 Then
             lblEstado.Caption = "Buscando..."
             tabDatos.Tab = 1
                Do While Rec1.EOF = False
                    GrdModulos.AddItem rec.Fields(0) & Chr(9) & rec.Fields(2) & Chr(9) & rec.Fields(1)
                    rec.MoveNext
                Loop
                Rec1.Close
                lblEstado.Caption = ""
             If GrdModulos.Enabled Then GrdModulos.SetFocus
             Exit Sub
            End If
            Consulta = True
            TxtCodigoCli.Text = Rec1!CLI_CODIGO
            TxtCodigoCli_LostFocus
            txtDescripcion.Text = Rec1!SUC_DESCRI
            Call BuscaCodigoProxItemData(Rec1!VEN_CODIGO, cboVendedor)
            
            Call BuscaCodigoProxItemData(Rec1!PAI_CODIGO, cboPais)
            CboPais_LostFocus
            Pais = cboPais.Text
            
            Call BuscaCodigoProxItemData(Rec1!PRO_CODIGO, cboProv)
            cboProv_LostFocus
            Provincia = cboProv.Text
            
            Call BuscaCodigoProxItemData(Rec1!LOC_CODIGO, cboLocal)
            txtDomici.Text = Rec1!SUC_DOMICI
            txtTelefono.Text = ChkNull(Rec1!SUC_TELEFONO)
            txtFax.Text = ChkNull(Rec1!SUC_FAX)
            txtMail.Text = ChkNull(Rec1!SUC_MAIL)
            txtContacto.Text = ChkNull(Rec1!SUC_CONTACTO)
            txtDescripcion.SetFocus
        Else
         MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
         Consulta = False
         Pais = ""
         Provincia = ""
         TxtCodigo.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtmail_GotFocus()
    SelecTexto txtMail
End Sub

Private Sub txtmail_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescripcion_GotFocus()
    SelecTexto txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txttelefono_GotFocus()
    SelecTexto txtTelefono
End Sub
