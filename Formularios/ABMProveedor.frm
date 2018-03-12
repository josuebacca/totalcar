VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ABMProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM de Proveedores"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   48
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMProveedor.frx":0000
      Height          =   720
      Left            =   4800
      Picture         =   "ABMProveedor.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6135
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMProveedor.frx":0614
      Height          =   720
      Left            =   5685
      Picture         =   "ABMProveedor.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6135
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMProveedor.frx":0C28
      Height          =   720
      Left            =   3030
      Picture         =   "ABMProveedor.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6135
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMProveedor.frx":123C
      Height          =   720
      Left            =   3915
      Picture         =   "ABMProveedor.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6135
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   5970
      Left            =   75
      TabIndex        =   17
      Top             =   105
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   10530
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
      TabPicture(0)   =   "ABMProveedor.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMProveedor.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraDatos 
         Caption         =   " Datos del Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5265
         Left            =   150
         TabIndex        =   23
         Top             =   540
         Width           =   6225
         Begin VB.TextBox txtcodpostal 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4800
            MaxLength       =   10
            TabIndex        =   46
            Top             =   3360
            Width           =   1005
         End
         Begin VB.CommandButton cmdNuevoTipoProveedor 
            Height          =   315
            Left            =   4800
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProveedor.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Agregar Tipo Proveedor"
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboTipoProveedor 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   420
            Width           =   3375
         End
         Begin VB.CommandButton cmdNuevaProvincia 
            Height          =   315
            Left            =   4275
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProveedor.frx":1C12
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Agregar Provincia"
            Top             =   2955
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaLocalidad 
            Height          =   315
            Left            =   4305
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProveedor.frx":1F9C
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Agregar Localidad"
            Top             =   3330
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoPais 
            Height          =   315
            Left            =   4275
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProveedor.frx":2326
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Agregar País"
            Top             =   2580
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboIva 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1485
            Width           =   3375
         End
         Begin VB.TextBox txtIngBrutos 
            Height          =   285
            Left            =   3975
            MaxLength       =   10
            TabIndex        =   5
            Top             =   1860
            Width           =   1005
         End
         Begin VB.ComboBox cboProv 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2955
            Width           =   2880
         End
         Begin VB.ComboBox cboPais 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2580
            Width           =   2880
         End
         Begin VB.ComboBox cboLocal 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3330
            Width           =   2910
         End
         Begin VB.TextBox txtMail 
            Height          =   285
            Left            =   1350
            LinkTimeout     =   0
            MaxLength       =   40
            TabIndex        =   12
            Top             =   4830
            Width           =   4455
         End
         Begin VB.TextBox txtFax 
            Height          =   285
            Left            =   3975
            MaxLength       =   25
            TabIndex        =   11
            Top             =   4485
            Width           =   1815
         End
         Begin VB.TextBox txtDomici 
            Height          =   285
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   9
            Top             =   3720
            Width           =   4440
         End
         Begin VB.TextBox txtRazSoc 
            Height          =   285
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   2
            Tag             =   "Descripción"
            Top             =   1140
            Width           =   4335
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   285
            Left            =   1350
            MaxLength       =   40
            TabIndex        =   1
            Top             =   795
            Width           =   975
         End
         Begin VB.TextBox txtTelefono 
            Height          =   285
            Left            =   1350
            MaxLength       =   25
            TabIndex        =   10
            Top             =   4485
            Width           =   1815
         End
         Begin MSMask.MaskEdBox txtCuit 
            Height          =   315
            Left            =   1350
            TabIndex        =   4
            Top             =   1845
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   13
            Mask            =   "##-########-#"
            PromptChar      =   "_"
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "C.P."
            Height          =   195
            Left            =   5160
            TabIndex        =   47
            Top             =   3120
            Width           =   300
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Prov.:"
            Height          =   195
            Left            =   450
            TabIndex        =   44
            Top             =   450
            Width           =   780
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Vias de Comunicación"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   480
            TabIndex        =   40
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
            TabIndex        =   39
            Top             =   2265
            Width           =   630
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   540
            X2              =   5760
            Y1              =   2415
            Y2              =   2415
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cond. I.V.A.:"
            Height          =   195
            Left            =   330
            TabIndex        =   38
            Top             =   1515
            Width           =   900
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos:"
            Height          =   195
            Left            =   3045
            TabIndex        =   37
            Top             =   1905
            Width           =   810
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   630
            TabIndex        =   36
            Top             =   1905
            Width           =   600
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "País:"
            Height          =   195
            Left            =   855
            TabIndex        =   35
            Top             =   2625
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail:"
            Height          =   195
            Left            =   750
            TabIndex        =   34
            Top             =   4890
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Left            =   3555
            TabIndex        =   33
            Top             =   4530
            Width           =   300
         End
         Begin VB.Label Label7 
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   555
            TabIndex        =   32
            Top             =   4545
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   525
            TabIndex        =   31
            Top             =   3000
            Width           =   705
         End
         Begin VB.Label Label4 
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   495
            TabIndex        =   30
            Top             =   3375
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   555
            TabIndex        =   29
            Top             =   3750
            Width           =   675
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Raz. Soc.:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   25
            Top             =   1155
            Width           =   750
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
            Left            =   690
            TabIndex        =   24
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74865
         TabIndex        =   18
         Top             =   600
         Width           =   6135
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   360
            Left            =   5580
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProveedor.frx":26B0
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   420
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   330
            Left            =   1275
            MaxLength       =   15
            TabIndex        =   19
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
         Height          =   4350
         Left            =   -74880
         TabIndex        =   26
         Top             =   1410
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   7673
         _Version        =   393216
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
         TabIndex        =   27
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
      TabIndex        =   28
      Top             =   6300
      Width           =   750
   End
End
Attribute VB_Name = "ABMProveedor"
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

Private Sub cboLocal_Click()
    BuscoCodPostal XN(cboLocal.ItemData(cboLocal.ListIndex))
End Sub

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
    'cargo combo Localidad
    If cboProv.ListIndex <> -1 Then
        If Consulta = True And cboProv.Text = Provincia Then
            Exit Sub
        Else
            Provincia = ""
        End If
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
    If Trim(TxtCodigo) <> "" Then
        resp = MsgBox("Seguro desea eliminar el Proveedor: " & Trim(txtRazSoc.Text) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Eliminando ..."
        
        sql = "DELETE FROM PROVEEDOR"
        sql = sql & " WHERE TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
        sql = sql & " AND PROV_CODIGO = " & XN(TxtCodigo)
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
    
    sql = "SELECT TPR_CODIGO,PROV_CODIGO,PROV_RAZSOC"
    sql = sql & " FROM PROVEEDOR"
    sql = sql & " WHERE PROV_RAZSOC"
    sql = sql & " LIKE '%" & TxtDescriB.Text & "%' ORDER BY PROV_RAZSOC"
        
    lblEstado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Do While Not rec.EOF
           GrdModulos.AddItem rec!PROV_CODIGO & Chr(9) & rec!PROV_RAZSOC & Chr(9) & rec!TPR_CODIGO
           rec.MoveNext
        Loop
        If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        TxtDescriB.SetFocus
    End If
    rec.Close
    Text1.Text = GrdModulos.Rows - 1
    MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub CmdGrabar_Click()

    If ValidarProveedor = False Then Exit Sub
    
    On Error GoTo HayError
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    DBConn.BeginTrans
    If TxtCodigo.Text <> "" Then
        sql = "UPDATE PROVEEDOR "
        sql = sql & " SET PROV_RAZSOC=" & XS(txtRazSoc)
        sql = sql & " , IVA_CODIGO=" & cboIva.ItemData(cboIva.ListIndex)
        sql = sql & " , PROV_CUIT=" & XS(txtCuit)
        sql = sql & " , PROV_INGBRU=" & XS(txtIngBrutos)
        sql = sql & " , PROV_DOMICI=" & XS(txtDomici)
        sql = sql & " , PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
        sql = sql & " , PRO_CODIGO=" & cboProv.ItemData(cboProv.ListIndex)
        sql = sql & " , LOC_CODIGO=" & cboLocal.ItemData(cboLocal.ListIndex)
        sql = sql & " , PROV_TELEFONO=" & XS(txtTelefono)
        sql = sql & " , PROV_FAX=" & XS(txtFax)
        sql = sql & " , PROV_MAIL=" & XS(txtMail)
        sql = sql & " , TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
        sql = sql & " WHERE "
        sql = sql & " PROV_CODIGO=" & XN(TxtCodigo)
        DBConn.Execute sql
        
    Else
        TxtCodigo.Text = "1"
        sql = "SELECT MAX(PROV_CODIGO) as maximo FROM PROVEEDOR"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = XN(rec.Fields!Maximo) + 1
        rec.Close
        
        sql = "INSERT INTO PROVEEDOR(TPR_CODIGO,PROV_CODIGO,PROV_RAZSOC,PROV_DOMICI,"
        sql = sql & "PROV_CUIT,PROV_INGBRU,PROV_TELEFONO,PROV_FAX,PROV_MAIL,"
        sql = sql & "IVA_CODIGO,PAI_CODIGO,PRO_CODIGO,LOC_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex) & ","
        sql = sql & XN(TxtCodigo) & ","
        sql = sql & XS(txtRazSoc) & ","
        sql = sql & XS(txtDomici) & ","
        sql = sql & XS(txtCuit) & ","
        sql = sql & XS(txtIngBrutos) & ","
        sql = sql & XS(txtTelefono) & ","
        sql = sql & XS(txtFax) & ","
        sql = sql & XS(txtMail) & ","
        sql = sql & cboIva.ItemData(cboIva.ListIndex) & ","
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

Private Function ValidarProveedor() As Boolean
    If cboTipoProveedor.ListIndex = -1 Or cboTipoProveedor.ListIndex = 0 Then
        MsgBox "No ha elegido el Tipo de Proveedor", vbExclamation, TIT_MSGBOX
        cboTipoProveedor.SetFocus
        ValidarProveedor = False
        Exit Function
    End If
    If txtRazSoc.Text = "" Then
        MsgBox "No ha ingresado la Razón Social", vbExclamation, TIT_MSGBOX
        txtRazSoc.SetFocus
        ValidarProveedor = False
        Exit Function
    End If
    If cboIva.ListIndex = -1 Then
        MsgBox "No ha seleccionado condición de I.V.A.", vbExclamation, TIT_MSGBOX
        cboIva.SetFocus
        ValidarProveedor = False
        Exit Function
    End If
    If cboIva.ItemData(cboIva.ListIndex) <> 3 Then
        If txtCuit.Text = "" Then
            MsgBox "No ha ingresado el Número de C.U.I.T.", vbExclamation, TIT_MSGBOX
            txtCuit.SetFocus
            ValidarProveedor = False
            Exit Function
        End If
    End If
    If cboPais.ListIndex = -1 Then
        MsgBox "No ha seleccionado Pais", vbExclamation, TIT_MSGBOX
        cboPais.SetFocus
        ValidarProveedor = False
        Exit Function
    End If
    If cboProv.ListIndex = -1 Then
        MsgBox "No ha seleccionado Provincia", vbExclamation, TIT_MSGBOX
        cboProv.SetFocus
        ValidarProveedor = False
        Exit Function
    End If
    If cboLocal.ListIndex = -1 Then
        MsgBox "No ha seleccionado Localidad", vbExclamation, TIT_MSGBOX
        cboLocal.SetFocus
        ValidarProveedor = False
        Exit Function
    End If
    If txtDomici.Text = "" Then
        MsgBox "No ha ingresado el Domicilio", vbExclamation, TIT_MSGBOX
        txtDomici.SetFocus
        ValidarProveedor = False
        Exit Function
    End If
    ValidarProveedor = True
End Function

Private Sub cmdNuevaLocalidad_Click()
    Dim Localidad As Integer
   ' Localidad = cboLocal.ItemData(cboLocal.ListIndex)
    Consulta = False
    ABMLocalidad.Show vbModal
    cboLocal.Clear
    cboProv_LostFocus
    'Call BuscaCodigoProxItemData(Localidad, cboLocal)
End Sub

Private Sub cmdNuevaProvincia_Click()
    Dim Provincia As Integer
    Provincia = cboProv.ItemData(cboProv.ListIndex)
    ABMProvincia.Show vbModal
    Consulta = False
    cboProv.Clear
    Provincia = ""
    CboPais_LostFocus
    Call BuscaCodigoProxItemData(Provincia, cboProv)
End Sub

Private Sub CmdNuevo_Click()
    tabDatos.Tab = 0
    TxtCodigo.Text = ""
    txtRazSoc.Text = ""
    txtTelefono.Text = ""
    txtFax.Text = ""
    txtMail.Text = ""
    lblEstado.Caption = ""
    txtDomici.Text = ""
    txtIngBrutos.Text = ""
    txtCuit.Text = ""
    'para la consulta
    Consulta = False 'no consulta true consulta
    Pais = ""
    Provincia = ""
    GrdModulos.Rows = 1
    cboProv.Clear
    cboLocal.Clear
    cboPais.ListIndex = 0
    cboIva.ListIndex = 0
    cboTipoProveedor.ListIndex = 0
    txtcodpostal.Text = ""
    cboTipoProveedor.SetFocus
End Sub

Private Sub cmdNuevoPais_Click()
    ABMPais.Show vbModal
    Consulta = False
    cboPais.Clear
    Pais = ""
    LlenarComboPais
    cboPais.SetFocus
End Sub

Private Sub cmdNuevoTipoProveedor_Click()
    ABMTipoProveedor.Show vbModal
    cboTipoProveedor.Clear
    LlenarComboTipoProv
    cboTipoProveedor.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set ABMProveedor = Nothing
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
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    lblEstado.Caption = ""
    GrdModulos.FormatString = "Código|Razón Social|codigo Tipo"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 5000
    GrdModulos.ColWidth(2) = 0
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    'cargo combo tipo Provedor
    LlenarComboTipoProv
    'cargo combo pais
    LlenarComboPais
    'cargo combo IVA
    LlenarComboIva
    '-----------------
    'para la consulta
    Consulta = False 'no consulta true consulta
    Pais = ""
    Provincia = ""
End Sub

Private Sub LlenarComboTipoProv()
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
            cboTipoProveedor.AddItem "TODOS"
        Do While rec.EOF = False
            cboTipoProveedor.AddItem rec!TPR_DESCRI
            cboTipoProveedor.ItemData(cboTipoProveedor.NewIndex) = rec!TPR_CODIGO
            rec.MoveNext
        Loop
        cboTipoProveedor.ListIndex = 0
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

Private Sub LlenarComboIva()
    sql = "SELECT * FROM CONDICION_IVA ORDER BY IVA_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboIva.AddItem rec!IVA_DESCRI
            cboIva.ItemData(cboIva.NewIndex) = rec!IVA_CODIGO
            rec.MoveNext
        Loop
        cboIva.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
           Call BuscaCodigoProxItemData(XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 2)), cboTipoProveedor)
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
        txtRazSoc.SetFocus
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
        'cboTipoProveedor.ListIndex = 0
        CmdNuevo_Click
        TxtCodigo.SetFocus
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

Private Sub txtCUIT_LostFocus()
    If txtCuit.Text <> "" Then
        If ValidoCuit(txtCuit.Text) = False Then
         txtCuit.SetFocus
        End If
    End If
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

Private Sub txtIngBrutos_GotFocus()
    SelecTexto txtIngBrutos
End Sub

Private Sub txtIngBrutos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtRazSoc_Change()
If Trim(txtRazSoc) = "" And cmdGrabar.Enabled Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If
End Sub

Private Sub TxtCodigo_LostFocus()
    If TxtCodigo.Text <> "" Then
        sql = "SELECT * FROM PROVEEDOR"
        sql = sql & " WHERE PROV_CODIGO=" & XN(TxtCodigo)
        If cboTipoProveedor.ListIndex > 0 Then
            sql = sql & " AND TPR_CODIGO = " & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
        End If
        
        
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec1.EOF = False Then
            Consulta = True
            txtRazSoc.Text = Rec1!PROV_RAZSOC
            Call BuscaCodigoProxItemData(Rec1!IVA_CODIGO, cboIva)
            txtCuit.Text = IIf(IsNull(Rec1!PROV_CUIT), "", Rec1!PROV_CUIT)
            txtIngBrutos.Text = IIf(IsNull(Rec1!PROV_INGBRU), "", Rec1!PROV_INGBRU)
            
            Call BuscaCodigoProxItemData(Rec1!TPR_CODIGO, cboTipoProveedor)
            
            Call BuscaCodigoProxItemData(Rec1!PAI_CODIGO, cboPais)
            CboPais_LostFocus
            Pais = cboPais.Text
            
            Call BuscaCodigoProxItemData(Rec1!PRO_CODIGO, cboProv)
            cboProv_LostFocus
            Provincia = cboProv.Text
            
            Call BuscaCodigoProxItemData(Rec1!LOC_CODIGO, cboLocal)
            BuscoCodPostal Rec1!LOC_CODIGO
            
            txtDomici.Text = Rec1!PROV_DOMICI
            txtTelefono.Text = IIf(IsNull(Rec1!PROV_TELEFONO), "", Rec1!PROV_TELEFONO)
            txtFax.Text = IIf(IsNull(Rec1!PROV_FAX), "", Rec1!PROV_FAX)
            txtMail.Text = IIf(IsNull(Rec1!PROV_MAIL), "", Rec1!PROV_MAIL)
            txtRazSoc.SetFocus
        Else
            MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
            TxtCodigo.Text = ""
            TxtCodigo.SetFocus
            Consulta = False
            Pais = ""
            Provincia = ""
        End If
        Rec1.Close
    End If
End Sub
Function BuscoCodPostal(Codigo As Integer) As String
    sql = "SELECT LOC_CODPOS FROM LOCALIDAD "
    sql = sql & "WHERE PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex) & ""
    sql = sql & " AND PRO_CODIGO = " & cboProv.ItemData(cboProv.ListIndex) & ""
    sql = sql & " AND LOC_CODIGO = " & Codigo
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        txtcodpostal.Text = IIf(IsNull(Rec2!LOC_CODPOS), "", Rec2!LOC_CODPOS)
    End If
    Rec2.Close
End Function

Private Sub txtmail_GotFocus()
    SelecTexto txtMail
End Sub

Private Sub txtmail_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtRazSoc_GotFocus()
    SelecTexto txtRazSoc
End Sub

Private Sub txtRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txttelefono_GotFocus()
    SelecTexto txtTelefono
End Sub
