VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form frmGastosVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos del Vendedor"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmGastosVendedor.frx":0000
      Height          =   720
      Left            =   3645
      Picture         =   "frmGastosVendedor.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3645
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmGastosVendedor.frx":0614
      Height          =   720
      Left            =   2760
      Picture         =   "frmGastosVendedor.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3645
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmGastosVendedor.frx":0C28
      Height          =   720
      Left            =   5400
      Picture         =   "frmGastosVendedor.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3645
      Width           =   870
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "frmGastosVendedor.frx":123C
      Height          =   720
      Left            =   4530
      Picture         =   "frmGastosVendedor.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3645
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   3510
      Left            =   45
      TabIndex        =   14
      Top             =   75
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
      TabPicture(0)   =   "frmGastosVendedor.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmGastosVendedor.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraDatos 
         Caption         =   "Gastos del Vendedor"
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
         Left            =   300
         TabIndex        =   13
         Top             =   495
         Width           =   5580
         Begin VB.CommandButton cmdNuevoGasto 
            Height          =   315
            Left            =   4845
            MaskColor       =   &H000000FF&
            Picture         =   "frmGastosVendedor.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Agergar Tipo de Gasto"
            Top             =   1455
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin FechaCtl.Fecha FechaGasto 
            Height          =   330
            Left            =   1065
            TabIndex        =   0
            Top             =   690
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.TextBox txtImporte 
            Height          =   315
            Left            =   1035
            MaxLength       =   6
            TabIndex        =   3
            Tag             =   "Descripción"
            Top             =   1830
            Width           =   885
         End
         Begin VB.ComboBox CboGastos 
            Height          =   315
            Left            =   1050
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1440
            Width           =   3765
         End
         Begin VB.ComboBox CboVend 
            Height          =   315
            Left            =   1050
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1065
            Width           =   3765
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Importe:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   21
            Top             =   1875
            Width           =   570
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   390
            TabIndex        =   18
            Top             =   1485
            Width           =   540
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   17
            Top             =   1110
            Width           =   735
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   435
            TabIndex        =   16
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74850
         TabIndex        =   12
         Top             =   390
         Width           =   5940
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   345
            Left            =   5325
            MaskColor       =   &H000000FF&
            Picture         =   "frmGastosVendedor.frx":1C12
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Buscar"
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   315
            Left            =   1230
            MaxLength       =   15
            TabIndex        =   8
            Top             =   255
            Width           =   3975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   270
            Width           =   735
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   2205
         Left            =   -74865
         TabIndex        =   10
         Top             =   1200
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   3889
         _Version        =   393216
         Cols            =   4
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
         TabIndex        =   19
         Top             =   570
         Width           =   1065
      End
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
      Left            =   135
      TabIndex        =   20
      Top             =   3795
      Width           =   750
   End
End
Attribute VB_Name = "frmGastosVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim resp As Integer

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
        resp = MsgBox("Seguro desea eliminar este Gasto del Vendedor?", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Eliminando ..."
        sql = "DELETE FROM GASTOS_VENDEDOR"
        sql = sql & "WHERE "
        sql = sql & " VEN_CODIGO=" & CboVend.ItemData(CboVend.ListIndex)
        sql = sql & " AND TGT_CODIGO=" & CboGastos.ItemData(CboGastos.ListIndex)
        sql = sql & " AND GAV_FECHA=" & FechaGasto.Text
        DBConn.Execute sql
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        cmdNuevo_Click
    
    Exit Sub
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    MousePointer = vbHourglass
    
    sql = "SELECT GV.GAV_FECHA,V.VEN_NOMBRE,G.TGT_DESCRI,GV.GAV_IMPORTE,V.VEN_CODIGO , G.TGT_CODIGO "
    sql = sql & " FROM GASTOS_VENDEDOR GV, VENDEDOR V,TIPO_GASTO G"
    sql = sql & " WHERE GV.VEN_CODIGO = V.VEN_CODIGO "
    sql = sql & " AND GV.TGT_CODIGO= G.TGT_CODIGO AND V.VEN_NOMBRE"
    sql = sql & " LIKE '" & TxtDescriB.Text & "%' ORDER BY V.VEN_NOMBRE"
        
    lblEstado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Do While Not rec.EOF
           GrdModulos.AddItem rec.Fields(0) & Chr(9) & rec.Fields(1) & Chr(9) & _
                              rec.Fields(2) & Chr(9) & Format(rec.Fields(3), "0.00") & Chr(9) & _
                              rec.Fields(4) & Chr(9) & rec.Fields(5)
                              
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

Private Sub CmdGrabar_Click()
    
    If ValidarGastos = False Then Exit Sub
    
    On Error GoTo HayError
    DBConn.BeginTrans
    sql = "SELECT * FROM GASTOS_VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO=" & CboVend.ItemData(CboVend.ListIndex)
    sql = sql & " AND TGT_CODIGO=" & CboGastos.ItemData(CboGastos.ListIndex)
    sql = sql & " AND GAV_FECHA = " & XDQ(FechaGasto.Text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
   
    If rec.EOF = False Then
        sql = "UPDATE GASTOS_VENDEDOR"
        sql = sql & " SET GAV_IMPORTE=" & XN(txtImporte)
        sql = sql & " WHERE  VEN_CODIGO =" & CboVend.ItemData(CboVend.ListIndex)
        sql = sql & " AND TGT_CODIGO=" & CboGastos.ItemData(CboGastos.ListIndex)
        sql = sql & " AND GAV_FECHA=" & XDQ(FechaGasto.Text)
        DBConn.Execute sql
        
    Else
        sql = "INSERT INTO GASTOS_VENDEDOR(TGT_CODIGO,VEN_CODIGO,GAV_FECHA,GAV_IMPORTE)"
        sql = sql & " VALUES ("
        sql = sql & CboGastos.ItemData(CboGastos.ListIndex) & ","
        sql = sql & CboVend.ItemData(CboVend.ListIndex) & ","
        sql = sql & XDQ(FechaGasto) & ","
        sql = sql & XN(txtImporte) & ")"
        DBConn.Execute sql
    End If
    
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.CommitTrans
    cmdNuevo_Click
    Exit Sub
    
HayError:
    If rec.State = 1 Then rec.Close
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub cmdNuevo_Click()
    txtImporte.Text = "0,00"
    lblEstado.Caption = ""
    GrdModulos.Rows = 1
    CboVend.Enabled = True
    CboGastos.Enabled = True
    CboVend.ListIndex = 0
    FechaGasto.Text = Date
    If CboGastos.ListCount > 0 Then
        CboGastos.ListIndex = 0
    End If
    CboVend.SetFocus
End Sub

Private Sub cmdNuevoGasto_Click()
    ABMTipoGasto.Show vbModal
    CboGastos.Clear
    llenarComboGastos
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set frmGastosVendedor = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    lblEstado.Caption = ""
    'ARMO LA GRILLA DE DATOS
    preparo_grilla
    tabDatos.Tab = 0
    'CARGO COMBO VENDEDOR
    LLenarComboVendedor
    'CARGO COMBO GASTOS
    llenarComboGastos
    FechaGasto.Text = Date
    txtImporte.Text = "0,00"
End Sub

Private Sub preparo_grilla()
    GrdModulos.FormatString = "Fecha|Vendedor|Gasto|Import($)|CODVEN|CODGAS"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 2500
    GrdModulos.ColWidth(2) = 2500
    GrdModulos.ColWidth(3) = 800
    GrdModulos.ColWidth(4) = 0
    GrdModulos.ColWidth(5) = 0
    GrdModulos.Rows = 1
End Sub

Private Sub LLenarComboVendedor()
    sql = "SELECT * FROM VENDEDOR ORDER BY VEN_NOMBRE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            CboVend.AddItem rec!VEN_NOMBRE
            CboVend.ItemData(CboVend.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        CboVend.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub llenarComboGastos()
    sql = "SELECT * FROM TIPO_GASTO ORDER BY TGT_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            CboGastos.AddItem rec!TGT_DESCRI
            CboGastos.ItemData(CboGastos.NewIndex) = rec!TGT_CODIGO
            rec.MoveNext
        Loop
        CboGastos.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_dblClick()
    If GrdModulos.row > 0 Then
           Call BuscaCodigoProxItemData(GrdModulos.TextMatrix(GrdModulos.RowSel, 4), CboVend)
           Call BuscaCodigoProxItemData(GrdModulos.TextMatrix(GrdModulos.RowSel, 5), CboGastos)
           txtImporte.Text = Format(GrdModulos.TextMatrix(GrdModulos.RowSel, 3), "0.00")
           FechaGasto.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
           cmdGrabar.Enabled = True
           cmdBorrar.Enabled = True
           CboVend.Enabled = False
           CboGastos.Enabled = False
           txtImporte.SetFocus
           tabDatos.Tab = 0
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 0 And Me.Visible Then
        FechaGasto.SetFocus
        cmdGrabar.Enabled = True
        cmdBorrar.Enabled = True
    End If
    If tabDatos.Tab = 1 Then
        TxtDescriB.Text = ""
        If TxtDescriB.Enabled Then TxtDescriB.SetFocus
        cmdGrabar.Enabled = False
        cmdBorrar.Enabled = False
    End If
End Sub

Private Sub TxtDescriB_GotFocus()
    SelecTexto TxtDescriB
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Function ValidarGastos() As Boolean
    If CboVend.ListIndex = -1 Then
        MsgBox "No ha seleccionado un Vendedor", vbExclamation, TIT_MSGBOX
        CboVend.SetFocus
        ValidarGastos = False
        Exit Function
    End If
    
    If CboGastos.ListIndex = -1 Then
        MsgBox "No ha seleccionado un Gasto", vbExclamation, TIT_MSGBOX
        CboGastos.SetFocus
        ValidarGastos = False
        Exit Function
    End If
    
    If txtImporte.Text = "" Then
        MsgBox "No ha ingresado el Importe del Gasto", vbExclamation, TIT_MSGBOX
        txtImporte.SetFocus
        ValidarGastos = False
        Exit Function
    End If
    If FechaGasto.Text = "" Then
        MsgBox "No ha ingresado la Fecha ", vbExclamation, TIT_MSGBOX
        FechaGasto.SetFocus
        ValidarGastos = False
        Exit Function
    End If
    
    ValidarGastos = True
End Function

Private Sub txtImporte_GotFocus()
    SelecTexto txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImporte, KeyAscii)
End Sub

Private Sub txtImporte_LostFocus()
    If txtImporte.Text <> "" Then
        txtImporte.Text = Valido_Importe(txtImporte)
    Else
        txtImporte.Text = "0,00"
    End If
End Sub
