VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ConsultaCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Cheque"
   ClientHeight    =   5685
   ClientLeft      =   150
   ClientTop       =   1065
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5685
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbSalirConsContra 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7170
      TabIndex        =   12
      Top             =   5175
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   150
      TabIndex        =   9
      Top             =   90
      Width           =   8490
      Begin VB.ComboBox cboOrden 
         Height          =   315
         ItemData        =   "ConsultaCh.frx":0000
         Left            =   5715
         List            =   "ConsultaCh.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1845
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   6450
         TabIndex        =   4
         Top             =   765
         Width           =   1110
      End
      Begin VB.ComboBox cboBusqueda 
         Height          =   315
         ItemData        =   "ConsultaCh.frx":0024
         Left            =   1665
         List            =   "ConsultaCh.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker FechaBusq 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   17170433
         CurrentDate     =   41098
      End
      Begin VB.TextBox txtCond_Busqueda 
         Height          =   285
         Left            =   3060
         TabIndex        =   3
         Top             =   750
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Ordenado por:"
         Height          =   255
         Left            =   4545
         TabIndex        =   13
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Búsqueda por:"
         Height          =   255
         Left            =   495
         TabIndex        =   11
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese la condición de Búsqueda:"
         Height          =   225
         Left            =   510
         TabIndex        =   10
         Top             =   780
         Width           =   2460
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5595
      TabIndex        =   6
      Top             =   5175
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid GrdCheques 
      Height          =   3645
      Left            =   150
      TabIndex        =   5
      Top             =   1425
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   6429
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColorSel    =   8388736
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.TextBox txtDes_Cons 
      Height          =   285
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4680
      Width           =   8430
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Seleccionado:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   330
      TabIndex        =   8
      Top             =   4410
      Width           =   1935
   End
End
Attribute VB_Name = "ConsultaCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CerrarSnapConsulta As Boolean
Dim snpConsulta As Recordset
Dim ventana As Form
Dim CrtlCodigo As Control

Public Function Parametros(auxVentana As Form, auxCrtlCodigo As Control)
    Set ventana = auxVentana 'Objeto ventana que llama a la ayuda
    Set CrtlCodigo = auxCrtlCodigo 'Objeto Control del form ventana al que se asigna el codigo
End Function

Private Sub cboBusqueda_Change()
    txtCond_Busqueda.Text = ""
End Sub

Private Sub cboBusqueda_Click()
    txtCond_Busqueda.Text = ""
End Sub

Private Sub cboBusqueda_LostFocus()
   If Me.cboBusqueda.List(Me.cboBusqueda.ListIndex) = "Número" Then
        Me.txtCond_Busqueda.Visible = True
        Me.FechaBusq.Visible = False
        Me.txtCond_Busqueda.SetFocus
   ElseIf Me.cboBusqueda.List(Me.cboBusqueda.ListIndex) = "Fecha de Vto" Then
        Me.txtCond_Busqueda.Visible = False
        Me.FechaBusq.Visible = True
        Me.FechaBusq.SetFocus
   End If
End Sub

Private Sub cmbSalirConsContra_Click()
    CrtlCodigo = ""
    Set ConsultaCheque = Nothing
    Unload Me
End Sub

Private Sub CmdAceptar_Click()
    Call ValidarIngreso
End Sub

Private Sub cmdBuscar_Click()
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    Screen.MousePointer = 11
    
    Select Case cboBusqueda.Text
       Case "Número"
              'TODOS los cheques, en CARTERA y con FECHA DE VTO. menor o igual a HOY
              sql = "SELECT c.che_numero, c.che_import, c.che_fecemi, c.che_fecvto, "
              sql = sql & " b.ban_banco, b.ban_localidad, b.ban_sucursal, b.ban_codigo, "
              sql = sql & " b.ban_descri, c.ech_descri "
              sql = sql & " FROM ChequeEstadoVigente c, Banco b "
              sql = sql & " WHERE c.ban_codint = b.ban_codint and c.ech_codigo = 1 And "
              sql = sql & " c.che_fecvto <= " & XDQ(Date)
              If Trim(txtCond_Busqueda.Text) <> "" Then
                 sql = sql & " and c.che_numero = " & XS(txtCond_Busqueda)
              End If
       Case "Fecha de Vto"
              sql = "SELECT c.che_numero, c.che_import, c.che_fecemi, c.che_fecvto, "
              sql = sql & " b.ban_banco, b.ban_localidad, b.ban_sucursal, b.ban_codigo, "
              sql = sql & " b.ban_descri, c.ech_descri "
              sql = sql & " FROM ChequeEstadoVigente c, Banco b "
              sql = sql & " WHERE c.ban_codint = b.ban_codint and c.ech_codigo = 1 And "
              sql = sql & " c.che_fecvto <= " & XDQ(Date)
              If Trim(txtCond_Busqueda.Text) <> "" Then
                sql = sql & " and c.che_fecvto = " & XDQ(Me.FechaBusq.Value)
              End If
    End Select
    
    Select Case cboOrden.Text
       Case "Número"
            sql = sql & " ORDER BY c.che_numero"
       Case "Fecha de Vto"
            sql = sql & " ORDER BY c.che_fecvto"
    End Select
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        GrdCheques.Rows = 1
        rec.MoveFirst
        Do While Not rec.EOF()
        
           GrdCheques.AddItem Trim(rec.Fields(0)) & Chr(9) & _
                              Format(rec.Fields(1), "$ #0.00") & Chr(9) & _
                              Trim(rec.Fields(2)) & Chr(9) & _
                              Trim(rec.Fields(3)) & Chr(9) & _
                              Trim(rec.Fields(4)) & Chr(9) & _
                              Trim(rec.Fields(5)) & Chr(9) & _
                              Trim(rec.Fields(6)) & Chr(9) & _
                              Trim(rec.Fields(7)) & Chr(9) & _
                              Trim(rec.Fields(8)) & Chr(9) & _
                              Trim(rec.Fields(9))
            rec.MoveNext
        Loop
        If Me.GrdCheques.Enabled Then
           Me.GrdCheques.row = 1
           Me.GrdCheques.SetFocus
        End If
    Else
        Me.GrdCheques.Rows = 1
        MsgBox "No se encontraron registros.", 16, TIT_MSGBOX
        txtDes_Cons.Text = ""
    End If
    rec.Close
    Screen.MousePointer = 1
    
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    CerrarSnapConsulta = False
    Call Centrar_pantalla(Me)
    
    GrdCheques.FormatString = "<Número|>Importe|>Fecha de Emisión|>Fecha de Vto.|^Banco|^Localidad|^Sucursal|<Código|<Banco|<Estado"
    GrdCheques.ColWidth(0) = 1000
    GrdCheques.ColWidth(1) = 1100
    GrdCheques.ColWidth(2) = 1500
    GrdCheques.ColWidth(3) = 1500
    GrdCheques.ColWidth(4) = 800
    GrdCheques.ColWidth(5) = 800
    GrdCheques.ColWidth(6) = 800
    GrdCheques.ColWidth(7) = 800
    GrdCheques.ColWidth(8) = 3500
    GrdCheques.ColWidth(9) = 3000
    GrdCheques.Rows = 1
    
    Me.cboBusqueda.Text = "Número"
    Me.cboOrden.Text = "Número"
    Me.Caption = "Consulta de Cheques"
    Me.lblDescripcion.Caption = "Ingrese Condición de Busqueda"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ConsultaCheque = Nothing
End Sub

Private Sub GrdCheques_Click()
    txtDes_Cons.Text = Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 0))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 1))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 2))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 3))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 4))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 5))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 6))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 7))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 8))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 9)))
End Sub

Private Sub GrdCheques_DblClick()
    CmdAceptar_Click
End Sub

Private Sub txtCond_Busqueda_KeyPress(KeyAscii As Integer)
    If cboBusqueda.Text = "Número" Then
        KeyAscii = NumeroEntero(KeyAscii)
    End If
End Sub

Private Function ValidarIngreso()
    If Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 0))) <> "" Then
        CrtlCodigo = Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 0)))
    End If
    Unload Me
End Function

Private Sub txtCond_Busqueda_LostFocus()
   If Me.cboBusqueda.Text = "Fecha de Vto" Then
       txtCond_Busqueda.Text = ValidarIngresoFecha(txtCond_Busqueda)
   End If
End Sub

Private Sub txtdes_cons_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


