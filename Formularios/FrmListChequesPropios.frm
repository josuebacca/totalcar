VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form FrmListChequesPropios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cheques Propios"
   ClientHeight    =   5955
   ClientLeft      =   1365
   ClientTop       =   975
   ClientWidth     =   8760
   Icon            =   "FrmListChequesPropios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraImp 
      Caption         =   "Impresión de Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5835
      Left            =   105
      TabIndex        =   21
      Top             =   75
      Width           =   8595
      Begin VB.Frame Frame1 
         Height          =   3330
         Left            =   3885
         TabIndex        =   25
         Top             =   390
         Width           =   4590
         Begin VB.ComboBox cboCtaBancaria 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1140
            Width           =   2100
         End
         Begin FechaCtl.Fecha TxtFecIngresoD 
            Height          =   300
            Left            =   1530
            TabIndex        =   9
            Top             =   1965
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.ComboBox CboEstado 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2385
            Width           =   2985
         End
         Begin VB.TextBox TxtNroCheque 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1530
            TabIndex        =   8
            Top             =   1545
            Width           =   1080
         End
         Begin VB.ComboBox CboBanco 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   735
            Width           =   2985
         End
         Begin FechaCtl.Fecha TxtFecIngresoH 
            Height          =   300
            Left            =   3060
            TabIndex        =   10
            Top             =   1965
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha Fecha2 
            Height          =   300
            Left            =   3060
            TabIndex        =   13
            Top             =   2790
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha Fecha1 
            Height          =   300
            Left            =   1530
            TabIndex        =   12
            Top             =   2790
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha TxtFecVtoD 
            Height          =   300
            Left            =   1530
            TabIndex        =   4
            Top             =   345
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha TxtFecVtoH 
            Height          =   300
            Left            =   3060
            TabIndex        =   5
            Top             =   345
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cuenta:"
            Height          =   195
            Index           =   4
            Left            =   525
            TabIndex        =   38
            Top             =   1170
            Width           =   855
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   870
            TabIndex        =   34
            Top             =   2850
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   2
            Left            =   2775
            TabIndex        =   33
            Top             =   2835
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   1
            Left            =   2790
            TabIndex        =   32
            Top             =   390
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   0
            Left            =   2790
            TabIndex        =   31
            Top             =   1980
            Width           =   120
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   840
            TabIndex        =   30
            Top             =   2460
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   870
            TabIndex        =   29
            Top             =   795
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nro de Cheque:"
            Height          =   195
            Left            =   255
            TabIndex        =   28
            Top             =   1575
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión:"
            Height          =   195
            Left            =   75
            TabIndex        =   27
            Top             =   2010
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Pago:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   375
            Width           =   1140
         End
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         DisabledPicture =   "FrmListChequesPropios.frx":27A2
         Height          =   720
         Left            =   7500
         Picture         =   "FrmListChequesPropios.frx":2BEC
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5025
         Width           =   915
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Aceptar"
         DisabledPicture =   "FrmListChequesPropios.frx":2EF6
         Height          =   720
         Left            =   5640
         Picture         =   "FrmListChequesPropios.frx":37C0
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5025
         Width           =   915
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         DisabledPicture =   "FrmListChequesPropios.frx":3ACA
         Height          =   720
         Left            =   6570
         Picture         =   "FrmListChequesPropios.frx":4394
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5025
         Width           =   915
      End
      Begin VB.Frame fraSentido 
         Caption         =   "Sentido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   165
         TabIndex        =   24
         Top             =   2760
         Width           =   3660
         Begin VB.OptionButton oDescendente 
            Caption         =   "Descendente"
            Height          =   255
            Left            =   1965
            TabIndex        =   15
            Top             =   435
            Width           =   1335
         End
         Begin VB.OptionButton oAscendente 
            Caption         =   "Ascendente"
            Height          =   255
            Left            =   210
            TabIndex        =   14
            Top             =   435
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame fraOrden 
         Height          =   2355
         Left            =   165
         TabIndex        =   23
         Top             =   390
         Width           =   3660
         Begin VB.OptionButton Option0 
            Caption         =   "... por Fecha de Pago"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   225
            Value           =   -1  'True
            Width           =   2910
         End
         Begin VB.OptionButton Option4 
            Caption         =   "... por Estado"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1875
            Width           =   2910
         End
         Begin VB.OptionButton Option1 
            Caption         =   "... por Banco y Nro de Cheque"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   765
            Width           =   2910
         End
         Begin VB.OptionButton Option3 
            Caption         =   "... por Fecha de Emisión"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1320
            Width           =   2910
         End
      End
      Begin VB.Frame fraImpresion 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   165
         TabIndex        =   22
         Top             =   3765
         Width           =   8310
         Begin VB.CommandButton CmdCambiarImp 
            Caption         =   "&Configurar Impresora"
            Height          =   495
            Left            =   195
            TabIndex        =   36
            Top             =   600
            Width           =   1890
         End
         Begin VB.OptionButton oImpresora 
            Caption         =   "Impresora"
            Height          =   255
            Left            =   2220
            TabIndex        =   17
            Top             =   270
            Width           =   990
         End
         Begin VB.OptionButton oPantalla 
            Caption         =   "Pantalla"
            Height          =   255
            Left            =   1185
            TabIndex        =   16
            Top             =   270
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label LBImpActual 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impresora Actual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2235
            TabIndex        =   37
            Top             =   840
            Width           =   1440
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Destino:"
            Height          =   195
            Left            =   480
            TabIndex        =   35
            Top             =   270
            Width           =   585
         End
      End
      Begin Crystal.CrystalReport Rep 
         Left            =   4515
         Top             =   5085
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin MSComDlg.CommonDialog CDImpresora 
         Left            =   4950
         Top             =   5025
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   64
      End
   End
End
Attribute VB_Name = "FrmListChequesPropios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Limpio_Campos()
   Me.TxtFecVtoD.Text = ""
   Me.TxtFecVtoH.Text = ""
   Me.CboBanco.ListIndex = -1
   Me.cboCtaBancaria.ListIndex = -1
   Me.TxtNroCheque.Text = ""
   Me.TxtFecIngresoD.Text = ""
   Me.TxtFecIngresoH.Text = ""
   Me.CboEstado.ListIndex = -1
   Me.Fecha1.Text = ""
   Me.Fecha2.Text = ""
End Sub

Private Sub CboBanco_LostFocus()
    If CboBanco.ListIndex <> -1 Then
        Set Rec1 = New ADODB.Recordset
        cboCtaBancaria.Clear
        sql = "SELECT CTA_NROCTA FROM CTA_BANCARIA"
        sql = sql & " WHERE BAN_CODINT=" & XN(CboBanco.ItemData(CboBanco.ListIndex))
        sql = sql & " AND CTA_FECCIE IS NULL"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
         Do While Rec1.EOF = False
             cboCtaBancaria.AddItem Trim(Rec1!CTA_NROCTA)
             Rec1.MoveNext
         Loop
         cboCtaBancaria.ListIndex = 0
        End If
        Rec1.Close
    End If
End Sub

Private Sub CboEstado_LostFocus()
    If Me.Option4.Value = True Then Fecha1.SetFocus
     If Me.Option0.Value = True Then Me.cmdAgregar.SetFocus
End Sub

Private Sub CmdAgregar_Click()
    sql = ""
    'VALIDO LAS FECHAS
    If Option0.Value = True Then
        If TxtFecVtoD.Text = "" Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            TxtFecVtoD.SetFocus
            Exit Sub
        End If
    ElseIf Option3.Value = True Then
        If TxtFecIngresoD.Text = "" Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            TxtFecIngresoD.SetFocus
            Exit Sub
        End If
    ElseIf Option4.Value = True Then
        If Fecha1.Text = "" Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            Fecha1.SetFocus
            Exit Sub
        End If
    End If
   
   On Error GoTo ErrorTrans
   
   Screen.MousePointer = 11
   
   'Sentido del Orden
   If oAscendente.Value = True Then
      wSentido = "+"
      Rep.Formulas(1) = "sentido ='Sentido: ASCENDENTE'"
   Else
      wSentido = "-"
      Rep.Formulas(1) = "sentido ='Sentido: DESCENDENTE '"
   End If
   
   If Me.Option0.Value = True Then 'Por Fecha de Vencimiento
       
       If Me.TxtFecVtoD.Text = "" Or Me.TxtFecVtoH.Text = "" Then
          If Me.TxtFecVtoD.Text = "" Then
            Me.TxtFecVtoD.Text = Format(Date, "dd/mm/yyyy")
          ElseIf Me.TxtFecVtoH.Text = "" Then
            Me.TxtFecVtoH.Text = Format(Date, "dd/mm/yyyy")
          End If
       End If
       
       '{ChequePropioEstadoVigente.ECH_CODIGO} = 1 Unicamente Cheques en Cartera
        sql = sql & " {ChequePropioEstadoVigente.CHEP_FECVTO} >= DATE(" & Mid(TxtFecVtoD.Text, 7, 4) & "," & _
                                                            Mid(TxtFecVtoD.Text, 4, 2) & "," & _
                                                            Mid(TxtFecVtoD.Text, 1, 2) & ") and " & _
                      "{ChequePropioEstadoVigente.CHEP_FECVTO} <= DATE(" & Mid(TxtFecVtoH.Text, 7, 4) & "," & _
                                                                    Mid(TxtFecVtoH.Text, 4, 2) & "," & _
                                                                    Mid(TxtFecVtoH.Text, 1, 2) & ")"
       wCondicion = wSentido & " {ChequePropioEstadoVigente.CHEP_FECVTO}"
       wCondicion1 = wSentido & " {ChequePropioEstadoVigente.CHEP_NUMERO}"
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE PAGO. Y NRO DE CHEQUE'"
       
   ElseIf Me.Option1.Value = True Then 'por Banco y Nº de Cheque
       
        sql = sql & " {ChequePropioEstadoVigente.BAN_CODINT} =  " & XN(CboBanco.ItemData(CboBanco.ListIndex)) _
                 & " AND {ChequePropioEstadoVigente.CTA_NROCTA} = " & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
        If TxtNroCheque.Text <> "" Then
            sql = sql & " AND {ChequePropioEstadoVigente.CHEP_NUMERO} =  " & XS(TxtNroCheque.Text)
        End If
       
       wCondicion = wSentido & " {ChequePropioEstadoVigente.CHEP_NUMERO}"
       wCondicion1 = ""
       Rep.Formulas(0) = "orden ='Ordenado por: NÚMERO DE CHEQUE'"
          
   ElseIf Me.Option3.Value = True Then 'por Fecha de Ingreso
   
       If Me.TxtFecIngresoD.Text = "" Or Me.TxtFecIngresoH.Text = "" Then
          If Me.TxtFecIngresoD.Text = "" Then
            Me.TxtFecIngresoD.Text = Format(Date, "dd/mm/yyyy")
          ElseIf Me.TxtFecIngresoH.Text = "" Then
            Me.TxtFecIngresoH.Text = Format(Date, "dd/mm/yyyy")
          End If
       End If
       
       sql = sql & "{ChequePropioEstadoVigente.CHEP_FECENT} >= DATE(" & Mid(TxtFecIngresoD.Text, 7, 4) & _
                                                      "," & Mid(TxtFecIngresoD.Text, 4, 2) & _
                                                      "," & Mid(TxtFecIngresoD.Text, 1, 2) & ")and " & _
                   "{ChequePropioEstadoVigente.CHEP_FECENT} <= DATE(" & Mid(TxtFecIngresoH.Text, 7, 4) & "," & _
                                                            Mid(TxtFecIngresoH.Text, 4, 2) & "," & _
                                                            Mid(TxtFecIngresoH.Text, 1, 2) & ")"
       
       wCondicion = wSentido & " {ChequePropioEstadoVigente.CHEP_FECENT}"
       wCondicion1 = wSentido & " {ChequePropioEstadoVigente.CHEP_FECVTO}"
       
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE INGRESO y FECHA DE PAGO.'"
   
   ElseIf Me.Option4.Value = True Then 'por Estado y Fecha de Ingreso
   
       If Fecha1.Text = "" Or Fecha2.Text = "" Then
          If Fecha1.Text = "" Then
            Fecha1.Text = Format(Date, "dd/mm/yyyy")
          ElseIf Fecha2.Text = "" Then
            Fecha2.Text = Format(Date, "dd/mm/yyyy")
          End If
       End If
    
       sql = sql & " {ChequePropioEstadoVigente.CHEP_FECENT} >= DATE(" & Mid(Fecha1.Text, 7, 4) & "," & _
                                                                    Mid(Fecha1.Text, 4, 2) & "," & _
                                                                    Mid(Fecha1.Text, 1, 2) & ") and " & _
                   "{ChequePropioEstadoVigente.CHEP_FECENT} <= DATE(" & Mid(Fecha2.Text, 7, 4) & "," & _
                                                                    Mid(Fecha2.Text, 4, 2) & "," & _
                                                                    Mid(Fecha2.Text, 1, 2) & ")"
       'por Estado
       If Me.CboEstado.List(Me.CboEstado.ListIndex) <> "(Todos)" Then
           If Me.CboEstado.List(Me.CboEstado.ListIndex) = "RECHAZADOS TODOS" Then
              sql = sql & " AND {ChequePropioEstadoVigente.ECH_CODIGO} >= 8 " & _
                            " AND {ChequePropioEstadoVigente.ECH_CODIGO} <= 24 "
           Else
              sql = sql & " AND {ChequePropioEstadoVigente.ECH_CODIGO} =  " & XN(CboEstado.ItemData(CboEstado.ListIndex))
           End If
       End If
       wCondicion = wSentido & " {ChequePropioEstadoVigente.CHEP_FECVTO}"
       wCondicion1 = wSentido & " {ChequePropioEstadoVigente.CHEP_NUMERO}"
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE PAGO. Y NRO. DE CHEQUE'"
   
   End If
   
   If oImpresora = True Then
       Rep.Destination = 1
   Else
       Rep.Destination = 0
       Rep.WindowMinButton = 0
       Rep.WindowTitle = "Consulta de Cheques Propios"
       Rep.WindowBorderStyle = 2
   End If
   
   Rep.SortFields(0) = wCondicion
   Rep.SortFields(1) = wCondicion1
   
   Rep.SelectionFormula = sql
   Rep.WindowState = crptNormal
   Rep.WindowBorderStyle = crptNoBorder
   Rep.Connect = "Provider=MSDASQL.1;Persst Security Info=False;Data Source=SIHDG"
   
   Rep.ReportFileName = DRIVE & DirReport & "chequepropio.rpt"
   Rep.Action = 1
   
   Rep.Formulas(0) = ""
   Rep.Formulas(1) = ""
   Rep.Formulas(2) = ""
   Rep.Formulas(3) = ""
   
   Screen.MousePointer = 1
   Exit Sub

ErrorTrans:
  Screen.MousePointer = 1
  MsgBox "Error intentando armar el reporte. " & Chr(13) & Err.Description, 16, TIT_MSGBOX
End Sub

Private Sub CmdCambiarImp_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub CmdCancelar_Click()
    Limpio_Campos
    Option0.Value = True
    Option1.Value = False
    Option3.Value = False
    Option4.Value = False
    oAscendente.Value = True
    oPantalla.Value = True
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set FrmListChequesPropios = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then 'avanza de campo
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    KeyPreview = True
    
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.cboCtaBancaria.Enabled = False
    Me.CboEstado.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    
    Call Centrar_pantalla(Me)

    Set rec = New ADODB.Recordset
    
    CboEstado.Clear
    CboEstado.AddItem "(Todos)"
    sql = "SELECT ECH_CODIGO, ECH_DESCRI FROM ESTADO_CHEQUE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While Not rec.EOF
            CboEstado.AddItem Trim(rec!ECH_DESCRI)
            CboEstado.ItemData(CboEstado.NewIndex) = Trim(rec!ECH_CODIGO)
            rec.MoveNext
        Loop
        Me.CboEstado.ListIndex = -1
    End If
    rec.Close
    CboEstado.AddItem "RECHAZADOS TODOS"
    Me.MousePointer = 1
    
    LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
    
    Option0_Click
End Sub

Private Sub oImpresora_Click()
  Me.LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
  Me.LBImpActual.Visible = True
End Sub

Private Sub oPantalla_Click()
 ' Me.CDImpresora.Visible = False
  Me.LBImpActual.Visible = False
End Sub

Private Sub Option0_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = True
    Me.TxtFecVtoH.Enabled = True
    Me.CboBanco.Enabled = False
    Me.cboCtaBancaria.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.CboEstado.Enabled = False
    Me.Fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    If Me.TxtFecVtoD.Visible = True Then Me.TxtFecVtoD.SetFocus
End Sub

Private Sub Option1_Click()
    Me.CboBanco.Clear
    Set rec = New ADODB.Recordset
    sql = "SELECT B.BAN_DESCRI, B.BAN_CODINT"
    sql = sql & " FROM BANCO B, CTA_BANCARIA CB"
    sql = sql & " WHERE B.BAN_CODINT=CB.BAN_CODINT"
    sql = sql & " ORDER BY B.BAN_DESCRI"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        rec.MoveFirst
        Do While Not rec.EOF
            Me.CboBanco.AddItem Trim(rec!BAN_DESCRI)
            Me.CboBanco.ItemData(Me.CboBanco.NewIndex) = rec!BAN_CODINT
            rec.MoveNext
        Loop
        Me.CboBanco.ListIndex = 0
    End If
    rec.Close
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.ListIndex = 0
    Me.CboBanco.Enabled = True
    Me.cboCtaBancaria.Enabled = True
    Me.TxtNroCheque.Enabled = True
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.CboEstado.Enabled = False
    Me.Fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.CboBanco.SetFocus
End Sub

Private Sub Option3_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.cboCtaBancaria.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = True
    Me.TxtFecIngresoH.Enabled = True
    Me.CboEstado.Enabled = False
    Me.Fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.TxtFecIngresoD.SetFocus
End Sub

Private Sub Option4_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.cboCtaBancaria.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.CboEstado.ListIndex = 0
    Me.CboEstado.Enabled = True
    Me.Fecha1.Enabled = True
    Me.Fecha2.Enabled = True
    Me.CboEstado.SetFocus
End Sub

Private Sub TxtFecIngresoD_LostFocus()
   'If Me.Option3.Value = True And TxtFecIngresoD.Text = "" Then TxtFecIngresoD.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub TxtFecIngresoH_LostFocus()
'If Me.Option3.Value = True And TxtFecIngresoH.Text = "" Then TxtFecIngresoH.Text = Format(Date, "dd/mm/yyyy")

  If IsDate(TxtFecIngresoD.Text) And IsDate(TxtFecIngresoH.Text) Then
    
    If CVDate(TxtFecIngresoD.Text) > CVDate(TxtFecIngresoH.Text) Then
      MsgBox "La Fecha Hasta no puede ser inferior a la Fecha Desde. Verifique!", 16, TIT_MSGBOX
      TxtFecIngresoH.Text = ""
      TxtFecIngresoH.SetFocus
      Exit Sub
    Else
      If Not IsDate(TxtFecIngresoD.Text) Then TxtFecIngresoD.Text = ""
      If Not IsDate(TxtFecIngresoH.Text) Then TxtFecIngresoH.Text = ""
    End If
    
 End If
End Sub

Private Sub TxtFecVtoD_LostFocus()
  'If Me.Option0.Value = True And TxtFecVtoD.Text = "" Then TxtFecVtoD.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub TxtFecVtoH_LostFocus()

  'If Me.Option0.Value = True And TxtFecVtoH.Text = "" Then TxtFecVtoH.Text = Format(Date, "dd/mm/yyyy")
  
  If IsDate(TxtFecVtoD.Text) And IsDate(TxtFecVtoH.Text) Then
  
    If CVDate(TxtFecVtoD.Text) > CVDate(TxtFecVtoH.Text) Then
      MsgBox "La Fecha Hasta no puede ser inferior a la Fecha Desde. Verifique!", 16, TIT_MSGBOX
      TxtFecVtoH.Text = ""
      TxtFecVtoD.SetFocus
      Exit Sub
    Else
      cmdAgregar.SetFocus
    End If
 End If
End Sub

Private Sub TxtNroCheque_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtNroCheque_LostFocus()
    If TxtNroCheque.Text <> "" Then
        If Len(TxtNroCheque.Text) < 10 Then TxtNroCheque.Text = CompletarConCeros(TxtNroCheque.Text, 10)
    End If
    cmdAgregar.SetFocus
End Sub
