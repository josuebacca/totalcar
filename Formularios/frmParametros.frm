VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Impuesto Transacciones Financieras (Impuesto al Cheque)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   30
      TabIndex        =   43
      Top             =   4020
      Width           =   9630
      Begin VB.CheckBox chkAplicoImpuesto 
         Caption         =   "Aplicar Impuesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2340
         TabIndex        =   46
         Top             =   375
         Width           =   1815
      End
      Begin VB.TextBox txtImpuestoCheque 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   705
         MaxLength       =   5
         TabIndex        =   45
         Top             =   345
         Width           =   825
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   1590
         TabIndex        =   47
         Top             =   420
         Width           =   120
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   225
         TabIndex        =   44
         Top             =   405
         Width           =   360
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Numeración Comprobantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   4590
      TabIndex        =   24
      Top             =   1470
      Width           =   5070
      Begin VB.TextBox txtremitoM 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3885
         TabIndex        =   30
         Top             =   420
         Width           =   1080
      End
      Begin VB.TextBox txtNroReciboB 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3885
         TabIndex        =   54
         Top             =   1470
         Width           =   1080
      End
      Begin VB.TextBox txtNroReciboA 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3885
         TabIndex        =   53
         Top             =   1125
         Width           =   1080
      End
      Begin VB.TextBox txtSalidaDeposito 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3885
         TabIndex        =   52
         Top             =   795
         Width           =   1080
      End
      Begin VB.TextBox txtRecepcionMerca 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1470
         TabIndex        =   49
         Top             =   420
         Width           =   1080
      End
      Begin VB.TextBox txtNotaDebitoB 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3885
         TabIndex        =   32
         Top             =   1815
         Width           =   1080
      End
      Begin VB.TextBox txtNotaDebitoA 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3885
         TabIndex        =   31
         Top             =   2160
         Width           =   1080
      End
      Begin VB.TextBox txtNotaCreditoB 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1470
         TabIndex        =   29
         Top             =   2145
         Width           =   1080
      End
      Begin VB.TextBox txtNotaCreditoA 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1470
         TabIndex        =   28
         Top             =   1800
         Width           =   1080
      End
      Begin VB.TextBox txtNroFacturaA 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1470
         TabIndex        =   27
         Top             =   1110
         Width           =   1080
      End
      Begin VB.TextBox txtNroRemito 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1470
         TabIndex        =   26
         Top             =   765
         Width           =   1080
      End
      Begin VB.TextBox txtNroFacturaB 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1470
         TabIndex        =   25
         Top             =   1455
         Width           =   1080
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Remito Multiple:"
         Height          =   195
         Left            =   2685
         TabIndex        =   51
         Top             =   465
         Width           =   1125
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Entrada Depósito:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   465
         Width           =   1275
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Salida Depósito:"
         Height          =   195
         Left            =   2655
         TabIndex        =   42
         Top             =   825
         Width           =   1155
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Recibo A:"
         Height          =   195
         Left            =   2760
         TabIndex        =   41
         Top             =   1170
         Width           =   1050
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Recibo B:"
         Height          =   195
         Left            =   2760
         TabIndex        =   40
         Top             =   1515
         Width           =   1050
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nota Débito B:"
         Height          =   195
         Left            =   2760
         TabIndex        =   39
         Top             =   1845
         Width           =   1050
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nota Débito A:"
         Height          =   195
         Left            =   2760
         TabIndex        =   38
         Top             =   2190
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nota Crédito B:"
         Height          =   195
         Left            =   315
         TabIndex        =   37
         Top             =   2190
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nota Crédito A:"
         Height          =   195
         Left            =   315
         TabIndex        =   36
         Top             =   1845
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Factura A:"
         Height          =   195
         Left            =   315
         TabIndex        =   35
         Top             =   1155
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Remito:"
         Height          =   195
         Left            =   510
         TabIndex        =   34
         Top             =   810
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Factura B:"
         Height          =   195
         Left            =   315
         TabIndex        =   33
         Top             =   1500
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Condición Impositiva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   30
      TabIndex        =   14
      Top             =   1470
      Width           =   4560
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1290
         TabIndex        =   18
         Top             =   1785
         Width           =   1080
      End
      Begin VB.ComboBox cboIva 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1080
         Width           =   3150
      End
      Begin VB.TextBox txtIngBrutos 
         Height          =   315
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   15
         Top             =   360
         Width           =   1350
      End
      Begin MSComCtl2.DTPicker fechaInicio 
         Height          =   315
         Left            =   1290
         TabIndex        =   48
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   56950785
         CurrentDate     =   41098
      End
      Begin MSMask.MaskEdBox Cuit1 
         Height          =   315
         Left            =   1290
         TabIndex        =   16
         Top             =   720
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
         Caption         =   "I.V.A.:"
         Height          =   195
         Left            =   780
         TabIndex        =   23
         Top             =   1830
         Width           =   435
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Inicio Actividad:"
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "IVA Condición:"
         Height          =   195
         Left            =   165
         TabIndex        =   21
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ing. Brutos:"
         Height          =   210
         Left            =   420
         TabIndex        =   20
         Top             =   405
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "C.U.I.T.:"
         Height          =   195
         Left            =   615
         TabIndex        =   19
         Top             =   765
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Generales..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   9615
      Begin VB.TextBox txtTelefono 
         Height          =   315
         Left            =   1215
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1005
         Width           =   2070
      End
      Begin VB.TextBox txtDireccion 
         Height          =   315
         Left            =   1215
         MaxLength       =   50
         TabIndex        =   7
         Top             =   645
         Width           =   4860
      End
      Begin VB.TextBox txtRazSoc 
         Height          =   315
         Left            =   1215
         MaxLength       =   50
         TabIndex        =   6
         Top             =   315
         Width           =   4860
      End
      Begin VB.TextBox txtNroRepresentada 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7590
         TabIndex        =   5
         Top             =   660
         Width           =   1080
      End
      Begin VB.TextBox txtSucursal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7590
         TabIndex        =   4
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         Height          =   195
         Left            =   465
         TabIndex        =   13
         Top             =   1065
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   420
         TabIndex        =   12
         Top             =   705
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social:"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Representada Nro:"
         Height          =   195
         Left            =   6165
         TabIndex        =   10
         Top             =   705
         Width           =   1350
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal:"
         Height          =   195
         Left            =   6855
         TabIndex        =   9
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   750
      Left            =   7785
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4935
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   8670
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4935
      Width           =   870
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
      TabIndex        =   2
      Top             =   5130
      Width           =   750
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGrabar_Click()
    If Validar_Parametros = False Then Exit Sub
    
    If MsgBox("¿Confirma los valores de Parámetros?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayError
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Actualizando..."
    
    DBConn.BeginTrans
    sql = "UPDATE PARAMETROS"
    sql = sql & " SET RAZ_SOCIAL=" & XS(txtRazSoc.Text)
    sql = sql & " ,DIRECCION=" & XS(txtDireccion.Text)
    sql = sql & " ,TELEFONO=" & XS(txtTelefono.Text)
    sql = sql & " ,CUIT=" & XS(Cuit1.Text)
    sql = sql & " ,ING_BRUTOS=" & XS(txtIngBrutos.Text)
    sql = sql & " ,IVA_CODIGO=" & cboIva.ItemData(cboIva.ListIndex)
    sql = sql & " ,INICIO_ACTIVIDAD=" & XDQ(fechaInicio.Value)
    sql = sql & " ,NRO_REMITO=" & XN(txtNroRemito.Text)
    sql = sql & " ,FACTURA_A=" & XN(txtNroFacturaA.Text)
    sql = sql & " ,FACTURA_B=" & XN(txtNroFacturaB.Text)
    sql = sql & " ,NOTA_CREDITO_A=" & XN(txtNotaCreditoA.Text)
    sql = sql & " ,NOTA_CREDITO_B=" & XN(txtNotaCreditoB.Text)
    sql = sql & " ,NOTA_DEBITO_A=" & XN(txtNotaDebitoA.Text)
    sql = sql & " ,NOTA_DEBITO_B=" & XN(txtNotaDebitoB.Text)
    sql = sql & " ,IVA=" & XN(txtIva.Text)
    sql = sql & " ,SUCURSAL=" & XN(txtSucursal.Text)
    sql = sql & " ,RECIBO_A=" & XN(txtNroReciboA.Text)
    sql = sql & " ,RECIBO_B=" & XN(txtNroReciboB.Text)
    sql = sql & " ,RECEPCION_MERCADERIA=" & XN(txtRecepcionMerca.Text)
    sql = sql & " ,SALIDA_MERCADERIA=" & XN(txtSalidaDeposito.Text)
    sql = sql & " ,VALOR_IMPUESTO=" & XN(txtImpuestoCheque.Text)
    sql = sql & " ,REMITOM=" & XN(txtremitoM.Text)
    If chkAplicoImpuesto.Value = Checked Then
        sql = sql & " ,APLICA_IMPUESTO=" & XS("S")
    Else
        sql = sql & " ,APLICA_IMPUESTO=" & XS("N")
    End If
    
    DBConn.Execute sql
    DBConn.CommitTrans
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox "Los cambios se registraron correctamente", vbInformation, TIT_MSGBOX
    Exit Sub
    
HayError:
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function Validar_Parametros() As Boolean
    If txtSucursal.Text = "" Then
        MsgBox "Debe Ingresar el número de Sucursal", vbExclamation, TIT_MSGBOX
        txtSucursal.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtRecepcionMerca.Text = "" Then
        MsgBox "Debe Ingresar el número de Recepción de Mercadería", vbExclamation, TIT_MSGBOX
        txtRecepcionMerca.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNroRemito.Text = "" Then
        MsgBox "Debe Ingresar el número de Remito", vbExclamation, TIT_MSGBOX
        txtNroRemito.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNroFacturaA.Text = "" Then
        MsgBox "Debe Ingresar el número de Factura A", vbExclamation, TIT_MSGBOX
        txtNroFacturaA.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNroFacturaB.Text = "" Then
        MsgBox "Debe Ingresar el número de Factura B", vbExclamation, TIT_MSGBOX
        txtNroFacturaB.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNotaCreditoA.Text = "" Then
        MsgBox "Debe Ingresar el número de Nota de Crédito A", vbExclamation, TIT_MSGBOX
        txtNotaCreditoA.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNotaCreditoB.Text = "" Then
        MsgBox "Debe Ingresar el número de Nota de Crédito B", vbExclamation, TIT_MSGBOX
        txtNotaCreditoB.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNotaDebitoA.Text = "" Then
        MsgBox "Debe Ingresar el número de Nota de Débito A", vbExclamation, TIT_MSGBOX
        txtNotaDebitoA.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNotaDebitoB.Text = "" Then
        MsgBox "Debe Ingresar el número de Nota de Débito B", vbExclamation, TIT_MSGBOX
        txtNotaDebitoB.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
      Validar_Parametros = True
End Function

Private Sub CmdSalir_Click()
    Set frmParametros = Nothing
    Unload Me
End Sub

Private Sub CUIT1_LostFocus()
    If Cuit1.Text <> "" Then
        If ValidoCuit(Cuit1.Text) = False Then
         Cuit1.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    'cargo combo iva
    LlenarComboIva
    'busco datos
    BuscarDatos
 
    lblEstado.Caption = ""
End Sub

Private Sub BuscarDatos()
    sql = "SELECT * FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtRazSoc.Text = IIf(IsNull(rec!RAZ_SOCIAL), "", rec!RAZ_SOCIAL)
        txtDireccion.Text = IIf(IsNull(rec!DIRECCION), "", rec!DIRECCION)
        txtTelefono.Text = IIf(IsNull(rec!TELEFONO), "", rec!TELEFONO)
        Cuit1.Text = IIf(IsNull(rec!cuit), "", rec!cuit)
        txtIngBrutos.Text = IIf(IsNull(rec!ING_BRUTOS), "", rec!ING_BRUTOS)
        Call BuscaCodigoProxItemData(IIf(IsNull(rec!IVA_CODIGO), 1, rec!IVA_CODIGO), cboIva)
        fechaInicio.Value = IIf(IsNull(rec!INICIO_ACTIVIDAD), "", rec!INICIO_ACTIVIDAD)
        txtIva.Text = IIf(IsNull(rec!IVA), "", rec!IVA)
        txtSucursal.Text = IIf(IsNull(rec!Sucursal), 1, rec!Sucursal)
        txtNroRemito.Text = IIf(IsNull(rec!NRO_REMITO), 1, rec!NRO_REMITO)
        txtNroFacturaA.Text = IIf(IsNull(rec!FACTURA_A), 1, rec!FACTURA_A)
        txtNroFacturaB.Text = IIf(IsNull(rec!FACTURA_B), 1, rec!FACTURA_B)
        txtNotaCreditoA.Text = IIf(IsNull(rec!NOTA_CREDITO_A), 1, rec!NOTA_CREDITO_A)
        txtNotaCreditoB.Text = IIf(IsNull(rec!NOTA_CREDITO_B), 1, rec!NOTA_CREDITO_B)
        txtNotaDebitoA.Text = IIf(IsNull(rec!NOTA_DEBITO_A), 1, rec!NOTA_DEBITO_A)
        txtNotaDebitoB.Text = IIf(IsNull(rec!NOTA_DEBITO_B), 1, rec!NOTA_DEBITO_B)
        txtNroReciboA.Text = IIf(IsNull(rec!RECIBO_A), 1, rec!RECIBO_A)
        txtNroReciboB.Text = IIf(IsNull(rec!RECIBO_B), 1, rec!RECIBO_B)
        txtRecepcionMerca.Text = IIf(IsNull(rec!RECEPCION_MERCADERIA), 1, rec!RECEPCION_MERCADERIA)
        txtSalidaDeposito.Text = IIf(IsNull(rec!SALIDA_MERCADERIA), 1, rec!SALIDA_MERCADERIA)
        txtImpuestoCheque.Text = IIf(IsNull(rec!VALOR_IMPUESTO), 0, rec!VALOR_IMPUESTO)
        txtremitoM.Text = IIf(IsNull(rec!REMITOM), 1, rec!REMITOM)
        If rec!APLICA_IMPUESTO = "S" Then
            chkAplicoImpuesto.Value = Checked
        Else
            chkAplicoImpuesto.Value = Unchecked
        End If
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


Private Sub txtDireccion_GotFocus()
    SelecTexto txtDireccion
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtImpuestoCheque_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImpuestoCheque, KeyAscii)
End Sub

Private Sub txtIngBrutos_GotFocus()
    SelecTexto txtIngBrutos
End Sub

Private Sub txtIngBrutos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtIva_GotFocus()
   SelecTexto txtIva
End Sub

Private Sub txtIva_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroDecimal(txtIva, KeyAscii)
End Sub

Private Sub txtIva_LostFocus()
    If txtIva.Text <> "" Then
        If ValidarPorcentaje(txtIva) = False Then txtIva.SetFocus
    End If
End Sub

Private Sub txtNotaCreditoA_GotFocus()
    SelecTexto txtNotaCreditoA
End Sub

Private Sub txtNotaCreditoA_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNotaCreditoB_GotFocus()
    SelecTexto txtNotaCreditoB
End Sub

Private Sub txtNotaCreditoB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNotaDebitoA_GotFocus()
    SelecTexto txtNotaDebitoA
End Sub

Private Sub txtNotaDebitoA_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNotaDebitoB_GotFocus()
    SelecTexto txtNotaDebitoB
End Sub

Private Sub txtNotaDebitoB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroFacturaA_GotFocus()
    SelecTexto txtNroFacturaA
End Sub

Private Sub txtNroFacturaA_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroFacturaB_GotFocus()
    SelecTexto txtNroFacturaB
End Sub

Private Sub txtNroFacturaB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroReciboA_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroReciboB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroRemito_GotFocus()
    SelecTexto txtNroRemito
End Sub

Private Sub txtNroRemito_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
Private Sub txtRazSoc_GotFocus()
    SelecTexto txtRazSoc
End Sub

Private Sub txtRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtRecepcionMerca_GotFocus()
    SelecTexto txtRecepcionMerca
End Sub

Private Sub txtRecepcionMerca_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
