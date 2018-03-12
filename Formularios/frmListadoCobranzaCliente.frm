VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListadoCobranzaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobranza a Clientes"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Ver..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   17
      Top             =   1455
      Width           =   6705
      Begin VB.OptionButton optVieja 
         Caption         =   "Cobranza Anterior"
         Height          =   195
         Left            =   2955
         TabIndex        =   4
         Top             =   240
         Width           =   1620
      End
      Begin VB.OptionButton optNuevo 
         Caption         =   "Cobranza Nueva"
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   1560
      End
   End
   Begin VB.Frame FrameImpresora 
      Caption         =   "impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   60
      TabIndex        =   13
      Top             =   2085
      Width           =   6690
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   6
         Top             =   330
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   5
         Top             =   330
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   330
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   4140
      Picture         =   "frmListadoCobranzaCliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2925
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   5895
      Picture         =   "frmListadoCobranzaCliente.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2925
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoCobranzaCliente.frx":0BD4
      Height          =   750
      Left            =   5010
      Picture         =   "frmListadoCobranzaCliente.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2925
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listar por..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   60
      TabIndex        =   11
      Top             =   -15
      Width           =   6690
      Begin VB.TextBox txtNroCobranza 
         Height          =   315
         Left            =   1785
         MaxLength       =   40
         TabIndex        =   2
         Top             =   915
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscarCliente 
         Height          =   315
         Left            =   1800
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoCobranzaCliente.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Buscar Cliente"
         Top             =   285
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.TextBox txtDesCli 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2235
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Descripción"
         Top             =   285
         Width           =   4305
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   795
         MaxLength       =   40
         TabIndex        =   0
         Top             =   285
         Width           =   975
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   15
         Top             =   330
         Width           =   525
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cobranza:"
         Height          =   195
         Left            =   630
         TabIndex        =   12
         Top             =   945
         Width           =   1020
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3330
      Top             =   3030
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2835
      Top             =   2940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   345
      Left            =   150
      TabIndex        =   16
      Top             =   3105
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoCobranzaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim NumeroCobranza As String

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCliente.Text = frmBuscar.grdBuscar.Text
        txtCliente_LostFocus
        txtDesCli.SetFocus
    Else
        txtCliente.SetFocus
    End If
End Sub

Private Sub cmdListar_Click()
    If optVieja.Value = True Then
        If txtNroCobranza.Text = "" Then
            MsgBox "Debe ingresar el número de Cobranza", vbExclamation, TIT_MSGBOX
            txtNroCobranza.SetFocus
            Exit Sub
        End If
    End If
    
    On Error GoTo AlCarajo
    DBConn.BeginTrans
    
    lblEstado.Caption = "Buscando Listado..."
    Screen.MousePointer = vbHourglass
    
    sql = "DELETE FROM TMP_RECIBO_CLIENTE"
    DBConn.Execute sql
    

    Call ReciboCobroComprobante(txtCliente)
    Call ReciboCobroCheques(txtCliente)
    Call ReciboCobroMoneda(txtCliente)
    Call ReciboCobroFacturas(txtCliente)
    If optNuevo.Value = True Then
        Call ReciboCobroActualizo(txtCliente)
    End If
'    End If
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    
    If txtCliente.Text <> "" Then
        Rep.SelectionFormula = "{TMP_RECIBO_CLIENTE.CLI_CODIGO}=" & txtCliente.Text
    End If
    
    Rep.Formulas(0) = "COBRANZA='" & "COBRANZA NRO: " & txtNroCobranza.Text & "'"
    
    Rep.WindowTitle = "Listado de Cobranza a Cliente"
    Rep.ReportFileName = DRIVE & DirReport & "rptcobranzaclientes.rpt"

    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     DBConn.CommitTrans
     
     Rep.Action = 1
     
     Screen.MousePointer = vbNormal
     lblEstado.Caption = ""
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     Exit Sub
     
AlCarajo:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    If rec.State = 1 Then rec.Close
    If Rec1.State = 1 Then Rec1.Close
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub BuscoNroCobranza()
    NumeroCobranza = ""
    sql = "SELECT MAX(REC_LISTADO) + 1 AS NUMERO_COBRANZA FROM RECIBO_CLIENTE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If IsNull(rec!NUMERO_COBRANZA) Then
        NumeroCobranza = 1
    Else
        NumeroCobranza = (rec!NUMERO_COBRANZA)
    End If
    rec.Close
End Sub

Private Sub ReciboCobroActualizo(CliCodigo As String)
    BuscoNroCobranza
    txtNroCobranza.Text = NumeroCobranza
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT TCO_CODIGO,REC_NUMERO,REC_SUCURSAL,REC_FECHA"
    sql = sql & " FROM RECIBO_CLIENTE "
    sql = sql & " WHERE"
    sql = sql & " EST_CODIGO=3" 'ESTADO DEFINITIVO
    
    If CliCodigo <> "" Then sql = sql & " AND CLI_CODIGO=" & XN(CliCodigo)
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            sql = "UPDATE RECIBO_CLIENTE"
            sql = sql & " SET REC_LISTADO=" & XS(NumeroCobranza)
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & XN(Rec1!TCO_CODIGO)
            sql = sql & " AND REC_NUMERO=" & XN(Rec1!REC_NUMERO)
            sql = sql & " AND REC_SUCURSAL=" & XN(Rec1!REC_SUCURSAL)
            sql = sql & " AND REC_FECHA=" & XDQ(Rec1!REC_FECHA)

            DBConn.Execute sql

            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub ReciboCobroComprobante(CliCodigo As String)
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, RC.CLI_CODIGO,RC.REC_TOTAL"
    sql = sql & ",TC.TCO_ABREVIA, DR.DRE_COMFECHA, DR.DRE_COMNUMERO, DR.DRE_COMSUCURSAL,"
    sql = sql & " DR.DRE_COMIMP, RC.TCO_CODIGO,RC.REC_NUMERO,RC.REC_SUCURSAL,RC.REC_FECHA,"
    sql = sql & " RC.REC_LISTADO"
    sql = sql & " FROM CLIENTE C, DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE RC"
    sql = sql & " ,TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " RC.TCO_CODIGO=DR.TCO_CODIGO"
    sql = sql & " AND DR.DRE_TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND RC.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND RC.REC_SUCURSAL=DR.REC_SUCURSAL"
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND RC.EST_CODIGO=3" 'ESTADO DEFINITIVO
    
    If optNuevo.Value = True Then
        sql = sql & " AND RC.REC_LISTADO IS NULL"
    Else
        sql = sql & " AND RC.REC_LISTADO=" & XN(txtNroCobranza)
    End If
    If CliCodigo <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(CliCodigo)
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False

            sql = "INSERT INTO TMP_RECIBO_CLIENTE ("
            sql = sql & "CLI_CODIGO,CLI_RAZSOC,CLI_DOMICI,"
            sql = sql & "TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_SUCURSAL,COM_IMPORTE"
            'sql = sql & "FAC_ABREVIA,FAC_NUMERO,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL"
            sql = sql & ",REC_TOTAL,TCO_CODIGO,REC_NUMERO,REC_SUCURSAL,REC_FECHA)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec1!CLI_CODIGO) & ","
            sql = sql & XS(Rec1!CLI_RAZSOC) & ","
            sql = sql & XS(Rec1!CLI_DOMICI) & ","
            sql = sql & XS(Rec1!TCO_ABREVIA) & ","
            sql = sql & XDQ(Rec1!DRE_COMFECHA) & ","
            sql = sql & XS(Format(Rec1!DRE_COMNUMERO, "00000000")) & ","
            sql = sql & XS(Format(Rec1!DRE_COMSUCURSAL, "0000")) & ","
            sql = sql & XN(Rec1!DRE_COMIMP) & ","
            'sql = sql & " NULL,NULL,NULL,NULL,NULL,"
            sql = sql & XN(Rec1!REC_TOTAL) & ","
            sql = sql & XN(Rec1!TCO_CODIGO) & ","
            sql = sql & XS(Format(Rec1!REC_NUMERO, "00000000")) & ","
            sql = sql & XS(Format(Rec1!REC_SUCURSAL, "0000")) & ","
            sql = sql & XDQ(Rec1!REC_FECHA) & ")"
            DBConn.Execute sql

            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub ReciboCobroCheques(CliCodigo As String)

    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, RC.CLI_CODIGO,"
    sql = sql & "B.BAN_NOMCOR, CH.CHE_FECVTO ,DR.CHE_NUMERO, CH.CHE_IMPORT,RC.REC_TOTAL,"
    sql = sql & "RC.TCO_CODIGO,RC.REC_NUMERO,RC.REC_SUCURSAL,RC.REC_FECHA,"
    sql = sql & " RC.REC_LISTADO"
    sql = sql & " FROM CLIENTE C, DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE RC"
    sql = sql & " ,CHEQUE CH, BANCO B"
    sql = sql & " WHERE"
    sql = sql & " RC.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND RC.REC_SUCURSAL=DR.REC_SUCURSAL"
    sql = sql & " AND RC.TCO_CODIGO=DR.TCO_CODIGO"
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND DR.BAN_CODINT=CH.BAN_CODINT"
    sql = sql & " AND DR.CHE_NUMERO=CH.CHE_NUMERO"
    sql = sql & " AND CH.BAN_CODINT=B.BAN_CODINT"
    sql = sql & " AND RC.EST_CODIGO=3" 'ESTADO DEFINITIVO
    
    If optNuevo.Value = True Then
        sql = sql & " AND REC_LISTADO IS NULL"
    Else
        sql = sql & " AND RC.REC_LISTADO=" & XN(txtNroCobranza)
    End If
    If CliCodigo <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(CliCodigo)
    
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_RECIBO_CLIENTE ("
            sql = sql & "CLI_CODIGO,CLI_RAZSOC,CLI_DOMICI,"
            sql = sql & "TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE"
            'sql = sql & "FAC_ABREVIA,FAC_NUMERO,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL"
            sql = sql & ",REC_TOTAL,TCO_CODIGO,REC_NUMERO,REC_SUCURSAL,REC_FECHA)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec1!CLI_CODIGO) & ","
            sql = sql & XS(Rec1!CLI_RAZSOC) & ","
            sql = sql & XS(Rec1!CLI_DOMICI) & ","
            sql = sql & XS(Rec1!BAN_NOMCOR) & ","
            sql = sql & XDQ(Rec1!CHE_FECVTO) & ","
            sql = sql & XS(Rec1!CHE_NUMERO) & ","
            sql = sql & XN(Rec1!che_import) & ","
            'sql = sql & " NULL,NULL,NULL,NULL,NULL,"
            sql = sql & XN(Rec1!REC_TOTAL) & ","
            sql = sql & XN(Rec1!TCO_CODIGO) & ","
            sql = sql & XS(Format(Rec1!REC_NUMERO, "00000000")) & ","
            sql = sql & XS(Format(Rec1!REC_SUCURSAL, "0000")) & ","
            sql = sql & XDQ(Rec1!REC_FECHA) & ")"
            DBConn.Execute sql
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub ReciboCobroMoneda(CliCodigo As String)
    Set Rec1 = New ADODB.Recordset
    
    sql = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, RC.CLI_CODIGO,"
    sql = sql & "M.MON_DESCRI, DR.DRE_MONIMP, RC.REC_TOTAL,"
    sql = sql & "RC.TCO_CODIGO,RC.REC_NUMERO,RC.REC_SUCURSAL,RC.REC_FECHA,"
    sql = sql & " RC.REC_LISTADO"
    sql = sql & " FROM CLIENTE C, DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE RC"
    sql = sql & " , MONEDA M"
    sql = sql & " WHERE"
    sql = sql & " RC.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND RC.REC_SUCURSAL=DR.REC_SUCURSAL"
    sql = sql & " AND RC.TCO_CODIGO=DR.TCO_CODIGO"
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND DR.MON_CODIGO=M.MON_CODIGO"
    sql = sql & " AND RC.EST_CODIGO=3" 'ESTADO DEFINITIVO
    If optNuevo.Value = True Then
        sql = sql & " AND REC_LISTADO IS NULL"
    Else
        sql = sql & " AND RC.REC_LISTADO=" & XN(txtNroCobranza)
    End If
    If CliCodigo <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(CliCodigo)
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_RECIBO_CLIENTE ("
            sql = sql & "CLI_CODIGO,CLI_RAZSOC,CLI_DOMICI,"
            sql = sql & "TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE"
            'sql = sql & "FAC_ABREVIA,FAC_NUMERO,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL"
            sql = sql & ",REC_TOTAL,TCO_CODIGO,REC_NUMERO,REC_SUCURSAL,REC_FECHA)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec1!CLI_CODIGO) & ","
            sql = sql & XS(Rec1!CLI_RAZSOC) & ","
            sql = sql & XS(Rec1!CLI_DOMICI) & ","
            sql = sql & XS(Rec1!MON_DESCRI) & ","
            sql = sql & "NULL,"
            sql = sql & "NULL,"
            sql = sql & XN(Rec1!DRE_MONIMP) & ","
            'sql = sql & " NULL,NULL,NULL,NULL,NULL,"
            sql = sql & XN(Rec1!REC_TOTAL) & ","
            sql = sql & XN(Rec1!TCO_CODIGO) & ","
            sql = sql & XS(Format(Rec1!REC_NUMERO, "00000000")) & ","
            sql = sql & XS(Format(Rec1!REC_SUCURSAL, "0000")) & ","
            sql = sql & XDQ(Rec1!REC_FECHA) & ")"
            DBConn.Execute sql
            
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub ReciboCobroFacturas(CliCodigo As String)
    Set Rec1 = New ADODB.Recordset
    
    sql = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, RC.CLI_CODIGO,RC.REC_TOTAL,"
    sql = sql & "TC.TCO_ABREVIA, FR.FCL_NUMERO,FR.FCL_SUCURSAL, FR.FCL_FECHA ,F.FCL_TOTAL,FR.REC_IMPORTE,"
    sql = sql & "RC.TCO_CODIGO,RC.REC_NUMERO,RC.REC_SUCURSAL,RC.REC_FECHA,"
    sql = sql & " RC.REC_LISTADO"
    sql = sql & " FROM CLIENTE C, RECIBO_CLIENTE RC, FACTURAS_RECIBO_CLIENTE FR"
    sql = sql & ", TIPO_COMPROBANTE TC, FACTURA_CLIENTE F"
    sql = sql & " WHERE"
    sql = sql & " RC.REC_NUMERO=FR.REC_NUMERO"
    sql = sql & " AND RC.REC_SUCURSAL=FR.REC_SUCURSAL"
    sql = sql & " AND RC.TCO_CODIGO=FR.TCO_CODIGO"
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND FR.FCL_TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND FR.FCL_TCO_CODIGO=F.TCO_CODIGO"
    sql = sql & " AND FR.FCL_NUMERO=F.FCL_NUMERO"
    sql = sql & " AND FR.FCL_SUCURSAL=F.FCL_SUCURSAL"
    sql = sql & " AND RC.EST_CODIGO=3" 'ESTADO DEFINITIVO
    
    If optNuevo.Value = True Then
        sql = sql & " AND REC_LISTADO IS NULL"
    Else
        sql = sql & " AND RC.REC_LISTADO=" & XN(txtNroCobranza)
    End If
    If CliCodigo <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(CliCodigo)
    'sql = sql & " GROUP BY FR.FCL_NUMERO, FR.FCL_FECHA,C.CLI_RAZSOC, C.CLI_DOMICI, RC.CLI_CODIGO, RC.REP_CODIGO,RC.REC_TOTAL,"
    'sql = sql & "TC.TCO_ABREVIA,F.FCL_TOTAL,FR.REC_IMPORTE,RC.TCO_CODIGO,RC.REC_NUMERO,RC.REC_FECHA"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
        
            sql = "INSERT INTO TMP_RECIBO_CLIENTE ("
            sql = sql & "CLI_CODIGO,CLI_RAZSOC,CLI_DOMICI,"
            sql = sql & "TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            sql = sql & "FAC_ABREVIA,FAC_NUMERO,FAC_SUCURSAL,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL"
            sql = sql & ",REC_TOTAL,TCO_CODIGO,REC_NUMERO,REC_SUCURSAL,REC_FECHA)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec1!CLI_CODIGO) & ","
            sql = sql & XS(Rec1!CLI_RAZSOC) & ","
            sql = sql & XS(Rec1!CLI_DOMICI) & ","
            sql = sql & "NULL,NULL,NULL,NULL,"
            sql = sql & XS(Rec1!TCO_ABREVIA) & ","
            sql = sql & XS(Format(Rec1!FCL_NUMERO, "00000000")) & ","
            sql = sql & XS(Format(Rec1!FCL_SUCURSAL, "0000")) & ","
            sql = sql & XS(Rec1!FCL_FECHA) & ","
            sql = sql & XN(Rec1!REC_IMPORTE) & ","
            sql = sql & XN(Rec1!FCL_TOTAL) & ","
            sql = sql & XN(Rec1!REC_TOTAL) & ","
            sql = sql & XN(Rec1!TCO_CODIGO) & ","
            sql = sql & XS(Format(Rec1!REC_NUMERO, "00000000")) & ","
            sql = sql & XS(Format(Rec1!REC_SUCURSAL, "0000")) & ","
            sql = sql & XDQ(Rec1!REC_FECHA) & ")"
            DBConn.Execute sql
            
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub CmdNuevo_Click()
    txtCliente.Text = ""
    txtNroCobranza.Text = ""
    txtDesCli.Text = ""
    optNuevo.Value = True
    txtCliente.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoCobranzaCliente = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""
    optNuevo.Value = True
End Sub
Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        rec.Open BuscoCliente(txtCliente), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtDesCli_Change()
    If txtDesCli.Text = "" Then
        txtCliente.Text = ""
    End If
End Sub

Private Sub txtDesCli_GotFocus()
    SelecTexto txtDesCli
End Sub

Private Sub txtDesCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesCli_LostFocus()
    If txtCliente.Text = "" And txtDesCli.Text <> "" Then
        rec.Open BuscoCliente(txtDesCli), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 1
                frmBuscar.TxtDescriB.Text = txtDesCli.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 0
                    txtCliente.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 1
                    txtDesCli.Text = frmBuscar.grdBuscar.Text
                Else
                    txtCliente.SetFocus
                End If
            Else
                txtCliente.Text = rec!CLI_CODIGO
                txtDesCli.Text = rec!CLI_RAZSOC
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Function BuscoCliente(Cli As String) As String
    sql = "SELECT CLI_CODIGO, CLI_RAZSOC"
    sql = sql & " FROM CLIENTE"
    sql = sql & " WHERE "
    If txtCliente.Text <> "" Then
        sql = sql & " CLI_CODIGO=" & XN(Cli)
    Else
        sql = sql & " CLI_RAZSOC LIKE '" & Cli & "%'"
    End If
    BuscoCliente = sql
End Function

Private Sub txtNroCobranza_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroCobranza_LostFocus()
    If txtNroCobranza.Text <> "" Then
        sql = "SELECT REC_NUMERO FROM RECIBO_CLIENTE"
        sql = sql & " WHERE REC_LISTADO=" & XN(txtNroCobranza)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = True Then
            MsgBox "El número de cobranza no Existe", vbExclamation, TIT_MSGBOX
            txtNroCobranza.Text = ""
            txtNroCobranza.SetFocus
        End If
        rec.Close
    End If
End Sub
