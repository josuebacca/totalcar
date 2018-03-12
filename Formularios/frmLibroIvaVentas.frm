VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmLibroIvaVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro IVA Ventas"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5655
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmLibroIvaVentas.frx":0000
      Height          =   735
      Left            =   3915
      Picture         =   "frmLibroIvaVentas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2910
      Width           =   840
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
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   2115
      Width           =   5610
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   375
         Left            =   3810
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1665
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   15
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   14
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2040
      Left            =   15
      TabIndex        =   6
      Top             =   60
      Width           =   5595
      Begin FechaCtl.Fecha FechaDesde 
         Height          =   300
         Left            =   1440
         TabIndex        =   0
         Top             =   510
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin FechaCtl.Fecha FechaHasta 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   210
         TabIndex        =   11
         Top             =   1395
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblPor 
         AutoSize        =   -1  'True
         Caption         =   "100 %"
         Height          =   195
         Left            =   5085
         TabIndex        =   12
         Top             =   1425
         Width           =   435
      End
      Begin VB.Label lblPeriodo2 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2625
         TabIndex        =   10
         Top             =   840
         Width           =   1785
      End
      Begin VB.Label lblPeriodo1 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2625
         TabIndex        =   9
         Top             =   510
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   405
         TabIndex        =   8
         Top             =   870
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   525
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4770
      Picture         =   "frmLibroIvaVentas.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2910
      Width           =   840
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   3060
      Picture         =   "frmLibroIvaVentas.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2910
      Width           =   840
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   1605
      Top             =   3090
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2145
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   105
      TabIndex        =   5
      Top             =   3075
      Width           =   750
   End
End
Attribute VB_Name = "frmLibroIvaVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Registro As Long
Dim Tamanio As Long
Dim TotIva As Double

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub CmdAceptar_Click()
     Registro = 0
     Tamanio = 0
     TotIva = 0
     
     If IsNull(FechaDesde.Value) Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaDesde.SetFocus
        Exit Sub
     End If
     If IsNull(FechaHasta.Value) Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaHasta.SetFocus
        Exit Sub
     End If
     
     On Error GoTo CLAVO
     Screen.MousePointer = vbHourglass
     DBConn.BeginTrans
     lblEstado.Caption = "Buscando Datos..."
     
        'BORRO LA TABLA TEMPORAL DE IVA VENTAS
        sql = "DELETE FROM TMP_LIBRO_IVA_VENTAS"
        DBConn.Execute sql
        
        'BUSCO FACTURAS
        sql = "SELECT FC.FCL_NUMEROTXT, FC.FCL_SUCURSAL, FC.FCL_FECHA, FC.FCL_IVA,"
        sql = sql & " FC.FCL_SUBTOTAL, FC.FCL_TOTAL,"
        sql = sql & " FC.EST_CODIGO, C.CLI_CUIT, C.CLI_INGBRU,"
        sql = sql & " C.CLI_RAZSOC, TC.TCO_ABREVIA"
        sql = sql & " FROM FACTURA_CLIENTE FC, REMITO_CLIENTE RC, CLIENTE C,"
        sql = sql & " TIPO_COMPROBANTE TC"
        sql = sql & " WHERE"
        sql = sql & " FC.RCL_NUMERO=RC.RCL_NUMERO"
        sql = sql & " AND FC.RCL_SUCURSAL=RC.RCL_SUCURSAL"
        sql = sql & " AND FC.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
        sql = sql & " AND FC.EST_CODIGO <> 1" 'ESTADO DEFINITIVO Y ANULADO
        If FechaDesde <> "" Then sql = sql & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY FC.FCL_NUMEROTXT,FC.FCL_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!FCL_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(Format(rec!FCL_SUCURSAL, "0000") & "-" & rec!FCL_NUMEROTXT) & ","
                sql = sql & XS(rec!CLI_RAZSOC) & ","
                sql = sql & XS(Format(rec!CLI_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                If rec!EST_CODIGO = 2 Then
                    sql = sql & "0" & ","
                    sql = sql & XN(rec!FCL_IVA) & ","
                    sql = sql & "0" & ","
                    sql = sql & "0" & ")"
                Else
                    sql = sql & XN(rec!FCL_SUBTOTAL) & ","
                    sql = sql & XN(rec!FCL_IVA) & ","
                    TotIva = (CDbl(rec!FCL_SUBTOTAL) * CDbl(rec!FCL_IVA)) / 100
                    sql = sql & XN(CStr(TotIva)) & ","
                    sql = sql & XN(rec!FCL_TOTAL) & ")"
                End If
                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
        
        'BUSCO NOTA DE CREDITO------------------------------------
         sql = "SELECT NC.NCC_NUMEROTXT, NC.NCC_SUCURSAL, NC.NCC_FECHA, NC.NCC_IVA,"
         sql = sql & " NC.NCC_SUBTOTAL, NC.NCC_TOTAL,"
         sql = sql & " NC.EST_CODIGO,C.CLI_CUIT,C.CLI_INGBRU,"
         sql = sql & " C.CLI_RAZSOC,TC.TCO_ABREVIA"
         sql = sql & " FROM NOTA_CREDITO_CLIENTE NC"
         sql = sql & ",TIPO_COMPROBANTE TC , CLIENTE C"
         sql = sql & " WHERE"
         sql = sql & " NC.TCO_CODIGO=TC.TCO_CODIGO"
         sql = sql & " AND NC.CLI_CODIGO=C.CLI_CODIGO"
         If FechaDesde <> "" Then sql = sql & " AND NC.NCC_FECHA>=" & XDQ(FechaDesde)
         If FechaHasta <> "" Then sql = sql & " AND NC.NCC_FECHA<=" & XDQ(FechaHasta)
         sql = sql & " ORDER BY NC.NCC_FECHA"
         
         rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!NCC_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(Format(rec!NCC_SUCURSAL, "0000") & "-" & rec!NCC_NUMEROTXT) & ","
                sql = sql & XS(rec!CLI_RAZSOC) & ","
                sql = sql & XS(Format(rec!CLI_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                If rec!EST_CODIGO = 2 Then
                    sql = sql & "0" & ","
                    sql = sql & XN(rec!NCC_IVA) & ","
                    sql = sql & "0" & ","
                    sql = sql & "0" & ")"
                Else
                    sql = sql & XN(CStr((-1) * CDbl(IIf(IsNull(rec!NCC_SUBTOTAL), 0, rec!NCC_SUBTOTAL)))) & ","
                    sql = sql & XN(rec!NCC_IVA) & ","
                    TotIva = (CDbl(IIf(IsNull(rec!NCC_SUBTOTAL), 0, rec!NCC_SUBTOTAL)) * CDbl(rec!NCC_IVA)) / 100
                    sql = sql & XN(CStr((-1) * CDbl(TotIva))) & ","
                    sql = sql & XN(CStr((-1) * CDbl(IIf(IsNull(rec!NCC_TOTAL), 0, rec!NCC_TOTAL)))) & ")"
                End If
                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
        
        'BUSCO NOTA DE DEBITO SERVICIOS Y CHEQUES DEVUELTOS-----
        sql = "SELECT ND.NDC_NUMEROTXT, ND.NDC_SUCURSAL, ND.NDC_FECHA, ND.NDC_IVA,"
        sql = sql & " ND.NDC_SUBTOTAL, ND.NDC_TOTAL,"
        sql = sql & " ND.EST_CODIGO, C.CLI_CUIT, C.CLI_INGBRU,"
        sql = sql & " C.CLI_RAZSOC,TC.TCO_ABREVIA"
        sql = sql & " FROM NOTA_DEBITO_CLIENTE ND,"
        sql = sql & " TIPO_COMPROBANTE TC , CLIENTE C"
        sql = sql & " WHERE ND.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND ND.CLI_CODIGO=C.CLI_CODIGO"
        If FechaDesde <> "" Then sql = sql & " AND ND.NDC_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND ND.NDC_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY ND.NDC_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!NDC_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(Format(rec!NDC_SUCURSAL, "0000") & "-" & rec!NDC_NUMEROTXT) & ","
                sql = sql & XS(rec!CLI_RAZSOC) & ","
                sql = sql & XS(Format(rec!CLI_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                If rec!EST_CODIGO = 2 Then
                    sql = sql & "0" & ","
                    sql = sql & XN(rec!NDC_IVA) & ","
                    sql = sql & "0" & ","
                    sql = sql & "0" & ")"
                Else
                    sql = sql & XN(rec!NDC_SUBTOTAL) & ","
                    sql = sql & XN(rec!NDC_IVA) & ","
                    TotIva = (CDbl(rec!NDC_SUBTOTAL) * CDbl(rec!NDC_IVA)) / 100
                    sql = sql & XN(CStr(TotIva)) & ","
                    sql = sql & XN(rec!NDC_TOTAL) & ")"
                End If
                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
        
    lblEstado.Caption = ""
    DBConn.CommitTrans
    'cargo el reporte
    ListarLibroIVA
        
    Screen.MousePointer = vbNormal
    
    Exit Sub

CLAVO:
 Screen.MousePointer = vbNormal
 lblEstado.Caption = ""
 DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub ListarLibroIVA()
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIHDG"
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
        
    sql = "SELECT CUIT,ING_BRUTOS,RAZ_SOCIAL FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Rep.Formulas(0) = "EMPRESA='     Empresa:  " & Trim(rec!RAZ_SOCIAL) & "'"
        Rep.Formulas(1) = "CUIT='       C.U.I.T.:  " & Format(rec!cuit, "##-########-#") & "'"
        Rep.Formulas(2) = "INGBRUTOS='Ing. Brutos:  " & Format(rec!ING_BRUTOS, "###-#####-##") & "'"
    End If
    rec.Close
    
    Rep.WindowTitle = "Libro I.V.A. Ventas"
    Rep.ReportFileName = DRIVE & DirReport & "rptlibroivaventas.rpt"
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     lblEstado.Caption = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     Rep.Formulas(2) = ""
End Sub

Private Sub CmdNuevo_Click()
    FechaDesde.Value = Null
    lblPeriodo1.Caption = ""
    FechaHasta.Value = Null
    lblPeriodo2.Caption = ""
    FechaDesde.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set frmLibroIvaVentas = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    lblEstado.Caption = ""
    lblPor.Caption = "100 %"
    Set rec = New ADODB.Recordset
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub FechaDesde_LostFocus()
    If Trim(FechaDesde.Value) <> "" Then
        FechaHasta.Value = FechaDesde.Value
        lblPeriodo1.Caption = UCase(Format(FechaDesde.Value, "mmmm/yyyy"))
    Else
        lblPeriodo1.Caption = ""
    End If
End Sub

Private Sub FechaHasta_LostFocus()
    If Trim(FechaHasta.Value) <> "" Then
        lblPeriodo2.Caption = UCase(Format(FechaHasta.Value, "mmmm/yyyy"))
    Else
        lblPeriodo2.Caption = ""
    End If
End Sub
