VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmLibroIvaCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro IVA Compras"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5685
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmLibroIvaCompras.frx":0000
      Height          =   735
      Left            =   3915
      Picture         =   "frmLibroIvaCompras.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2445
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
      Top             =   1650
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
      Height          =   1575
      Left            =   15
      TabIndex        =   6
      Top             =   60
      Width           =   5595
      Begin FechaCtl.Fecha FechaDesde 
         Height          =   300
         Left            =   1305
         TabIndex        =   0
         Top             =   255
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin FechaCtl.Fecha FechaHasta 
         Height          =   285
         Left            =   1305
         TabIndex        =   1
         Top             =   630
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   75
         TabIndex        =   11
         Top             =   1185
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
         Left            =   4950
         TabIndex        =   12
         Top             =   1215
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
         Left            =   2490
         TabIndex        =   10
         Top             =   630
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
         Left            =   2490
         TabIndex        =   9
         Top             =   255
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   660
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4770
      Picture         =   "frmLibroIvaCompras.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2445
      Width           =   840
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   3060
      Picture         =   "frmLibroIvaCompras.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2445
      Width           =   840
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   1605
      Top             =   2625
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2145
      Top             =   2595
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
      Top             =   2610
      Width           =   750
   End
End
Attribute VB_Name = "frmLibroIvaCompras"
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
     If FechaHasta.Value = Null Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaHasta.SetFocus
        Exit Sub
     End If
     
     On Error GoTo CLAVO
     Screen.MousePointer = vbHourglass
     DBConn.BeginTrans
     lblEstado.Caption = "Buscando Datos..."
     
        'BORRO LA TABLA TMP_LIBRO_IVA_COMPRAS
        sql = "DELETE FROM TMP_LIBRO_IVA_COMPRAS"
        DBConn.Execute sql
        
        'BUSCO FACTURAS
        sql = "SELECT FP.FPR_NROSUCTXT,FP.FPR_NUMEROTXT,"
        sql = sql & " FP.FPR_FECHA,FP.FPR_IVA,FP.FPR_SUBTOTAL,FP.FPR_TOTAL,"
        'sql = sql & " FP.FPR_IVA1,FP.FPR_NETO1,FP.FPR_IMPUESTOS,"
        sql = sql & " P.PROV_CUIT,P.PROV_INGBRU,"
        sql = sql & " P.PROV_RAZSOC,TC.TCO_ABREVIA"
        sql = sql & " FROM FACTURA_PROVEEDOR FP, PROVEEDOR P"
        sql = sql & " ,TIPO_COMPROBANTE TC"
        sql = sql & " WHERE"
        sql = sql & " FP.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND FP.TPR_CODIGO=P.TPR_CODIGO"
        sql = sql & " AND FP.PROV_CODIGO=P.PROV_CODIGO"
        sql = sql & " AND FP.EST_CODIGO<> 2"
        If FechaDesde <> "" Then sql = sql & " AND FP.FPR_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND FP.FPR_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY FP.FPR_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,"
                sql = sql & "TOTAL)"
                sql = sql & " VALUES ("
                sql = sql & XDQ(rec!FPR_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(rec!FPR_NROSUCTXT & "-" & rec!FPR_NUMEROTXT) & ","
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                sql = sql & XN(rec!FPR_SUBTOTAL) & ","
                sql = sql & XN(rec!FPR_IVA) & ","
                    TotIva = (CDbl(rec!FPR_SUBTOTAL) * CDbl(rec!FPR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & "," 'IVA BUENO
                
                sql = sql & XN(Chk0("")) & "," 'OTRO NETO
                sql = sql & XN(Chk0("")) & "," 'OTRO IVA
                sql = sql & XN(Chk0("")) & "," 'IMPUESTOS
                
'                sql = sql & XN(Chk0(rec!FPR_NETO1)) & ","
'                    TotIva = (CDbl(Chk0(rec!FPR_NETO1)) * CDbl(Chk0(rec!FPR_IVA1))) / 100
'                sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
'                sql = sql & XN(Chk0(rec!FPR_IMPUESTOS)) & ","
                sql = sql & XN(rec!FPR_TOTAL) & ")"
                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
        
        'BUSCO NOTA DE CREDITO------------------------------------
         sql = "SELECT NP.CPR_NROSUCTXT,NP.CPR_NUMEROTXT,"
         sql = sql & " NP.CPR_FECHA,NP.CPR_IVA,NP.CPR_SUBTOTAL,NP.CPR_TOTAL,"
         sql = sql & " P.PROV_CUIT,P.PROV_INGBRU,"
         sql = sql & " P.PROV_RAZSOC,TC.TCO_ABREVIA"
         sql = sql & " FROM NOTA_CREDITO_PROVEEDOR NP"
         sql = sql & ",TIPO_COMPROBANTE TC , PROVEEDOR P"
         sql = sql & " WHERE"
         sql = sql & " NP.TCO_CODIGO=TC.TCO_CODIGO"
         sql = sql & " AND NP.TPR_CODIGO=P.TPR_CODIGO"
         sql = sql & " AND NP.PROV_CODIGO=P.PROV_CODIGO"
         sql = sql & " AND NP.EST_CODIGO <> 2 "
         If FechaDesde <> "" Then sql = sql & " AND NP.CPR_FECHA>=" & XDQ(FechaDesde)
         If FechaHasta <> "" Then sql = sql & " AND NP.CPR_FECHA<=" & XDQ(FechaHasta)
         sql = sql & " ORDER BY NP.CPR_FECHA"
         
         rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!CPR_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(rec!CPR_NROSUCTXT & "-" & rec!CPR_NUMEROTXT) & ","
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                sql = sql & XN(CStr((-1) * CDbl(rec!CPR_SUBTOTAL))) & ","
                sql = sql & XN(CStr((-1) * CDbl(rec!CPR_IVA))) & ","
                TotIva = (CDbl(rec!CPR_SUBTOTAL) * CDbl(rec!CPR_IVA)) / 100
                sql = sql & XN(CStr((-1) * CDbl(TotIva))) & ","
                sql = sql & XN(Chk0("")) & "," 'OTRO NETO
                sql = sql & XN(Chk0("")) & "," 'OTRO IVA
                sql = sql & XN(Chk0("")) & "," 'IMPUESTOS
                sql = sql & XN(CStr((-1) * CDbl(rec!CPR_TOTAL))) & ")"

                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
        
        'BUSCO NOTA DE DEBITO SERVICIOS Y CHEQUES DEVUELTOS-----
        sql = "SELECT NP.DPR_NROSUCTXT,NP.DPR_NUMEROTXT,NP.DPR_FECHA,NP.DPR_IVA,NP.DPR_NETO,NP.DPR_TOTAL,"
        sql = sql & " P.PROV_CUIT,P.PROV_INGBRU,"
        sql = sql & " P.PROV_RAZSOC,TC.TCO_ABREVIA"
        sql = sql & " FROM NOTA_DEBITO_PROVEEDOR NP,"
        sql = sql & " TIPO_COMPROBANTE TC, PROVEEDOR P"
        sql = sql & " WHERE NP.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND NP.TPR_CODIGO=P.TPR_CODIGO"
        sql = sql & " AND NP.PROV_CODIGO=P.PROV_CODIGO"
        sql = sql & " AND NP.EST_CODIGO <> 2"
        If FechaDesde <> "" Then sql = sql & " AND NP.DPR_PERIODO>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND NP.DPR_PERIODO<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY NP.DPR_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!DPR_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(rec!DPR_NROSUCTXT & "-" & rec!DPR_NUMEROTXT) & ","
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                sql = sql & XN(rec!DPR_NETO) & ","
                sql = sql & XN(rec!DPR_IVA) & ","
                TotIva = (CDbl(rec!DPR_NETO) * CDbl(rec!DPR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(Chk0("")) & "," 'OTRO NETO
                sql = sql & XN(Chk0("")) & "," 'OTRO IVA
                sql = sql & XN(Chk0("")) & "," 'IMPUESTOS
                sql = sql & XN(rec!DPR_TOTAL) & ")"
                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
        
        'BUSCO COMPROBANTES DENTRO DE LOS GASTOS GENERALES -----
        sql = "SELECT GG.GGR_NROSUCTXT,GG.GGR_NROCOMPTXT,GG.GGR_FECHACOMP,GG.GGR_IVA,GG.GGR_NETO,GG.GGR_TOTAL,"
        sql = sql & " GG.GGR_IVA1,GG.GGR_NETO1,GG.GGR_IMPUESTOS,"
        sql = sql & " P.PROV_CUIT,P.PROV_INGBRU,"
        sql = sql & " P.PROV_RAZSOC,TC.TCO_ABREVIA"
        sql = sql & " FROM GASTOS_GENERALES GG,"
        sql = sql & " TIPO_COMPROBANTE TC, PROVEEDOR P"
        sql = sql & " WHERE GG.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND GG.TPR_CODIGO=P.TPR_CODIGO"
        sql = sql & " AND GG.PROV_CODIGO=P.PROV_CODIGO"
        sql = sql & " AND GG.GGR_LIBROIVA='S'"
        If FechaDesde <> "" Then sql = sql & " AND GG.GGR_PERIODO>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND GG.GGR_PERIODO<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY GG.GGR_FECHACOMP"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!GGR_FECHACOMP) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(rec!GGR_NROSUCTXT & "-" & rec!GGR_NROCOMPTXT) & ","
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                sql = sql & XN(rec!GGR_NETO) & ","
                sql = sql & XN(rec!GGR_IVA) & ","
                    TotIva = (CDbl(rec!GGR_NETO) * CDbl(rec!GGR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(Chk0(rec!GGR_NETO1)) & "," 'OTRO NETO
                    TotIva = (CDbl(Chk0(rec!GGR_NETO1)) * CDbl(Chk0(rec!GGR_IVA1))) / 100
                sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                sql = sql & XN(Chk0(rec!GGR_IMPUESTOS)) & "," 'IMPUESTOS
                sql = sql & XN(rec!GGR_TOTAL) & ")"
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
    Rep.WindowState = crptNormal
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
    
    
    Rep.WindowTitle = "Libro I.V.A. Compras"
    Rep.ReportFileName = DRIVE & DirReport & "rptlibroivacompras.rpt"
    
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
    Set frmLibroIvaCompras = Nothing
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
