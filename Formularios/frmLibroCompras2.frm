VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLibroCompras2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro IVA Compras"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAfip 
      Caption         =   "A&fip"
      Height          =   735
      Left            =   2040
      Picture         =   "frmLibroCompras2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Generar Archivo AFIP"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   3060
      Picture         =   "frmLibroCompras2.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2385
      Width           =   840
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4770
      Picture         =   "frmLibroCompras2.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2385
      Width           =   840
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   15
      TabIndex        =   8
      Top             =   0
      Width           =   5595
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   75
         TabIndex        =   9
         Top             =   1185
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   56623105
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   56623105
         CurrentDate     =   41098
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   660
         Width           =   960
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
         Left            =   2850
         TabIndex        =   12
         Top             =   255
         Width           =   1785
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
         Left            =   2850
         TabIndex        =   11
         Top             =   630
         Width           =   1785
      End
      Begin VB.Label lblPor 
         AutoSize        =   -1  'True
         Caption         =   "100 %"
         Height          =   195
         Left            =   4950
         TabIndex        =   10
         Top             =   1215
         Width           =   435
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
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   1590
      Width           =   5610
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   6
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   375
         Left            =   3810
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmLibroCompras2.frx":1016
      Height          =   735
      Left            =   3915
      Picture         =   "frmLibroCompras2.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2385
      Width           =   840
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   885
      Top             =   2565
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   1425
      Top             =   2535
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
      TabIndex        =   18
      Top             =   2550
      Width           =   750
   End
End
Attribute VB_Name = "frmLibroCompras2"
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

Private Sub buscarComprobantes()

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
        If Not IsNull(FechaDesde.Value) Then sql = sql & " AND FP.FPR_FECHA>=" & XDQ(FechaDesde.Value)
        If Not IsNull(FechaHasta.Value) Then sql = sql & " AND FP.FPR_FECHA<=" & XDQ(FechaHasta.Value)
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
                sql = sql & XS("F") & ","
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
         sql = sql & " P.PROV_RAZSOC,TC.TCO_ABREVIA,NP.CPR_IIBB"
         sql = sql & " FROM NOTA_CREDITO_PROVEEDOR NP"
         sql = sql & ",TIPO_COMPROBANTE TC , PROVEEDOR P"
         sql = sql & " WHERE"
         sql = sql & " NP.TCO_CODIGO=TC.TCO_CODIGO"
         sql = sql & " AND NP.TPR_CODIGO=P.TPR_CODIGO"
         sql = sql & " AND NP.PROV_CODIGO=P.PROV_CODIGO"
         sql = sql & " AND NP.EST_CODIGO <> 2 "
         If Not IsNull(FechaDesde) Then sql = sql & " AND NP.CPR_FECHA>=" & XDQ(FechaDesde)
         If Not IsNull(FechaHasta) Then sql = sql & " AND NP.CPR_FECHA<=" & XDQ(FechaHasta)
         sql = sql & " ORDER BY NP.CPR_FECHA"

         rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL,IVAGASTO)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!CPR_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(rec!CPR_NROSUCTXT & "-" & rec!CPR_NUMEROTXT) & ","
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
                sql = sql & XS("F") & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                sql = sql & XN(CStr((-1) * CDbl(rec!CPR_SUBTOTAL))) & ","
                sql = sql & XN(CDbl(rec!CPR_IVA)) & ","
                TotIva = (CDbl(rec!CPR_SUBTOTAL) * CDbl(rec!CPR_IVA)) / 100
                sql = sql & XN(CStr((-1) * CDbl(TotIva))) & ","
                sql = sql & XN(Chk0("")) & "," 'OTRO NETO
                sql = sql & XN(Chk0("")) & "," 'OTRO IVA
                sql = sql & XN(Chk0(rec!CPR_IIBB)) & "," 'IMPUESTOS
                sql = sql & XN(CStr((-1) * CDbl(rec!CPR_TOTAL))) & ","
                sql = sql & XN(Chk0("")) & ")"
                
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
        If Not IsNull(FechaDesde) Then sql = sql & " AND NP.DPR_PERIODO>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND NP.DPR_PERIODO<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY NP.DPR_FECHA"

        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL,IVAGASTO)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!DPR_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(rec!DPR_NROSUCTXT & "-" & rec!DPR_NUMEROTXT) & ","
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
                sql = sql & XS("F") & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                sql = sql & XN(rec!DPR_NETO) & ","
                sql = sql & XN(rec!DPR_IVA) & ","
                TotIva = (CDbl(rec!DPR_NETO) * CDbl(rec!DPR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(Chk0("")) & "," 'OTRO NETO
                sql = sql & XN(Chk0("")) & "," 'OTRO IVA
                sql = sql & XN(Chk0("")) & "," 'IMPUESTOS
                sql = sql & XN(rec!DPR_TOTAL) & ","
                sql = sql & XN(Chk0("")) & ")"
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
        'sql = sql & " AND GG.GGR_LIBROIVA='N'"
        If Not IsNull(FechaDesde) Then sql = sql & " AND GG.GGR_FECHACOMP>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND GG.GGR_FECHACOMP<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY GG.GGR_FECHACOMP"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL,IVAGASTO)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!GGR_FECHACOMP) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(rec!GGR_NROSUCTXT & "-" & rec!GGR_NROCOMPTXT) & ","
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
                sql = sql & XS("G") & ","
                sql = sql & XN(rec!GGR_NETO) & ","
                sql = sql & XN(rec!GGR_IVA) & ","
                    TotIva = (CDbl(rec!GGR_NETO) * CDbl(rec!GGR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & "," 'IVA BUENO
                sql = sql & XN(Chk0(rec!GGR_NETO)) & "," 'OTRO NETO
                    TotIva = (CDbl(Chk0(rec!GGR_NETO)) * CDbl(Chk0(rec!GGR_IVA))) / 100
                sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                If rec!GGR_IVA = 0 Then
                    'PONGO EL TOTAL CUANDO EL GASTO CON IVA = 0
                    ' EJ MONOTRIBUTISTA
                    sql = sql & 0 & "," 'IMPUESTOS
                Else
                    sql = sql & XN(Chk0(rec!GGR_IMPUESTOS)) & "," 'IMPUESTOS
                End If
                sql = sql & XN(rec!GGR_TOTAL) & ","
                sql = sql & XN(rec!GGR_IVA) & ")"
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
        
    Screen.MousePointer = vbNormal
    
    Exit Sub

CLAVO:
 Screen.MousePointer = vbNormal
 lblEstado.Caption = ""
 DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub


Private Sub CmdAceptar_Click()
    buscarComprobantes
    ListarLibroIVA
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub ListarLibroIVA()
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    'Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
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
    Rep.Action = 1
    
'    If optPantalla.Value = True Then
'        Rep.Destination = crptToWindow
'    ElseIf optImpresora.Value = True Then
'        Rep.Destination = crptToPrinter
 '   End If
     'Rep.Action = 1
     
     lblEstado.Caption = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     Rep.Formulas(2) = ""
End Sub

Private Sub cmdAfip_Click()
    buscarComprobantes
    CrearArchvioAFIP
    Screen.MousePointer = vbNormal
End Sub
Private Function CrearArchvioAFIP()
    Dim Cadena() As String
    Dim Alicuota() As String
    
    Dim Fecha As String
    Dim TipoCbte As String
    Dim PtoVenta As String
    Dim NroCbte As String
    Dim NroDespacho As String
    Dim CodDocProv As String
    Dim NroIdVend As String
    Dim NombreProv As String
    Dim TOTALAux As String
    Dim TOTAL As String
    Dim NoNetoGravado As String
    Dim OpExentas As String
    Dim IVA As String
    Dim OtroImp As String
    Dim PerIIBB As String
    Dim ImpMuni As String
    Dim ImpInt As String
    Dim Moneda As String
    Dim TCambio As String
    Dim CantIVA As String
    Dim CodOp As String
    Dim CredFiscalComp As String
    Dim OtrosTrib As String
    Dim CuitEmisor As String
    Dim DenEmisor As String
    Dim IVAcom As String
    Dim MesAnio As String
    Dim CredFiscalCompAux As String
    
    Dim I As Integer
    Dim cantRegistros As Integer
    sql = "SELECT * FROM TMP_LIBRO_IVA_COMPRAS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cantRegistros = rec.RecordCount
        ReDim Cadena(cantRegistros) As String
        ReDim Alicuota(cantRegistros) As String
        I = 1
        Do While rec.EOF = False
            '--------Fecha Comprobante
            Fecha = Replace(rec!Fecha, "/", "")
            Fecha = Right(Fecha, 4) + Right(Left(Fecha, 4), 2) + Left(Left(Fecha, 4), 2)
            
            '--------Tipo Comprobante
            TipoCbte = Tipo_Cbte_Afip(rec!COMPROBANTE)
            TipoCbte = String(3 - Len(TipoCbte), "0") & TipoCbte
            
            '--------Punto de Venta
            PtoVenta = Left(rec!Numero, 4)
            PtoVenta = String(5 - Len(PtoVenta), "0") & PtoVenta
            
            '--------Numero de Comprobante
            NroCbte = Right(rec!Numero, 8)
            NroCbte = String(20 - Len(NroCbte), "0") & NroCbte
            
            'Nro de despacho de Importacion
            NroDespacho = " "
            NroDespacho = String(16 - Len(NroDespacho), " ") & NroDespacho
            
            'Codigo de documento del vendedor
            CodDocProv = 80 'CUIT
            
            'Numero de Identificacion del Vendedor
            NroIdVend = Replace(rec!cuit, "-", "")
            NroIdVend = String(20 - Len(NroIdVend), "0") & NroIdVend
            
            'Apellido y Nombre del vendedor
            NombreProv = Left(rec!Proveedor, 30)
            NombreProv = NombreProv & String(30 - Len(NombreProv), " ")
            
            'Importe Total de la operacion
            If rec!TOTAL < 0 Then
                TOTALAux = Replace(Format(rec!TOTAL, "#,##0.00"), ",", "")
                TOTAL = Replace(TOTALAux, "-", "0")
                TOTAL = Replace(TOTAL, ".", "")
                TOTAL = String(15 - Len(TOTAL), "0") & TOTAL
                
            Else
                TOTAL = Replace(Format(rec!TOTAL, "#,##0.00"), ",", "")
                TOTAL = Replace(TOTAL, ".", "")
                TOTAL = String(15 - Len(TOTAL), "0") & TOTAL
            End If
            
            
            'Importe Total de conceptos que no integran el precio neto gravado. Ej. Facturas Pignata
            If rec!COMPROBANTE = "FAC-B" Or rec!COMPROBANTE = "FAC-C" Or rec!COMPROBANTE = "T" Or rec!COMPROBANTE = "OCRG" Then
                NoNetoGravado = 0
            Else
                NoNetoGravado = Replace(Format(rec!IMPUESTOS, "#,##0.00"), ",", "")
            End If
            NoNetoGravado = String(15 - Len(NoNetoGravado), "0") & NoNetoGravado
            
            
            'Importe de operaciones exentas
            OpExentas = 0
            OpExentas = String(15 - Len(OpExentas), "0") & OpExentas
            
            'Importe de percepciones o pagos a cuenta del Impuesto al Valor Agregado
            IVA = 0
            IVA = String(15 - Len(IVA), "0") & IVA
            
            'Importe de percepciones o pagos a cuenta de otros impuestos nacionales
            OtroImp = 0
            OtroImp = String(15 - Len(OtroImp), "0") & OtroImp
            
            'Importe de percepciones de Ingresos Brutos
            PerIIBB = 0
            PerIIBB = String(15 - Len(PerIIBB), "0") & PerIIBB
            
            'Importe de percepciones Impuestos Municipales
            ImpMuni = 0
            ImpMuni = String(15 - Len(ImpMuni), "0") & ImpMuni
            
            'Importe de impuestos internos
            ImpInt = 0
            ImpInt = String(15 - Len(ImpInt), "0") & ImpInt
            
            'Codigo de Moneda
            Moneda = "PES"
            Moneda = String(3 - Len(Moneda), "0") & Moneda
            
            'Tipo de Cambio
            TCambio = "1000000"
            TCambio = String(10 - Len(TCambio), "0") & TCambio
            
            'Cantidad de Alicuotas de IVA
            If rec!COMPROBANTE = "FAC-B" Or rec!COMPROBANTE = "FAC-C" Or rec!COMPROBANTE = "T" Or rec!COMPROBANTE = "TFB" Then
                CantIVA = 0
            Else
                CantIVA = 1
            End If
            CantIVA = String(1 - Len(CantIVA), "0") & CantIVA
            
            'Codigo de Operacion
            CodOp = 0
            If rec!COMPROBANTE = "OCRG" Or rec!COMPROBANTE = "FAC-A" Then
                If rec!IVA = 0 Then 'CENTRO COMERCIAL
                    CodOp = "A"
                End If
            End If
            CodOp = String(1 - Len(CodOp), "0") & CodOp
            
            'Credito Fiscal computable //monto iva
            If rec!COMPROBANTE = "FAC-B" Or rec!COMPROBANTE = "FAC-C" Or rec!COMPROBANTE = "T" Then
                CredFiscalComp = 0
            Else
                If rec!TotIva < 0 Then
                    CredFiscalCompAux = Replace(Format(rec!TotIva, "#,##0.00"), ",", "")
                    CredFiscalComp = Replace(CredFiscalCompAux, "-", "0")
                    CredFiscalComp = Replace(CredFiscalComp, ".", "")
                    CredFiscalComp = String(15 - Len(CredFiscalComp), "0") & CredFiscalComp
                    
                Else
                    CredFiscalComp = Replace(Format(rec!TotIva, "#,##0.00"), ",", "")
                    CredFiscalComp = Replace(CredFiscalComp, ".", "")
                    CredFiscalComp = String(15 - Len(CredFiscalComp), "0") & CredFiscalComp
                End If
            End If
            CredFiscalComp = String(15 - Len(CredFiscalComp), "0") & CredFiscalComp
            
            'Otros Tributos
            OtrosTrib = 0
            OtrosTrib = String(15 - Len(OtrosTrib), "0") & OtrosTrib
            
            'CUIT emisor/corredor
'            Rec1.Open "SELECT CUIT FROM PARAMETROS", DBConn, adOpenStatic, adLockOptimistic
'            If Rec1.EOF = False Then
'                CuitEmisor = Rec1!cuit
'            End If
'            Rec1.Close
            CuitEmisor = 0
            CuitEmisor = String(11 - Len(CuitEmisor), "0") & CuitEmisor
            
            'Denominacion del emisor/corredor
'            Rec1.Open "SELECT RAZ_SOCIAL FROM PARAMETROS", DBConn, adOpenStatic, adLockOptimistic
'            If Rec1.EOF = False Then
'                DenEmisor = Left(Rec1!RAZ_SOCIAL, 30)
'            End If
'            Rec1.Close
            DenEmisor = " "
            DenEmisor = DenEmisor & String(30 - Len(DenEmisor), " ")
            
            'IVA Comision
            IVAcom = 0
            IVAcom = String(15 - Len(IVAcom), "0") & IVAcom
            
            
            '----------archivo alicuotas iva--------------
            Dim neto As String
            Dim netoaux As String
            Dim alicuotaIVA As String
            
            'importe neto gravado
            'Importe Total de la operacion
            If rec!subtotal < 0 Then
                netoaux = Replace(Format(rec!subtotal, "#,##0.00"), ",", "")
                neto = Replace(netoaux, "-", "0")
                neto = Replace(neto, ".", "")
                neto = String(15 - Len(neto), "0") & neto
                
            Else
                neto = Replace(Format(rec!subtotal, "#,##0.00"), ",", "")
                neto = Replace(neto, ".", "")
                neto = String(15 - Len(neto), "0") & neto
            End If
            
            '----------- alicuota IVA ------------
            ' no se cargan alicuotas de IVA para
            ' (006) - FACTURA B
            ' (011) - FACTURA C
            ' (083 -  TIQUE
            
            Select Case rec!IVA
                Case "0"
                    alicuotaIVA = "0003"
                Case "10,5"
                    alicuotaIVA = "0004"
                Case "21"
                    alicuotaIVA = "0005"
                Case "27"
                    alicuotaIVA = "0006"
                Case "5"
                    alicuotaIVA = "0008"
                Case "2,5"
                    alicuotaIVA = "0009"
                
                
            End Select
            'ARMO UNA LINEA DEL ARCHIVO COMPROBANTES
            Cadena(I) = Fecha & TipoCbte & PtoVenta & NroCbte & NroDespacho & _
                        CodDocProv & NroIdVend & NombreProv & TOTAL & NoNetoGravado & _
                        OpExentas & IVA & OtroImp & PerIIBB & ImpMuni & ImpInt & Moneda & _
                        TCambio & CantIVA & CodOp & CredFiscalComp & OtrosTrib & CuitEmisor & _
                        DenEmisor & IVAcom
            
            'If rec!COMPROBANTE <> "FAC-B" And rec!COMPROBANTE <> "FAC-C" And rec!COMPROBANTE <> "T" Then
            'ARMO UNA LINEA DEL ARCHIVO ALICUOTAS
                Alicuota(I) = TipoCbte & PtoVenta & NroCbte & _
                            CodDocProv & NroIdVend & neto & alicuotaIVA & CredFiscalComp
            
            'End If
            rec.MoveNext
            I = I + 1
            
        Loop
        
    End If
    
    rec.Close

    'Cadena = "Texto a enviar"
    ' ver aca como hago el append de cada cadena
    If lblPeriodo1 <> "" Then
        MesAnio = Replace(lblPeriodo1.Caption, "/", "")
    End If
    
    'BORRO LOS ARCHIVOS SI EXISTEN
        
    If EstadoDeArchivo(DirAFIP & "Compras_AFIP_" & MesAnio & ".txt") Then
        Kill (DirAFIP & "Compras_AFIP_" & MesAnio & ".txt")
    End If
    If EstadoDeArchivo(DirAFIP & "Alicuotas_AFIP_" & MesAnio & ".txt") Then
        Kill (DirAFIP & "Alicuotas_AFIP_" & MesAnio & ".txt")
    End If
    
    'GENERO LOS ARCHIVOS
    For I = 0 To cantRegistros
        Open DirAFIP & "Compras_AFIP_" & MesAnio & ".txt" For Append As #1
        Print #1, Cadena(I)
        Close #1
    Next
    MsgBox "Se genero correctamente el archivo " & DirAFIP & "Compras_AFIP_" & MesAnio & ".txt", vbInformation, TIT_MSGBOX
    
    
    For I = 0 To cantRegistros
        Open DirAFIP & "Alicuotas_AFIP_" & MesAnio & ".txt" For Append As #1
        
        Print #1, Alicuota(I)
        Close #1
        
    Next
    MsgBox "Se genero correctamente el archivo " & DirAFIP & "Alicuotas_AFIP_" & MesAnio & ".txt", vbInformation, TIT_MSGBOX
    
End Function

Public Function EstadoDeArchivo(ByVal Archivo As String) As Boolean
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If (fso.FileExists(Archivo)) Then
    EstadoDeArchivo = True
Else
    EstadoDeArchivo = False
End If
End Function

Function Tipo_Cbte_Afip(Codigo As String) As String
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT TCO_CODTABLA FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_ABREVIA LIKE '" & Codigo & "'"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Tipo_Cbte_Afip = IIf(IsNull(Rec1!TCO_CODTABLA), 0, Rec1!TCO_CODTABLA)
    Else
        Tipo_Cbte_Afip = 0
    End If
    Rec1.Close
End Function
Private Sub CmdNuevo_Click()
    FechaDesde.Value = Null
    lblPeriodo1.Caption = ""
    FechaHasta.Value = Null
    lblPeriodo2.Caption = ""
    FechaDesde.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmLibroCompras2 = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    lblEstado.Caption = ""
    lblPor.Caption = "100 %"
    Call Centrar_pantalla(Me)
    Set rec = New ADODB.Recordset
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    'otroejemplo
End Sub

Private Sub FechaDesde_LostFocus()
    If Not IsNull(FechaDesde.Value) Then
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
Sub otroejemplo()

Open DirAFIP & "Compras_AFIP_AGOSTO2015.txt" For Input As #1

'Luego se tiene que leer con

Dim Linea As String, TOTAL As String
Do Until EOF(1)
Line Input #1, Linea
TOTAL = TOTAL + Linea + vbCrLf
Loop
Close #1


End Sub
