VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLibroVentas2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro IVA Ventas"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   15
      Top             =   2040
      Width           =   5610
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   16
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   2
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   375
         Left            =   3810
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1665
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
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   3025
      Picture         =   "frmLibroVentas2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2850
      Width           =   840
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4755
      Picture         =   "frmLibroVentas2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2850
      Width           =   840
   End
   Begin VB.Frame Frame2 
      Height          =   2040
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5595
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   1395
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   16777217
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   16777217
         CurrentDate     =   41098
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   405
         TabIndex        =   12
         Top             =   870
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
         Left            =   2985
         TabIndex        =   11
         Top             =   510
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
         Left            =   2985
         TabIndex        =   10
         Top             =   840
         Width           =   1785
      End
      Begin VB.Label lblPor 
         AutoSize        =   -1  'True
         Caption         =   "100 %"
         Height          =   195
         Left            =   5085
         TabIndex        =   9
         Top             =   1425
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmLibroVentas2.frx":0614
      Height          =   735
      Left            =   3890
      Picture         =   "frmLibroVentas2.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2850
      Width           =   840
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   870
      Top             =   3030
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   1290
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAfip 
      Caption         =   "A&fip"
      Height          =   735
      Left            =   2160
      Picture         =   "frmLibroVentas2.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Generar Archivo AFIP"
      Top             =   2850
      Width           =   840
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
      Left            =   90
      TabIndex        =   14
      Top             =   3015
      Width           =   750
   End
End
Attribute VB_Name = "frmLibroVentas2"
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
    buscarComprobantes
    ListarLibroIVA
End Sub
Private Function buscarComprobantes()
     Registro = 0
     Tamanio = 0
     TotIva = 0
     
     If IsNull(FechaDesde.Value) Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaDesde.SetFocus
        Exit Function
     End If
     If IsNull(FechaHasta.Value) Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaHasta.SetFocus
        Exit Function
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
        'sql = sql & " AND FC.EST_CODIGO <> 2 " 'ESTADO DISTINTO A ANULADO"
        If Not IsNull(FechaDesde) Then sql = sql & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY TC.TCO_ABREVIA" ', FC.FCL_NUMEROTXT,FC.FCL_FECHA"
        
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
                    TotIva = CDbl(rec!FCL_SUBTOTAL) * (CDbl(rec!FCL_IVA) / 100)
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
         sql = sql & " C.CLI_RAZSOC,TC.TCO_ABREVIA,TC.TCO_CODIGO"
         sql = sql & " FROM NOTA_CREDITO_CLIENTE NC"
         sql = sql & ",TIPO_COMPROBANTE TC , CLIENTE C"
         sql = sql & " WHERE"
         sql = sql & " NC.TCO_CODIGO=TC.TCO_CODIGO"
         sql = sql & " AND NC.CLI_CODIGO=C.CLI_CODIGO"
         If Not IsNull(FechaDesde) Then sql = sql & " AND NC.NCC_FECHA>=" & XDQ(FechaDesde)
         If Not IsNull(FechaHasta) Then sql = sql & " AND NC.NCC_FECHA<=" & XDQ(FechaHasta)
         sql = sql & " ORDER BY NC.NCC_FECHA"
         
         rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Dim subtotal As Double
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
'                    If rec!TCO_CODIGO = 4 Then ' NC A
                        sql = sql & XN(CStr((-1) * CDbl(IIf(IsNull(rec!NCC_SUBTOTAL), 0, rec!NCC_SUBTOTAL)))) & ","
                        sql = sql & XN(rec!NCC_IVA) & ","
                        TotIva = (CDbl(IIf(IsNull(rec!NCC_SUBTOTAL), 0, rec!NCC_SUBTOTAL)) * CDbl(rec!NCC_IVA)) / 100
                        sql = sql & XN(CStr((-1) * CDbl(TotIva))) & ","
                        sql = sql & XN(CStr((-1) * CDbl(IIf(IsNull(rec!NCC_TOTAL), 0, rec!NCC_TOTAL)))) & ")"
'                    Else
'                        subtotal = CDbl(rec!NCC_TOTAL / (1 + (rec!NCC_IVA / 100)))
'                        sql = sql & XN(CStr((-1) * subtotal)) & ","
'                        sql = sql & XN(rec!NCC_IVA) & ","
'                        TotIva = CDbl(rec!NCC_TOTAL - subtotal)
'                        sql = sql & XN(CStr((-1) * CDbl(TotIva))) & ","
'                        sql = sql & XN(CStr((-1) * CDbl(IIf(IsNull(rec!NCC_TOTAL), 0, rec!NCC_TOTAL)))) & ")"
'                    End If
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
        If Not IsNull(FechaDesde) Then sql = sql & " AND ND.NDC_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND ND.NDC_FECHA<=" & XDQ(FechaHasta)
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
    'ListarLibroIVA
        
    Screen.MousePointer = vbNormal
    
    Exit Function

CLAVO:
 Screen.MousePointer = vbNormal
 lblEstado.Caption = ""
 DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Function

Private Sub ListarLibroIVA()
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
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
    Dim NroCbteHasta As String
    Dim NroDespacho As String
    Dim CodDocProv As String
    Dim NroIdVend As String
    Dim NombreProv As String
    Dim TOTALAux As String
    Dim TOTAL As String
    Dim NoNetoGravado As String
    Dim PerNoCat As String
    Dim OpExentas As String
    Dim ImpPerPagImp As String
    Dim PerIIBB As String
    Dim ImpMuni As String
    Dim ImpInt As String
    Dim Moneda As String
    Dim TCambio As String
    Dim CantIVA As String
    Dim CodOp As String
    Dim OtrosTrib As String
    Dim FecVen As String
    
    
    Dim CuitEmisor As String
    Dim DenEmisor As String
    Dim IVAcom As String
    Dim MesAnio As String
    Dim ImpLiq As String
    Dim ImpLiqaux As String
    
    Dim I As Integer
    Dim cantRegistros As Integer
    sql = "SELECT * FROM TMP_LIBRO_IVA_VENTAS WHERE TOTAL <> 0"
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
            
            '--------Numero de Comprobante Hasta
'            If rec!COMPROBANTE = "FAC-B" Then
'                NroCbteHasta = Right(rec!Numero, 8)
'
'            Else
'                NroCbteHasta = 0
'            End If
'            NroCbteHasta = String(20 - Len(NroCbteHasta), "0") & NroCbteHasta
            
            'Codigo de documento del Comprador
            CodDocProv = 80 'CUIT
            
             'Numero de Identificacion del Comprador
            NroIdVend = Replace(rec!cuit, "-", "")
            NroIdVend = String(20 - Len(NroIdVend), "0") & NroIdVend
                       
            'Apellido y Nombre del vendedor
            NombreProv = Left(rec!Cliente, 30)
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
            NoNetoGravado = 0
            NoNetoGravado = String(15 - Len(NoNetoGravado), "0") & NoNetoGravado
            
            'Percepcion a no categorizados
            PerNoCat = 0
            PerNoCat = String(15 - Len(PerNoCat), "0") & PerNoCat
                      
            'Importe de operaciones exentas
            If rec!cuit = "20-16837135-8" Then 'Arbore alquiler
                OpExentas = TOTAL
                
            Else
                OpExentas = 0
            End If
            OpExentas = String(15 - Len(OpExentas), "0") & OpExentas
            
            'Importe de percepciones o pagos a cuenta de impuestos nacionales
            ImpPerPagImp = 0
            ImpPerPagImp = String(15 - Len(ImpPerPagImp), "0") & ImpPerPagImp
           
            
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
            CantIVA = 1
            CantIVA = String(1 - Len(CantIVA), "0") & CantIVA
            
            'Codigo de Operacion
            CodOp = 0
            If rec!cuit = "20-16837135-8" Then
                    CodOp = "E"
            End If
            CodOp = String(1 - Len(CodOp), "0") & CodOp
            
                        
            'Otros Tributos
            OtrosTrib = 0
            OtrosTrib = String(15 - Len(OtrosTrib), "0") & OtrosTrib
            
            'Fecha de vencimiento de pago
            FecVen = "0"
            FecVen = String(8 - Len(FecVen), "0") & FecVen
 
            
            
            '----------archivo alicuotas iva--------------
            Dim neto As String
            Dim netoaux As String
            Dim alicuotaIVA As String
            
            'importe neto gravado
            'Importe Total de la operacion
            If CodOp = "E" Then
                neto = 0
            Else
                If rec!subtotal < 0 Then
                    netoaux = Replace(Format(rec!subtotal, "#,##0.00"), ",", "")
                    neto = Replace(netoaux, "-", "0")
                    neto = Replace(neto, ".", "")
                Else
                    neto = Replace(Format(rec!subtotal, "#,##0.00"), ",", "")
                    neto = Replace(neto, ".", "")
                End If
            End If
            neto = String(15 - Len(neto), "0") & neto
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
                
            'Impuesto Liquidado
            If rec!subtotal < 0 Then
                ImpLiqaux = Replace(Format(rec!TotIva, "#,##0.00"), ",", "")
                ImpLiq = Replace(ImpLiqaux, "-", "0")
                ImpLiq = Replace(ImpLiq, ".", "")
            Else
                ImpLiq = Replace(Format(rec!TotIva, "#,##0.00"), ",", "")
                ImpLiq = Replace(ImpLiq, ".", "")
                
            End If
                        
            ImpLiq = String(15 - Len(ImpLiq), "0") & ImpLiq
                
           
            'ARMO UNA LINEA DEL ARCHIVO COMPROBANTES
            Cadena(I) = Fecha & TipoCbte & PtoVenta & NroCbte & NroCbte & _
                        CodDocProv & NroIdVend & NombreProv & TOTAL & NoNetoGravado & _
                        PerNoCat & OpExentas & ImpPerPagImp & PerIIBB & ImpMuni & ImpInt & Moneda & _
                        TCambio & CantIVA & CodOp & OtrosTrib & FecVen
            
            
            'ARMO UNA LINEA DEL ARCHIVO ALICUOTAS
                Alicuota(I) = TipoCbte & PtoVenta & NroCbte & _
                             neto & alicuotaIVA & ImpLiq
            
            
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
        
    If EstadoDeArchivo(DirAFIP & "Ventas_AFIP_" & MesAnio & ".txt") Then
        Kill (DirAFIP & "Ventas_AFIP_" & MesAnio & ".txt")
    End If
    If EstadoDeArchivo(DirAFIP & "AlicuotasVentas_AFIP_" & MesAnio & ".txt") Then
        Kill (DirAFIP & "AlicuotasVentas_AFIP_" & MesAnio & ".txt")
    End If
    
    'GENERO LOS ARCHIVOS
    For I = 1 To cantRegistros
        Open DirAFIP & "Ventas_AFIP_" & MesAnio & ".txt" For Append As #1
        Print #1, Cadena(I)
        Close #1
    Next
    MsgBox "Se genero correctamente el archivo " & DirAFIP & "Ventas_AFIP_" & MesAnio & ".txt", vbInformation, TIT_MSGBOX
    
    
    For I = 1 To cantRegistros
        Open DirAFIP & "AlicuotasVentas_AFIP_" & MesAnio & ".txt" For Append As #1
        
        Print #1, Alicuota(I)
        Close #1
        
    Next
    MsgBox "Se genero correctamente el archivo " & DirAFIP & "AlicuotasVentas_AFIP_" & MesAnio & ".txt", vbInformation, TIT_MSGBOX
    
End Function
Private Sub CmdNuevo_Click()
    FechaDesde.Value = Null
    lblPeriodo1.Caption = ""
    FechaHasta.Value = Null
    lblPeriodo2.Caption = ""
    FechaDesde.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmLibroVentas2 = Nothing
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
Public Function EstadoDeArchivo(ByVal Archivo As String) As Boolean
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If (fso.FileExists(Archivo)) Then
        EstadoDeArchivo = True
    Else
        EstadoDeArchivo = False
    End If
End Function
