VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmControlStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Stock"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabLista 
      Height          =   855
      Left            =   7320
      TabIndex        =   34
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1508
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Accesorios"
      TabPicture(0)   =   "frmControlStock.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cboListaPrecio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Repuestos"
      TabPicture(1)   =   "frmControlStock.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboLPrecioRep"
      Tab(1).ControlCount=   1
      Begin VB.ComboBox cboLPrecioRep 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   420
         Width           =   3345
      End
      Begin VB.ComboBox cboListaPrecio 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   420
         Width           =   3345
      End
   End
   Begin VB.CheckBox chkLista 
      Caption         =   "Lista de Precios:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   37
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame freOpciones 
      Caption         =   "Opciones de Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   50
      TabIndex        =   24
      Top             =   960
      Width           =   11075
      Begin VB.CheckBox chkProducto 
         Caption         =   "Producto"
         Height          =   195
         Left            =   360
         TabIndex        =   33
         Top             =   285
         Width           =   960
      End
      Begin VB.TextBox txtProducto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1485
         TabIndex        =   32
         Top             =   225
         Width           =   3165
      End
      Begin VB.CheckBox chkRepres 
         Caption         =   "Marca"
         Height          =   255
         Left            =   5295
         TabIndex        =   31
         Top             =   615
         Width           =   780
      End
      Begin VB.ComboBox cborubro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   225
         Width           =   3870
      End
      Begin VB.ComboBox cbolinea 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1485
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   585
         Width           =   3180
      End
      Begin VB.ComboBox cboRep 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6165
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   585
         Width           =   3870
      End
      Begin VB.CommandButton CmdBuscAprox 
         Height          =   570
         Left            =   10125
         Picture         =   "frmControlStock.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   225
         Width           =   690
      End
      Begin VB.CheckBox chklinea 
         Caption         =   "Línea"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   615
         Width           =   780
      End
      Begin VB.CheckBox chkrubro 
         Caption         =   "Rubro"
         Height          =   285
         Left            =   5295
         TabIndex        =   25
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame frastock 
      Caption         =   "Actualizción de Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   4200
      TabIndex        =   21
      Top             =   3000
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   495
         Left            =   960
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdsalgo 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtNuevoStock 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtStockA 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Stock:"
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   780
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Stock Actual:"
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   300
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gráficar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   6765
      TabIndex        =   18
      Top             =   5625
      Width           =   4290
      Begin VB.CommandButton cmdGrafico 
         Caption         =   "&Gráfico"
         Height          =   750
         Left            =   3240
         Picture         =   "frmControlStock.frx":27DA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   855
      End
      Begin VB.ComboBox cboTipoGrafico 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   525
         Width           =   2355
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Gráfico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   19
         Top             =   285
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   8415
      Picture         =   "frmControlStock.frx":2AE4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6690
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   10155
      Picture         =   "frmControlStock.frx":33AE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6690
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmControlStock.frx":36B8
      Height          =   750
      Left            =   9285
      Picture         =   "frmControlStock.frx":39C2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6690
      Width           =   870
   End
   Begin VB.ComboBox cboStock 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   2685
   End
   Begin VB.Frame Frame3 
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
      Height          =   1800
      Left            =   15
      TabIndex        =   13
      Top             =   5625
      Width           =   6705
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   435
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   660
         Width           =   1665
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   6
         Top             =   360
         Width           =   1050
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
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label lblImpresora 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1800
         TabIndex        =   16
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   360
         Width           =   585
      End
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   6795
      Top             =   6870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   7380
      Top             =   6930
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3675
      Left            =   45
      TabIndex        =   12
      Top             =   1935
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   6482
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorSel    =   8388736
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione el Stock:"
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
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   2145
   End
End
Attribute VB_Name = "frmControlStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub chklinea_Click()
    If chklinea.Value = 1 Then
        cbolinea.Enabled = True
        cbolinea.ListIndex = 0
        cbolinea.SetFocus
    Else
         cbolinea.Enabled = False
        cbolinea.ListIndex = -1
    End If
End Sub

Private Sub chkLista_Click()
    If chkLista.Value = Checked Then
        tabLista.Enabled = True
        If cboListaPrecio.ListCount > 1 Then
            cboListaPrecio.ListIndex = 0
        End If
        If cboLPrecioRep.ListCount > 1 Then
            cboLPrecioRep.ListIndex = 0
        End If
    Else
        tabLista.Enabled = False
        cboListaPrecio.ListIndex = -1
        cboLPrecioRep.ListIndex = -1
    End If
    
End Sub

Private Sub chkProducto_Click()
    If chkProducto.Value = 1 Then
        txtProducto.Enabled = True
        txtProducto.SetFocus
    Else
        txtProducto.Enabled = False
        txtProducto.Text = ""
    End If
End Sub

Private Sub chkRepres_Click()
    If (chklinea.Value = 1) And (chkrubro.Value = 1) Then
        cargocboRepres cbolinea.ItemData(cbolinea.ListIndex), cborubro.ItemData(cborubro.ListIndex)
    Else
        If chklinea.Value = 1 Then
            cargocboRepres cbolinea.ItemData(cbolinea.ListIndex), -1
        Else
            If chkrubro.Value = 1 Then
                cargocboRepres -1, cborubro.ItemData(cborubro.ListIndex)
            Else
                    cargocboRepres -1, -1
            End If
        End If
        
    End If
    If chkRepres.Value = 1 Then
        cboRep.Enabled = True
'        cboRep.ListIndex = 0
        cboRep.SetFocus
    Else
        cboRep.Enabled = False
        cboRep.ListIndex = -1
    End If
End Sub

Private Sub chkRubro_Click()
    If chklinea.Value = 0 Then
        cargocboRubro (-1)
    Else
        cargocboRubro (cbolinea.ItemData(cbolinea.ListIndex))
    End If
    If chkrubro.Value = 1 Then
        cborubro.Enabled = True
        cborubro.ListIndex = 0
        cborubro.SetFocus
    Else
        cborubro.Enabled = False
        cborubro.ListIndex = -1
    End If
    
End Sub
'Private Sub chkLinea_Click()
'    If chkLinea.Value = 1 Then
'        txtLinea.Enabled = True
'        txtLinea.SetFocus
'    Else
'        txtLinea.Text = ""
'        txtDesLin.Text = ""
'        txtLinea.Enabled = False
'    End If
'
'End Sub
'
'Private Sub chkLinea_GotFocus()
'    SelecTexto txtLinea
'End Sub
'
'Private Sub chkLinea_KeyPress(KeyAscii As Integer)
'    KeyAscii = CarNumeroEntero(KeyAscii)
'End Sub

'Private Sub chkProducto_Click()
'    If chkProducto.Value = 1 Then
'        txtProducto.Enabled = True
'        txtProducto.SetFocus
'    Else
'        txtProducto.Text = ""
'        txtDesProd.Text = ""
'        txtProducto.Enabled = False
'    End If
'
'    End Sub
'
'Private Sub chkRepresentada_Click()
'    If chkRepresentada.Value = 1 Then
'        txtRepresentada.Enabled = True
'        txtRepresentada.SetFocus
'    Else
'        txtRepresentada.Text = ""
'        txtDesRep.Text = ""
'        txtRepresentada.Enabled = False
'
'    End If
'End Sub
'
'Private Sub chkRubro_Click()
'    If chkRubro.Value = 1 Then
'        txtRubro.Enabled = True
'        txtRubro.SetFocus
'    Else
'        txtRubro.Text = ""
'        txtDesRub.Text = ""
'        txtRubro.Enabled = False
'
'    End If
'End Sub

Private Sub CmdAceptar_Click()
    Dim resp As String
    resp = MsgBox("Seguro desea Actualizar el stock del Producto: " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & "? ", 36, "Actualizar Stock")
    If resp <> 6 Then Exit Sub
    
    sql = "UPDATE DETALLE_STOCK"
    sql = sql & " SET"
    sql = sql & " DST_STKFIS = " & XN(txtNuevoStock.Text)
    sql = sql & " WHERE STK_CODIGO = 1"
    '& cboStock.ItemData(cboStock.ListIndex)
    sql = sql & " AND PTO_CODIGO LIKE '" & GrdModulos.TextMatrix(GrdModulos.RowSel, 0) & "' "
    DBConn.Execute sql
    
    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = txtNuevoStock.Text
    frastock.Visible = False
    GrdModulos.SetFocus

End Sub

Private Sub CmdBuscAprox_Click()
    Dim j As Integer
    
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    sql = "SELECT DISTINCT   P.PTO_CODIGO,P.PTO_DESCRI,L.LNA_DESCRI,R.RUB_DESCRI,TP.TPRE_DESCRI,"
    sql = sql & " P.PTO_STKMIN,D.DST_STKFIS"
    sql = sql & " FROM PRODUCTO P,LINEAS L,RUBROS R, REPRESENTADA RE,DETALLE_STOCK D,TIPO_PRESENTACION TP"
    sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO "
    sql = sql & " AND L.LNA_CODIGO = P.LNA_CODIGO "
    sql = sql & " AND P.TPRE_CODIGO = TP.TPRE_CODIGO "
    sql = sql & " AND R.RUB_CODIGO = P.RUB_CODIGO  "
    sql = sql & " AND L.LNA_CODIGO = R.LNA_CODIGO "
    
    'AND P.REP_CODIGO = RE.REP_CODIGO "
    sql = sql & " AND D.STK_CODIGO = 1 "
    '" & cboStock.ItemData(cboStock.ListIndex) & " "
    If txtProducto.Text <> "" Then
        sql = sql & " AND (P.PTO_CODIGO LIKE '" & txtProducto.Text & "%' "
        sql = sql & " OR P.PTO_DESCRI LIKE '" & txtProducto.Text & "%')"
    End If
    If chklinea.Value = 1 Then sql = sql & " AND L.LNA_CODIGO=" & cbolinea.ItemData(cbolinea.ListIndex)
    If chkrubro.Value = 1 Then sql = sql & "AND R.RUB_CODIGO=" & cborubro.ItemData(cborubro.ListIndex)
    If chkRepres = 1 Then sql = sql & " AND TP.TPRE_CODIGO=" & cboRep.ItemData(cboRep.ListIndex)
    'Muestro de acuerdo a la Lista de Precio
    If chkLista.Value = Checked Then
        If tabLista.Tab = 0 Then
            sql = sql & "AND P.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
        Else
            sql = sql & "AND P.LIS_CODIGO=" & cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex)
        End If
        
    End If
    
    'sql = sql & " GROUP BY P.PTO_CODIGO,P.PTO_DESCRI,L.LNA_DESCRITP.TPRE_DESCRI,D.DST_STKPEN,D.DST_STKFIS"
    sql = sql & " ORDER BY P.PTO_CODIGO"
    'ES PARA MOSTRAR O NO EL STOCK DE TOTALCAR
    If cboStock.ItemData(cboStock.ListIndex) = 1 Then
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        GrdModulos.ColWidth(1) = 3000
        GrdModulos.ColWidth(7) = 0
        If rec.EOF = False Then
            GrdModulos.Rows = 1
            Do While Not rec.EOF
            GrdModulos.AddItem rec!PTO_CODIGO & Chr(9) & rec!PTO_DESCRI & Chr(9) & _
                               rec!LNA_DESCRI & Chr(9) & rec!RUB_DESCRI & Chr(9) & _
                               rec!TPRE_DESCRI & Chr(9) & rec!PTO_STKMIN & Chr(9) & _
                               IIf(IsNull(rec!DST_STKFIS), "0", rec!DST_STKFIS)
                              ' "0" & Chr(9) & _
                              ' IIf(IsNull(rec!DST_STKPEN), "0", rec!DST_STKPEN) & Chr(9) & _
                              ' IIf(IsNull(rec!DST_STKFIS), "0", rec!DST_STKFIS) & Chr(9) & _
                              ' IIf(IsNull(rec!DST_STKPEN), rec!DST_STKFIS, (rec!DST_STKFIS - rec!DST_STKPEN))
    
            rec.MoveNext
            Loop
        Else
            MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
            GrdModulos.Rows = 1
        End If
        rec.Close
        Screen.MousePointer = vbNormal
        lblEstado.Caption = ""
    Else 'ACA ENTRA CUANDO NO ES COVIT
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        GrdModulos.ColWidth(1) = 2500
'        GrdModulos.ColWidth(7) = 1000
'        If rec.EOF = False Then
'            GrdModulos.Rows = 1
'            Do While Not rec.EOF
'            'ACA VER BIEN LA CANTIDAD DE COLUMNAS
'            GrdModulos.AddItem rec!PTO_CODIGO & Chr(9) & rec!PTO_DESCRI & Chr(9) & rec!TPRE_DESCRI & Chr(9) & _
'                              "0" & Chr(9) & _
'                              IIf(IsNull(rec!DST_STKPEN), "0", rec!DST_STKPEN) & Chr(9) & _
'                              IIf(IsNull(rec!DST_STKFIS), "0", rec!DST_STKFIS) & Chr(9) & _
'                              IIf(IsNull(rec!DST_STKPEN), IIf(IsNull(rec!DST_STKFIS), "0", rec!DST_STKFIS), (rec!DST_STKFIS - rec!DST_STKPEN))
'
'            rec.MoveNext
'            Loop
'        Else
'            MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
'            GrdModulos.Rows = 1
'        End If
        rec.Close
        Screen.MousePointer = vbNormal
        lblEstado.Caption = ""
        
'        If GrdModulos.Rows > 1 Then
'
'            'select para ver los stockes de COVIT
'            sql = "SELECT DISTINCT D.PTO_CODIGO,D.DST_STKPEN,D.DST_STKFIS"
'            sql = sql & " FROM DETALLE_STOCK D"
'            sql = sql & " WHERE D.STK_CODIGO = 1 "
'            sql = sql & " ORDER BY D.PTO_CODIGO"
'            lblEstado.Caption = "Buscando..."
'            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'            If rec.EOF = False Then
'            Do While Not rec.EOF
'                For j = 1 To GrdModulos.Rows - 1
'                    If rec!PTO_CODIGO = GrdModulos.TextMatrix(j, 0) Then
'                        If IsNull(rec!DST_STKFIS) Then
'                            GrdModulos.TextMatrix(j, 7) = "0"
'                        Else
'                            GrdModulos.TextMatrix(j, 7) = IIf(IsNull(rec!DST_STKPEN), rec!DST_STKFIS, (rec!DST_STKFIS - rec!DST_STKPEN))
'                        End If
'                        Exit For
'                    End If
'                Next
'                rec.MoveNext
'                Loop
'            End If
'            rec.Close
'        End If
        
    End If
'    If GrdModulos.Rows > 1 Then 'PREGUNTO SI LA GRILLA NO ESTA VACIA
'        'Aca calculo y agrego los pedidos pendientes en la grilla
'        sql = "SELECT DISTINCT DR.PTO_CODIGO,SUM(DR.DRC_CANTIDAD) AS PEDPEN"
'        sql = sql & " FROM DETALLE_REMITO_CLIENTE DR, REMITO_CLIENTE RC"
'        sql = sql & " WHERE RC.RCL_NUMERO = DR.RCL_NUMERO AND RC.EST_CODIGO = 1"
'        sql = sql & " GROUP BY DR.PTO_CODIGO"
'        sql = sql & " ORDER BY DR.PTO_CODIGO"
'
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'        If rec.EOF = False Then
'            Do While Not rec.EOF
'                For j = 1 To GrdModulos.Rows - 1
'                    If rec!PTO_CODIGO = GrdModulos.TextMatrix(j, 0) Then
'                       GrdModulos.TextMatrix(j, 3) = rec!PEDPEN
'                       Exit For
'                    End If
'                Next
'            rec.MoveNext
'            Loop
'        End If
'        rec.Close
'    End If
    lblEstado.Caption = ""
    If Stock = 0 Then
        If GrdModulos.Rows > 1 Then
            GrdModulos.ToolTipText = "Doble Click en la Grilla para actualizar el Stock"
        Else
            GrdModulos.ToolTipText = ""
        End If
    Else
        frmControlStock.Caption = "Consulta del Stock"
        GrdModulos.ToolTipText = ""
    End If
    
End Sub

Private Sub cmdBuscarLin_Click()

End Sub

Private Sub cmdBuscarProd_Click()
'    frmBuscar.TipoBusqueda = 2
'    frmBuscar.Show vbModal
'    If frmBuscar.grdBuscar.Text <> "" Then
'        frmBuscar.grdBuscar.Col = 0
'        txtproducto.Text = frmBuscar.grdBuscar.Text
'        frmBuscar.grdBuscar.Col = 1
'        txtDesProd.Text = frmBuscar.grdBuscar.Text
'    Else
'        txtproducto.Enabled = True
'        txtproducto.SetFocus
'    End If
End Sub

Private Sub cmdGrafico_Click()
    If GrdModulos.Rows > 1 Then
        frmGraficoStock.Show vbModal
    Else
        MsgBox "No hay ningún producto seleccionado", vbExclamation, TIT_MSGBOX
    End If
End Sub

Private Sub cmdListar_Click()
    
    lblEstado.Caption = "Buscando Listado..."
    Screen.MousePointer = vbHourglass
    
    Rep.WindowState = crptMaximized 'crptMinimized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    
    'If cboStock.ItemData(cboStock.ListIndex) = 1 Then
    '    Rep.SelectionFormula = "{STOCK.STK_CODIGO}=" & cboStock.ItemData(cboStock.ListIndex)
    '    Rep.ReportFileName = DRIVE & DirReport & "rptstockcovit.rpt"
    'Else
    LlenoTablaTemporal
    Rep.SelectionFormula = "{STOCK.STK_CODIGO}=" & cboStock.ItemData(cboStock.ListIndex)
    Rep.ReportFileName = DRIVE & DirReport & "rptstockotros.rpt"
    'End If
    Screen.MousePointer = vbNormal
    Rep.Destination = crptToWindow
    Rep.Action = 1
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    lblEstado.Caption = ""
    
End Sub

Private Sub LlenoTablaTemporal()
    Dim CantidadStock As String
    CantidadStock = ""
  
    sql = "DELETE FROM TMP_LISTADO_DETALLE_STOCK"
    DBConn.Execute sql
    
    sql = "INSERT INTO TMP_LISTADO_DETALLE_STOCK"
    sql = sql & " (STK_CODIGO,PTO_CODIGO,DST_STKFIS,DST_STKPEN)"
    sql = sql & " SELECT DISTINCT  D.STK_CODIGO, P.PTO_CODIGO, D.DST_STKFIS,P.PTO_STKMIN"
    sql = sql & " FROM PRODUCTO P,LINEAS L,RUBROS R, REPRESENTADA RE,DETALLE_STOCK D,TIPO_PRESENTACION TP"
    sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO "
    sql = sql & " AND L.LNA_CODIGO = P.LNA_CODIGO "
    sql = sql & " AND P.TPRE_CODIGO = TP.TPRE_CODIGO "
    sql = sql & " AND R.RUB_CODIGO = P.RUB_CODIGO  "
    sql = sql & " AND L.LNA_CODIGO = R.LNA_CODIGO "
    
    sql = sql & " AND D.STK_CODIGO = 1 "  ' " & cboStock.ItemData(cboStock.ListIndex)
    If txtProducto.Text <> "" Then
        sql = sql & " AND (P.PTO_CODIGO LIKE '" & txtProducto.Text & "%' "
        sql = sql & " OR P.PTO_DESCRI LIKE '" & txtProducto.Text & "%')"
    End If
    If chklinea.Value = 1 Then sql = sql & " AND L.LNA_CODIGO=" & cbolinea.ItemData(cbolinea.ListIndex)
    If chkrubro.Value = 1 Then sql = sql & " AND R.RUB_CODIGO=" & cborubro.ItemData(cborubro.ListIndex)
    If chkRepres = 1 Then sql = sql & " AND TP.TPRE_CODIGO=" & cboRep.ItemData(cboRep.ListIndex)
    
    DBConn.Execute sql
    
    'select para ver el stock de COVIT
'    sql = "SELECT DISTINCT D.PTO_CODIGO,D.DST_STKPEN,D.DST_STKFIS"
'    sql = sql & " FROM DETALLE_STOCK D"
'    sql = sql & " WHERE D.STK_CODIGO = 1 "
'    sql = sql & " ORDER BY D.PTO_CODIGO"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While Not rec.EOF
'            CantidadStock = IIf(IsNull(rec!DST_STKPEN), rec!DST_STKFIS, (rec!DST_STKFIS - rec!DST_STKPEN))
'            sql = "UPDATE TMP_LISTADO_DETALLE_STOCK"
'            sql = sql & " SET DST_COVIT=" & XN(CantidadStock)
'            sql = sql & " WHERE PTO_CODIGO=" & XN(rec!PTO_CODIGO)
'            DBConn.Execute sql
'            rec.MoveNext
'        Loop
'    End If
'    rec.Close
End Sub

Private Sub CmdNuevo_Click()
    cboStock.SetFocus
    HacerNuevo 'Funcion que limpia los controles al presionar boton Nuevo
    CmdBuscAprox_Click
End Sub
Private Sub HacerNuevo()
    chklinea.Value = 0
    chkrubro.Value = 0
    chkRepres.Value = 0
    chkProducto.Value = 0
    txtProducto.Text = ""
'    txtLinea.Text = ""
'    txtRubro.Text = ""
'    txtRepresentada.Text = ""
'    txtDesLin.Text = ""
'    txtDesProd.Text = ""
'    txtDesRub.Text = ""
'    txtDesRep.Text = ""
End Sub

Private Sub cmdsalgo_Click()
    frastock.Visible = False
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
Set rec = New ADODB.Recordset

Call Centrar_pantalla(Me)
CargocboStock
preparogrilla
'cobmo tipo gráfico----------------
cboTipoGrafico.AddItem "Gráfico 2D"
cboTipoGrafico.AddItem "Gráfico 3D"
cboTipoGrafico.ListIndex = 0
'----------------------------------
'cargo combos
If Stock = 0 Then
    frmControlStock.Caption = "Actualización del Stock"
    If GrdModulos.Rows > 1 Then
        GrdModulos.ToolTipText = "Doble Click en la Grilla para actualizar el Stock"
    Else
        GrdModulos.ToolTipText = ""
    End If
Else
    frmControlStock.Caption = "Consulta del Stock"
    GrdModulos.ToolTipText = ""
End If
cargocboLinea
cargocboRubro (-1)
cargocboRepres -1, -1  ' Para Cargar Marcas sin Lineas y Rubros

chkLista.Value = Checked
tabLista.Tab = 0
CargoCboLPrecioRep
CargoCboListaPrecio

lblEstado.Caption = ""
lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
'CmdBuscAprox_Click
End Sub
Private Sub CargoCboListaPrecio() '' Lista de Precios de Accesorios
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 2"   '6: Accesorios
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboListaPrecio.AddItem rec!LIS_DESCRI
            cboListaPrecio.ItemData(cboListaPrecio.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboListaPrecio.ListIndex = 0
    End If
    rec.Close
End Sub
Private Sub CargoCboLPrecioRep() '' Lista de Precios de Repuestos
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 1"   '1: Repuestos
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboLPrecioRep.AddItem rec!LIS_DESCRI
            cboLPrecioRep.ItemData(cboLPrecioRep.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboLPrecioRep.ListIndex = 0
    End If
    rec.Close
End Sub
Function cargocboLinea()
    cbolinea.Clear
    sql = "SELECT * FROM LINEAS WHERE LNA_CODIGO <> 8 ORDER BY LNA_DESCRI"
    If rec.State = 1 Then rec.Close
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cbolinea.AddItem rec!LNA_DESCRI
            cbolinea.ItemData(cbolinea.NewIndex) = rec!LNA_CODIGO
            rec.MoveNext
        Loop
        cbolinea.ListIndex = -1
    End If
    rec.Close
End Function
Function cargocboRubro(cod As Integer)
    
    cborubro.Clear
    sql = "SELECT * FROM RUBROS "
    sql = sql & "WHERE RUB_CODIGO <> 34"
    If cod <> -1 Then
        sql = sql & " AND LNA_CODIGO= " & cod
    End If
    sql = sql & " ORDER BY RUB_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cborubro.AddItem rec!RUB_DESCRI
            cborubro.ItemData(cborubro.NewIndex) = rec!RUB_CODIGO
            rec.MoveNext
        Loop
        cborubro.ListIndex = -1
    End If
    rec.Close
End Function
Function cargocboRepres(codL As Integer, codR As Integer)
    cboRep.Clear
    sql = "SELECT * FROM TIPO_PRESENTACION WHERE TPRE_CODIGO <> 0 AND TPRE_CODIGO <> 60"
    If codL <> -1 Then
        sql = sql & " AND LNA_CODIGO = " & cbolinea.ItemData(cbolinea.ListIndex) & ""
    End If
    If codR <> -1 Then
        sql = sql & "AND RUB_CODIGO = " & cborubro.ItemData(cborubro.ListIndex) & ""
    End If
    sql = sql & " ORDER BY TPRE_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboRep.AddItem rec!TPRE_DESCRI
            cboRep.ItemData(cboRep.NewIndex) = rec!TPRE_CODIGO
            rec.MoveNext
        Loop
        cboRep.ListIndex = -1
    End If
    rec.Close
End Function
Function preparogrilla()
    GrdModulos.FormatString = "Código|Producto|Linea|Rubro|Marca|Stock Min|Stock Actual|COVIT"
    GrdModulos.ColWidth(0) = 1000 'codigo
    GrdModulos.ColWidth(1) = 3000 'producto
    GrdModulos.ColWidth(2) = 1500 'linea
    GrdModulos.ColWidth(3) = 1500 'rubro.
    GrdModulos.ColWidth(4) = 1500 'Marca
    GrdModulos.ColWidth(5) = 1000 'Stock Min
    GrdModulos.ColWidth(6) = 1200 'Stock Actual
    GrdModulos.ColWidth(7) = 0    'COVIT
    GrdModulos.Rows = 1
End Function
Private Sub CargocboStock()
    sql = "SELECT STK_CODIGO, REP_RAZSOC "
    sql = sql & " FROM STOCK S, REPRESENTADA R "
    sql = sql & " WHERE S.REP_CODIGO=R.REP_CODIGO"
    sql = sql & " ORDER BY STK_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboStock.AddItem rec!REP_RAZSOC
            cboStock.ItemData(cboStock.NewIndex) = rec!STK_CODIGO
            rec.MoveNext
        Loop
        cboStock.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub Text1_Change()

End Sub

Private Sub GrdModulos_DblClick()
    'variable de stock = 1 modificacion
    If GrdModulos.Rows > 1 Then
        If Stock = 0 Then
            frastock.Visible = True
            txtStockA.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
            txtNuevoStock.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
            txtNuevoStock.SetFocus
        End If
    End If
End Sub

Private Sub GrdModulos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub txtLinea_Change()
'    If txtLinea.Text = "" Then
'        txtLinea.Text = ""
'        txtDesLin.Text = ""
'    End If
    
End Sub

Private Sub txtLinea_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtLinea_LostFocus()
'    If txtLinea.Text <> "" Then
'        sql = "SELECT L.LNA_CODIGO,L.LNA_DESCRI"
'        sql = sql & " FROM LINEAS L,PRODUCTO P,DETALLE_STOCK D"
'        sql = sql & " WHERE L.LNA_CODIGO = P.LNA_CODIGO "
'        sql = sql & " AND P.PTO_CODIGO = D.PTO_CODIGO "
'        sql = sql & " AND D.STK_CODIGO = " & cboStock.ItemData(cboStock.ListIndex) & "  "
'        sql = sql & " AND L.LNA_CODIGO = " & CInt(txtLinea.Text) & ""
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            txtDesLin.Text = rec!LNA_DESCRI
'            CmdBuscAprox.SetFocus
'        Else
'            MsgBox "El codigo no existe", vbInformation
'            txtLinea.SetFocus
'        End If
'        rec.Close
'     End If
End Sub

Private Sub txtNuevoStock_GotFocus()
    SelecTexto txtNuevoStock
End Sub

Private Sub txtNuevoStock_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtNuevoStock, KeyAscii)
End Sub

Private Sub txtProducto_Change()
'    If txtProducto.Text = "" Then
'        txtProducto.Text = ""
'        'txtDesProd.Text = ""
'
'    End If
End Sub

Private Sub txtProducto_GotFocus()
    SelecTexto txtProducto
    
End Sub

Private Sub txtproducto_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtProducto_LostFocus()
'    If txtProducto.Text <> "" Then
'        sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI"
'        sql = sql & " FROM PRODUCTO P, DETALLE_STOCK D "
'        sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO"
'        sql = sql & " AND D.STK_CODIGO = " & cboStock.ItemData(cboStock.ListIndex) & "  "
'        sql = sql & " AND P.PTO_CODIGO = " & CInt(txtProducto.Text) & ""
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'        If rec.EOF = False Then
'            'txtDesProd.Text = rec!PTO_DESCRI
'            CmdBuscAprox.SetFocus
'        Else
'            MsgBox "El código no existe", vbInformation
'            txtProducto.SetFocus
'        End If
'        rec.Close
'    End If
    
End Sub

Private Sub txtRepresentada_GotFocus()
    'SelecTexto txtRepresentada
End Sub

Private Sub txtRepresentada_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtRepresentada_LostFocus()
'    If txtRepresentada.Text <> "" Then
'        sql = "SELECT R.REP_CODIGO,R.REP_RAZSOC"
'        sql = sql & " FROM REPRESENTADA R,PRODUCTO P,DETALLE_STOCK D"
'        sql = sql & " WHERE R.REP_CODIGO = P.REP_CODIGO "
'        sql = sql & " AND P.PTO_CODIGO = D.PTO_CODIGO "
'        sql = sql & " AND D.STK_CODIGO = " & cboStock.ItemData(cboStock.ListIndex) & "  "
'        sql = sql & " AND R.REP_CODIGO = " & CInt(txtRepresentada.Text) & ""
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            txtDesRep.Text = rec!REP_RAZSOC
'            CmdBuscAprox.SetFocus
'        Else
'            MsgBox "El codigo no existe", vbInformation
'            txtRepresentada.SetFocus
'        End If
'        rec.Close
'     End If
End Sub

Private Sub txtRubro_Change()
'    If txtRubro.Text = "" Then
'        txtRubro.Text = ""
'        txtDesRub.Text = ""
'    End If
End Sub

Private Sub txtRubro_GotFocus()
'    SelecTexto txtRubro
End Sub

Private Sub txtRubro_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtRubro_LostFocus()
'    If txtRubro.Text <> "" Then
'        sql = "SELECT R.RUB_CODIGO,R.RUB_DESCRI"
'        sql = sql & " FROM RUBROS R,PRODUCTO P,DETALLE_STOCK D"
'        sql = sql & " WHERE R.RUB_CODIGO = P.RUB_CODIGO "
'        sql = sql & " AND P.PTO_CODIGO = D.PTO_CODIGO "
'        sql = sql & " AND D.STK_CODIGO = " & cboStock.ItemData(cboStock.ListIndex) & "  "
'        sql = sql & " AND R.RUB_CODIGO = " & CInt(txtRubro.Text) & ""
'        If txtLinea.Text <> "" Then
'            sql = sql & "AND R.LNA_CODIGO = " & CInt(txtLinea.Text) & ""
'        End If
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            txtDesRub.Text = rec!RUB_DESCRI
'            CmdBuscAprox.SetFocus
'        Else
'            MsgBox "El codigo no existe", vbInformation
'            txtRubro.SetFocus
'        End If
'        rec.Close
'     End If
End Sub
