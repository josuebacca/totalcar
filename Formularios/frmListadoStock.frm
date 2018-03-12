VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmListadoStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Stock - Productos Faltantes!!!"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
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
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   10740
      Begin VB.TextBox txtcodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8880
         TabIndex        =   14
         Top             =   255
         Width           =   1110
      End
      Begin VB.CheckBox chkRubro 
         Caption         =   "Rubro"
         Height          =   285
         Left            =   5055
         TabIndex        =   13
         Top             =   1080
         Width           =   825
      End
      Begin VB.CheckBox chkLinea 
         Caption         =   "Línea"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1455
         Width           =   780
      End
      Begin VB.CheckBox chkProducto 
         Caption         =   "Producto"
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   1110
         Width           =   990
      End
      Begin VB.CommandButton cmdfiltrar 
         Caption         =   "&Filtrar"
         Height          =   690
         Left            =   10125
         Picture         =   "frmListadoStock.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1065
         Width           =   570
      End
      Begin VB.ComboBox cboRep 
         Height          =   315
         Left            =   6165
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1425
         Width           =   3870
      End
      Begin VB.ComboBox cbolinea 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1425
         Width           =   3300
      End
      Begin VB.ComboBox cborubro 
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1065
         Width           =   3870
      End
      Begin VB.CheckBox chkRepres 
         Caption         =   "Marca"
         Height          =   255
         Left            =   5055
         TabIndex        =   6
         Top             =   1455
         Width           =   1020
      End
      Begin VB.TextBox txtproducto 
         Height          =   315
         Left            =   1245
         TabIndex        =   5
         Top             =   1065
         Width           =   3285
      End
      Begin TabDlg.SSTab tabLista 
         Height          =   855
         Left            =   2520
         TabIndex        =   17
         Top             =   120
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1508
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Maquinarias"
         TabPicture(0)   =   "frmListadoStock.frx":27A2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cboListaPrecio"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Repuestos"
         TabPicture(1)   =   "frmListadoStock.frx":27BE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cboLPrecioRep"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.ComboBox cboListaPrecio 
            Height          =   315
            Left            =   260
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   400
            Width           =   5385
         End
         Begin VB.ComboBox cboLPrecioRep 
            Height          =   315
            Left            =   -74740
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   400
            Width           =   5385
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lista de Precios"
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
         Index           =   0
         Left            =   960
         TabIndex        =   16
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   9120
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmListadoStock.frx":27DA
      Height          =   750
      Left            =   9990
      Picture         =   "frmListadoStock.frx":2AE4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   870
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      DisabledPicture =   "frmListadoStock.frx":2DEE
      Height          =   750
      Left            =   9120
      Picture         =   "frmListadoStock.frx":30F8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   870
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3975
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorSel    =   8388736
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.Label lblestado 
      AutoSize        =   -1  'True
      Caption         =   "estado"
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
      Left            =   240
      TabIndex        =   3
      Top             =   6000
      Width           =   585
   End
End
Attribute VB_Name = "frmListadoStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboLinea_LostFocus()
    If cbolinea.ListIndex <> -1 Then
        chkLinea.Value = 1
        cborubro.Clear
        cargocboRubro (cbolinea.ItemData(cbolinea.ListIndex))
    Else
        cborubro.Clear
        cargocboRubro (-1)
        chkRubro.Value = 0
    End If
End Sub

Private Sub cboListaPrecio_Click()
    If cboListaPrecio.ListIndex <> -1 Then
        txtcodigo.Text = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
    Else
        txtcodigo.Text = ""
    End If
End Sub

Private Sub cboLPrecioRep_Click()
    If cboLPrecioRep.ListIndex <> -1 Then
        txtcodigo.Text = cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex)
    Else
        txtcodigo.Text = ""
    End If
End Sub

Private Sub cboRep_LostFocus()
    If cboRep.ListIndex <> -1 Then
        chkRepres.Value = 1
    End If
End Sub

Private Sub cboRubro_LostFocus()
    If cborubro.ListIndex <> -1 Then
        chkRubro.Value = 1
    End If
End Sub

Private Sub chklinea_Click()
    If chkLinea.Value = 1 Then
        cbolinea.Enabled = True
        cbolinea.ListIndex = 0
        cbolinea.SetFocus
    Else
         cbolinea.Enabled = False
        cbolinea.ListIndex = -1
    End If
End Sub

Private Sub chkProducto_Click()
  If chkProducto.Value = 1 Then
        txtproducto.Enabled = True
        txtproducto.SetFocus
    Else
        txtproducto.Enabled = False
        txtproducto.Text = ""
    End If
End Sub

Private Sub chkRepres_Click()
    If (chkLinea.Value = 1) And (chkRubro.Value = 1) Then
        cargocboRepres cbolinea.ItemData(cbolinea.ListIndex), cborubro.ItemData(cborubro.ListIndex)
    Else
        If chkLinea.Value = 1 Then
            cargocboRepres cbolinea.ItemData(cbolinea.ListIndex), -1
        Else
            If chkRubro.Value = 1 Then
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
    If chkLinea.Value = 0 Then
        cargocboRubro (-1)
    Else
        cargocboRubro (cbolinea.ItemData(cbolinea.ListIndex))
    End If
    If chkRubro.Value = 1 Then
        cborubro.Enabled = True
        cborubro.ListIndex = 0
        cborubro.SetFocus
    Else
        cborubro.Enabled = False
        cborubro.ListIndex = -1
    End If
End Sub

Private Sub cmdfiltrar_Click()
    sql = " SELECT P.PTO_DESCRI,L.LNA_DESCRI, "
    sql = sql & " R.RUB_DESCRI,RE.TPRE_DESCRI,P.PTO_PRECIO,P.PTO_PRECIOC,P.PTO_CODIGO,P.PTO_PRECIVA, "
    sql = sql & " P.PTO_STKMIN,DS.DST_STKFIS "
    sql = sql & " FROM PRODUCTO P,LINEAS L,RUBROS R,TIPO_PRESENTACION RE,LISTA_PRECIO LP,STOCK ST,DETALLE_STOCK DS"
    sql = sql & " WHERE P.LNA_CODIGO = L.LNA_CODIGO  AND "
    sql = sql & " LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.PTO_STKMIN <> " & 0 & " "
    sql = sql & " AND P.LIS_CODIGO <> " & 0 & " "
    sql = sql & " AND P.RUB_CODIGO = R.RUB_CODIGO AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
    sql = sql & " AND ST.STK_CODIGO = DS.STK_CODIGO AND P.PTO_CODIGO = DS.PTO_CODIGO"
    sql = sql & " AND DS.DST_STKFIS < P.PTO_STKMIN"
    If txtcodigo.Text <> "" Then
        sql = sql & " AND LP.LIS_CODIGO = " & txtcodigo.Text & " "
    End If
    If chkProducto.Value = 1 Then
        'sql = sql & " AND (P.PTO_DESCRI LIKE '" & txtproducto.Text & "%' "
        'sql = sql & " OR P.PTO_CODIGO LIKE '" & txtproducto.Text & "%' )"
        txtproducto.Text = Replace(txtproducto, "'", "´")
        
        sql = sql & " AND (P.PTO_DESCRI LIKE '" & txtproducto.Text & "%' "
        sql = sql & " OR P.PTO_CODIGO LIKE '" & txtproducto.Text & "%' )"
    End If
    If chkLinea.Value = 1 Then
        sql = sql & " AND L.LNA_CODIGO = " & cbolinea.ItemData(cbolinea.ListIndex) & " "
    End If
    If chkRubro.Value = 1 Then
        sql = sql & " AND R.RUB_CODIGO = " & cborubro.ItemData(cborubro.ListIndex) & " "
    End If
    If chkRepres.Value = 1 Then
        sql = sql & " AND RE.TPRE_CODIGO = " & cboRep.ItemData(cboRep.ListIndex) & " "
    End If
    
    sql = sql & " ORDER BY P.PTO_DESCRI"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    GrdModulos.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!PTO_CODIGO & Chr(9) & rec!PTO_DESCRI & Chr(9) & _
                               rec!LNA_DESCRI & Chr(9) & rec!RUB_DESCRI & Chr(9) & _
                               rec!TPRE_DESCRI & Chr(9) & rec!PTO_STKMIN & Chr(9) & _
                               rec!DST_STKFIS & Chr(9) & rec!DST_STKFIS - rec!PTO_STKMIN
            rec.MoveNext
        Loop
    Else
        MsgBox "No se encontraron productos con faltantes", vbInformation, TIT_MSGBOX
    End If
    rec.Close
    lblEstado.Caption = "Se encontraron " & GrdModulos.Rows - 1 & " productos debajo del stock mininmo"
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    Centrar_pantalla Me
    preparogrilla
    lblEstado = ""
    'cmdfiltrar_Click
'    sql = "UPDATE DETALLE_STOCK SET "
'    sql = sql & " DST_STKFIS = 120"
'    sql = sql & " WHERE PTO_CODIGO < 4000"
'    DBConn.Execute sql
'
'    sql = "UPDATE PRODUCTO SET "
'    sql = sql & " PTO_STKMIN = 100"
'    'sql = sql & " WHERE PTO_CODIGO LIKE '" & "11" & "'%"
'    DBConn.Execute sql
    
    cargocboLinea
    cargocboRubro (-1)
    cargocboRepres -1, -1  ' Para Cargar Marcas sin Lineas y Rubros
    'cargocboLista
    CargoCboListaPrecio ' maquina
    CargoCboLPrecioRep ' repuesto
    'CargoCboListaAdicionales ' adicionales
    tabLista.Tab = 1
    
End Sub
Private Sub CargoCboListaPrecio() '' Lista de Precios de Repuestos
    cboListaPrecio.Clear
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 6"   '6: Maquinaria
    sql = sql & " ORDER BY LP.LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboListaPrecio.AddItem rec!LIS_DESCRI
            cboListaPrecio.ItemData(cboListaPrecio.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboListaPrecio.ListIndex = -1
    End If
    rec.Close
End Sub
Function cargocboLista()
    cbodescri.Clear
    sql = "SELECT DISTINCT LIS_CODIGO,LIS_DESCRI,LIS_FECHA FROM LISTA_PRECIO "
    sql = sql & "ORDER BY LIS_DESCRI"
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    'cbodescri.AddItem "<TODOS>"
    If Rec2.EOF = False Then
        Do While Rec2.EOF = False
            cbodescri.AddItem Rec2!LIS_DESCRI
            cbodescri.ItemData(cbodescri.NewIndex) = Rec2!LIS_CODIGO
            Rec2.MoveNext
        Loop
        
        cbodescri.ListIndex = -1
    End If
'    Rec2.MoveFirst
'    txtcodigo.Text = Rec2!LIS_CODIGO
'    Fecha1.Text = Rec2!LIS_FECHA
'    Rec2.MoveFirst
    txtcodigo.Text = ""
    Fecha1.Text = ""
    Rec2.Close
    
End Function
Function cargocboLinea()
    cbolinea.Clear
    sql = "SELECT * FROM LINEAS  ORDER BY LNA_DESCRI"
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
    If cod <> -1 Then
        sql = sql & " WHERE LNA_CODIGO= " & cod
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
    sql = "SELECT * FROM TIPO_PRESENTACION WHERE TPRE_CODIGO <> 0 "
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



Private Sub CargoCboLPrecioRep() '' Lista de Precios de Repuestos
    cboLPrecioRep.Clear
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 7"   '6: Repuestos
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboLPrecioRep.AddItem rec!LIS_DESCRI
            cboLPrecioRep.ItemData(cboLPrecioRep.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboLPrecioRep.ListIndex = -1
    End If
    rec.Close
End Sub

Private Function preparogrilla()
    GrdModulos.FormatString = "Código|Producto|Linea|Rubro|Marca|Stock Minimo|Stock Actual |Stock Dif"
    GrdModulos.ColWidth(0) = 900 ' codigo
    GrdModulos.ColWidth(1) = 2300 ' Producto
    GrdModulos.ColWidth(2) = 1400 ' linea
    GrdModulos.ColWidth(3) = 1400 ' Rubro
    GrdModulos.ColWidth(4) = 1400 ' Marca
    GrdModulos.ColWidth(5) = 1200 ' Stock minimo
    GrdModulos.ColWidth(6) = 1100 ' Stock actual
    GrdModulos.ColWidth(7) = 1100 ' Stock Diferencia entre el fisico y el minimo
    GrdModulos.Rows = 1
End Function

