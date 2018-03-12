VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   ControlBox      =   0   'False
   Icon            =   "frmBuscar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescriB 
      Height          =   315
      Left            =   765
      TabIndex        =   0
      Top             =   105
      Width           =   2280
   End
   Begin VB.CommandButton cmdBuscaAprox 
      Height          =   330
      Left            =   3120
      MaskColor       =   &H8000000F&
      Picture         =   "frmBuscar.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ejecutar B�squeda"
      Top             =   105
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3645
      Picture         =   "frmBuscar.frx":2AAC
      TabIndex        =   3
      ToolTipText     =   " Salir "
      Top             =   90
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSFlexGridLib.MSFlexGrid grdBuscar 
      Height          =   4260
      Left            =   0
      TabIndex        =   2
      Top             =   495
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   7514
      _Version        =   393216
      Rows            =   1
      Cols            =   0
      FixedCols       =   0
      BackColorSel    =   8388736
      ForeColorSel    =   16777215
      GridColor       =   8421504
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "<Esc> Salir"
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
      Left            =   4710
      TabIndex        =   6
      Top             =   150
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<F3> Buscar"
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
      Left            =   3210
      TabIndex        =   5
      Top             =   150
      Width           =   1290
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Buscar:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   540
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TipoBusqueda As Integer
Public TipoEnt As Integer
Public CodigoCli As String
Dim Importe As Double
Public CodListaPrecio As Integer

Public Sub ArmaSQL()
    Select Case TipoBusqueda
    
    Case 1 'CLIENTE
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC"
        sql = sql & " FROM CLIENTE "
        sql = sql & " WHERE "
        sql = sql & " CLI_RAZSOC LIKE '%" & Trim(TxtDescriB) & "%'"
        sql = sql & " AND CLI_ESTADO=1"
        sql = sql & " ORDER BY CLI_RAZSOC"
    
    Case 2 'PRODUCTOS precio Venta
        If CodListaPrecio = 0 Then
            sql = "SELECT TOP 1000 P.PTO_CODIGO,P.PTO_DESCRI, P.PTO_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI"
            sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, TIPO_PRESENTACION RE"
            sql = sql & " WHERE"
            sql = sql & " P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
            'If IsNumeric(TxtDescriB) Then
            '    sql = sql & " P.PTO_CODIGO LIKE '" & XS(TxtDescriB) & "'"
            'Else
            '    sql = sql & " P.PTO_DESCRI LIKE '" & Trim(TxtDescriB) & "%'"
            'End If
            sql = sql & " AND (P.PTO_DESCRI LIKE '" & Trim(TxtDescriB) & "%' "
            sql = sql & " OR P.PTO_CODIGO LIKE '" & Trim(TxtDescriB) & "%' )"
            
            
            'sql = sql & " AND P.PTO_ESTADO=1"
        Else
            sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI, D.LIS_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI"
            sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, DETALLE_LISTA_PRECIO D, TIPO_PRESENTACION RE"
            sql = sql & " WHERE"
            sql = sql & " D.LIS_CODIGO=" & CodListaPrecio
            sql = sql & " AND P.PTO_CODIGO=D.PTO_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
            'If IsNumeric(TxtDescriB) Then
            '    sql = sql & " P.PTO_CODIGO=" & XN(TxtDescriB)
            'Else
            '    sql = sql & " P.PTO_DESCRI LIKE '" & Trim(TxtDescriB) & "%'"
            'End If
            sql = sql & " AND (P.PTO_DESCRI LIKE '" & Trim(TxtDescriB) & "%' "
            sql = sql & " OR P.PTO_CODIGO LIKE '" & Trim(TxtDescriB) & "%' )"
            
            
            'sql = sql & " AND PTO_ESTADO=1"
        End If
            sql = sql & " ORDER BY PTO_DESCRI"
            
    Case 3 'SUCURSALES
        sql = "SELECT S.SUC_CODIGO,S.SUC_DESCRI,C.CLI_RAZSOC,S.CLI_CODIGO"
        sql = sql & " FROM SUCURSAL S, CLIENTE C"
        sql = sql & " WHERE S.CLI_CODIGO=C.CLI_CODIGO"
        If IsNumeric(TxtDescriB) Then
            sql = sql & " AND  S.SUC_CODIGO=" & XN(TxtDescriB)
        Else
            sql = sql & " AND S.SUC_DESCRI LIKE '" & Trim(TxtDescriB) & "%' "
        End If
        If CodigoCli <> "" Then
            sql = sql & " AND C.CLI_CODIGO=" & XN(CodigoCli)
        End If
        sql = sql & " AND C.CLI_ESTADO=1"
        sql = sql & " ORDER BY S.SUC_DESCRI"
        
    Case 4 'VENDEDORES
        sql = "SELECT VEN_CODIGO,VEN_NOMBRE,VEN_DOMICI"
        sql = sql & " FROM VENDEDOR"
        sql = sql & " WHERE"
        sql = sql & " VEN_NOMBRE LIKE '" & Trim(TxtDescriB) & "%' "
        sql = sql & " ORDER BY VEN_NOMBRE"
    
    Case 5 'PROVEEDORES
        sql = "SELECT TP.TPR_CODIGO,TP.TPR_DESCRI,P.PROV_CODIGO,P.PROV_RAZSOC"
        sql = sql & " FROM TIPO_PROVEEDOR TP, PROVEEDOR P"
        sql = sql & " WHERE"
        sql = sql & " TP.TPR_CODIGO=P.TPR_CODIGO"
        sql = sql & " AND PROV_RAZSOC LIKE '%" & Trim(TxtDescriB) & "%' "
        sql = sql & " ORDER BY PROV_RAZSOC"
    
    Case 6 'CHEQUES EN CARTERA
        sql = "SELECT DISTINCT TOP 200 CE.CHE_NUMERO, CH.CHE_IMPORT, CH.CHE_FECVTO, CE.BAN_CODINT, B.BAN_BANCO, B.BAN_LOCALIDAD,"
        sql = sql & " B.BAN_SUCURSAL, B.BAN_CODIGO, B.BAN_NOMCOR,CE.CES_DESCRI,B.BAN_DESCRI"
        sql = sql & " FROM CHEQUE_ESTADOS CE, CHEQUE CH, BANCO B,ESTADO_CHEQUE E"
        sql = sql & " Where "
        sql = sql & " CE.CHE_NUMERO = CH.CHE_NUMERO And "
        sql = sql & " CE.BAN_CODINT = CH.BAN_CODINT And "
        sql = sql & " CH.BAN_CODINT=B.BAN_CODINT AND "
        sql = sql & " CE.ECH_CODIGO= E.ECH_CODIGO AND" '
        sql = sql & " E.ECH_CODIGO=1"
        If TxtDescriB.Text <> "" Then
            sql = sql & " AND CH.CHE_NUMERO LIKE '" & Trim(TxtDescriB) & "'"  'CODIGO (1) ES CHEQUE EN CARTERA
        End If
    Case 7 'PRODUCTOS precio compra
        If CodListaPrecio = 0 Then
            sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI, P.PTO_PRECIOC, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI"
            sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, TIPO_PRESENTACION RE"
            sql = sql & " WHERE"
            If IsNumeric(TxtDescriB) Then
                sql = sql & " P.PTO_CODIGO=" & XN(TxtDescriB)
            Else
                sql = sql & " P.PTO_DESCRI LIKE '" & Trim(TxtDescriB) & "%'"
            End If
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
            sql = sql & " AND P.PTO_ESTADO=1"
        Else
            sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI, D.LIS_PRECIOC, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI"
            sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, DETALLE_LISTA_PRECIO D, TIPO_PRESENTACION RE"
            sql = sql & " WHERE"
            If IsNumeric(TxtDescriB) Then
                sql = sql & " P.PTO_CODIGO=" & XN(TxtDescriB)
            Else
                sql = sql & " P.PTO_DESCRI LIKE '" & Trim(TxtDescriB) & "%'"
            End If
            sql = sql & " AND D.LIS_CODIGO=" & CodListaPrecio
            sql = sql & " AND P.PTO_CODIGO=D.PTO_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
            sql = sql & " AND PTO_ESTADO=1"
        End If
            sql = sql & " ORDER BY PTO_DESCRI"
            
     Case 8 ' cheques entregados
        sql = "SELECT DISTINCT CE.CHE_NUMERO, CH.CHE_IMPORT, CH.CHE_FECVTO, CE.BAN_CODINT, B.BAN_BANCO, B.BAN_LOCALIDAD,"
        sql = sql & " B.BAN_SUCURSAL, B.BAN_CODIGO, B.BAN_NOMCOR,CE.CES_DESCRI,B.BAN_DESCRI"
        sql = sql & " FROM CHEQUE_ESTADOS CE, CHEQUE CH, BANCO B,ESTADO_CHEQUE E"
        sql = sql & " Where "
        sql = sql & " CE.CHE_NUMERO = CH.CHE_NUMERO And "
        sql = sql & " CE.BAN_CODINT = CH.BAN_CODINT And "
        sql = sql & " CH.BAN_CODINT=B.BAN_CODINT AND "
        sql = sql & " CE.ECH_CODIGO= E.ECH_CODIGO AND" '
        sql = sql & " E.ECH_CODIGO=7" ' 7-entregado
        If TxtDescriB.Text <> "" Then
            sql = sql & " AND CH.CHE_NUMERO LIKE '" & Trim(TxtDescriB) & "'"  'CODIGO (1) ES CHEQUE EN CARTERA
        End If
      Case 9 'cheques propios
        sql = "SELECT DISTINCT CE.CHEP_NUMERO, CH.CHEP_IMPORT, CH.CHEP_FECVTO, CE.BAN_CODINT, B.BAN_BANCO, B.BAN_LOCALIDAD,"
        sql = sql & " B.BAN_SUCURSAL, B.BAN_CODIGO, B.BAN_NOMCOR,CE.CPES_DESCRI,B.BAN_DESCRI"
        sql = sql & " FROM CHEQUE_PROPIO_ESTADO CE, CHEQUE_PROPIO CH, BANCO B,ESTADO_CHEQUE E"
        sql = sql & " Where "
        sql = sql & " CE.CHEP_NUMERO = CH.CHEP_NUMERO And "
        sql = sql & " CE.BAN_CODINT = CH.BAN_CODINT And "
        sql = sql & " CH.BAN_CODINT=B.BAN_CODINT AND "
        sql = sql & " CE.ECH_CODIGO= E.ECH_CODIGO AND" '
        sql = sql & " E.ECH_CODIGO<>5" ' 5-anulado
        If TxtDescriB.Text <> "" Then
            sql = sql & " AND CH.CHEP_NUMERO LIKE '" & Trim(TxtDescriB) & "'"  'CODIGO (1) ES CHEQUE EN CARTERA
        End If
     End Select
End Sub
Public Sub RellenaGrilla(Registro As ADODB.Recordset)
    Select Case TipoBusqueda
    
    Case 1 'CLIENTES
        Do While Not Registro.EOF
            grdBuscar.AddItem Trim(Registro!CLI_CODIGO) & Chr(9) & _
                Trim(Registro!CLI_RAZSOC)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
    
    Case 2 'PRODUCTOS
        Do While Not Registro.EOF
            If CodListaPrecio = 0 Then
                grdBuscar.AddItem Registro!PTO_CODIGO & Chr(9) & _
                    Trim(Registro!PTO_DESCRI) & Chr(9) & _
                    Valido_Importe(Registro!PTO_PRECIO) & Chr(9) & _
                    Trim(Registro!RUB_DESCRI) & Chr(9) & _
                    Trim(Registro!LNA_DESCRI) & Chr(9) & _
                    Trim(Registro!TPRE_DESCRI)
            Else 'SI USO LISTA DE PRECIO ENTRA ACA
                grdBuscar.AddItem Registro!PTO_CODIGO & Chr(9) & _
                    Trim(Registro!PTO_DESCRI) & Chr(9) & _
                    Valido_Importe(Registro!LIS_PRECIO) & Chr(9) & _
                    Trim(Registro!RUB_DESCRI) & Chr(9) & _
                    Trim(Registro!LNA_DESCRI) & Chr(9) & _
                    Trim(Registro!TPRE_DESCRI)
            End If
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
    
    Case 3 'SUCURSAL
        Do While Not Registro.EOF
            grdBuscar.AddItem Trim(Registro!SUC_CODIGO) & Chr(9) & _
                Trim(Registro!SUC_DESCRI) & Chr(9) & _
                Trim(Registro!CLI_RAZSOC) & Chr(9) & _
                Trim(Registro!CLI_CODIGO)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
        
    Case 4 'VENDEDORES
        Do While Not Registro.EOF
            grdBuscar.AddItem Trim(Registro!VEN_CODIGO) & Chr(9) & _
                Trim(Registro!VEN_NOMBRE) & Chr(9) & _
                Trim(Registro!VEN_DOMICI) & Chr(9) & _
                Trim(Registro!VEN_CODIGO)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
    
    Case 5 'PROVEEDORES
        Do While Not Registro.EOF
            grdBuscar.AddItem Trim(Registro!TPR_DESCRI) & Chr(9) & _
                Trim(Registro!PROV_CODIGO) & Chr(9) & _
                Trim(Registro!PROV_RAZSOC) & Chr(9) & _
                Trim(Registro!TPR_CODIGO)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
        'Text1.Text = grdBuscar.Rows - 1
    Case 6 'CHEQUES EN CARTERA
        Do While Not Registro.EOF
            grdBuscar.AddItem Trim(Registro!BAN_NOMCOR) & Chr(9) & _
                Trim(Registro!CHE_NUMERO) & Chr(9) & _
                Trim(Registro!CHE_FECVTO) & Chr(9) & _
                Trim(Valido_Importe(Registro!che_import)) & Chr(9) & _
                Trim(Registro!BAN_CODINT) & Chr(9) & _
                Trim(Registro!BAN_BANCO) & Chr(9) & _
                Trim(Registro!BAN_LOCALIDAD) & Chr(9) & _
                Trim(Registro!BAN_SUCURSAL) & Chr(9) & _
                Trim(Registro!BAN_CODIGO) & Chr(9) & _
                Trim(Registro!CES_DESCRI) & Chr(9) & _
                Trim(Registro!BAN_DESCRI)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
    Case 7 'PRODUCTOS PRECIO COMPRA
        Do While Not Registro.EOF
            If CodListaPrecio = 0 Then
                grdBuscar.AddItem Registro!PTO_CODIGO & Chr(9) & _
                    Trim(Registro!PTO_DESCRI) & Chr(9) & _
                    Valido_Importe(Registro!PTO_PRECIOC) & Chr(9) & _
                    Trim(Registro!RUB_DESCRI) & Chr(9) & _
                    Trim(Registro!LNA_DESCRI) & Chr(9) & _
                    Trim(Registro!TPRE_DESCRI)
            Else 'SI USO LISTA DE PRECIO ENTRA ACA
                grdBuscar.AddItem Registro!PTO_CODIGO & Chr(9) & _
                    Trim(Registro!PTO_DESCRI) & Chr(9) & _
                    Valido_Importe(Registro!LIS_PRECIOC) & Chr(9) & _
                    Trim(Registro!RUB_DESCRI) & Chr(9) & _
                    Trim(Registro!LNA_DESCRI) & Chr(9) & _
                    Trim(Registro!TPRE_DESCRI)
            End If
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
     Case 8 'cheque entregado
        Do While Not Registro.EOF
               grdBuscar.AddItem Trim(Registro!BAN_NOMCOR) & Chr(9) & _
                   Trim(Registro!CHE_NUMERO) & Chr(9) & _
                   Trim(Registro!CHE_FECVTO) & Chr(9) & _
                   Trim(Valido_Importe(Registro!che_import)) & Chr(9) & _
                   Trim(Registro!BAN_CODINT) & Chr(9) & _
                   Trim(Registro!BAN_BANCO) & Chr(9) & _
                   Trim(Registro!BAN_LOCALIDAD) & Chr(9) & _
                   Trim(Registro!BAN_SUCURSAL) & Chr(9) & _
                   Trim(Registro!BAN_CODIGO) & Chr(9) & _
                   Trim(Registro!CES_DESCRI) & Chr(9) & _
                   Trim(Registro!BAN_DESCRI)
               Registro.MoveNext
               grdBuscar.Refresh
           Loop
      Case 9 'cheque PROPIO entregado
        Do While Not Registro.EOF
               grdBuscar.AddItem Trim(Registro!BAN_NOMCOR) & Chr(9) & _
                   Trim(Registro!CHEP_NUMERO) & Chr(9) & _
                   Trim(Registro!CHEP_FECVTO) & Chr(9) & _
                   Trim(Valido_Importe(Registro!CHEP_IMPORT)) & Chr(9) & _
                   Trim(Registro!BAN_CODINT) & Chr(9) & _
                   Trim(Registro!BAN_BANCO) & Chr(9) & _
                   Trim(Registro!BAN_LOCALIDAD) & Chr(9) & _
                   Trim(Registro!BAN_SUCURSAL) & Chr(9) & _
                   Trim(Registro!BAN_CODIGO) & Chr(9) & _
                   Trim(Registro!CPES_DESCRI) & Chr(9) & _
                   Trim(Registro!BAN_DESCRI)
               Registro.MoveNext
               grdBuscar.Refresh
           Loop
    End Select
End Sub
Private Sub cmdBuscaAprox_Click()
'    If Trim(TxtDescriB) = "" Then
'        MsgBox "Debe especificar un detalle de B�squeda"
'        Exit Sub
'    End If
    Screen.MousePointer = vbHourglass
    Set Rec1 = New ADODB.Recordset
    ArmaSQL
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    EliminarFilasDeGrilla grdBuscar
    If Rec1.EOF = False Then
        RellenaGrilla Rec1
        grdBuscar.SetFocus
    Else
        MsgBox "No se han encontrado datos relacionados"
        SelecTexto TxtDescriB
        TxtDescriB.SetFocus
    End If
    Rec1.Close
    Set Rec1 = Nothing
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdSalir_Click()
    grdBuscar.Clear
    Me.Hide
End Sub

Private Sub Form_Activate()
    grdBuscar.Rows = 1
    Importe = 0
    
    Select Case TipoBusqueda
    
    Case 1 'CLIENTES
        Me.Caption = "Buscar:[Clientes]"
        grdBuscar.Width = 6200
        grdBuscar.Cols = 2
        grdBuscar.FormatString = "C�digo|Raz�n Social"
        grdBuscar.ColWidth(0) = 1000
        grdBuscar.ColWidth(1) = 5000
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
    
    Case 2 'PRODUCTOS
        Me.Caption = "Buscar:[Productos]"
        grdBuscar.Width = 11200
        grdBuscar.Cols = 6
        grdBuscar.FormatString = ">C�digo|Descripci�n|>Precio|Rubro|Linea|Representada"
        grdBuscar.ColWidth(0) = 1000
        grdBuscar.ColWidth(1) = 5000
        grdBuscar.ColWidth(2) = 1000
        grdBuscar.ColWidth(3) = 2000
        grdBuscar.ColWidth(4) = 2000
        grdBuscar.ColWidth(5) = 2000
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
    
    Case 3 'SUCURSALES
        Me.Caption = "Buscar:[Sucursales]"
        grdBuscar.Width = 11200
        grdBuscar.Cols = 4
        grdBuscar.FormatString = "C�digo|Raz�n Social|Cliente|CODCLI"
        grdBuscar.ColWidth(0) = 1000
        grdBuscar.ColWidth(1) = 5000
        grdBuscar.ColWidth(2) = 5000
        grdBuscar.ColWidth(3) = 0
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
    
    Case 4 'VENDEDORES
        Me.Caption = "Buscar:[Vendedores]"
        grdBuscar.Width = 7700
        grdBuscar.Cols = 4
        grdBuscar.FormatString = "N�mero|Nombre|Domicilio|NUMVENDEDOR"
        grdBuscar.ColWidth(0) = 1000
        grdBuscar.ColWidth(1) = 3500
        grdBuscar.ColWidth(2) = 3000
        grdBuscar.ColWidth(3) = 0
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
        
    Case 5 'PROVEDORES
        Me.Caption = "Buscar:[Proveedores]"
        grdBuscar.Width = 8000
        grdBuscar.Cols = 4
        grdBuscar.FormatString = "Tipo Proveedor|N�mero|Raz�n Social|Cod Tipo Prov"
        grdBuscar.ColWidth(0) = 3000
        grdBuscar.ColWidth(1) = 800
        grdBuscar.ColWidth(2) = 4000
        grdBuscar.ColWidth(3) = 0
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
    
    Case 6 'CHEQUES EN CARTERA
        Me.Caption = "Buscar:[Cheques en Cartera (de Terceros)]"
        grdBuscar.Width = 9000
        grdBuscar.Cols = 11
        grdBuscar.FormatString = "Banco|^Cheuqe Nro|^Fecha Vto|>Importe|BAN_CODINT|BAN_BANCO" _
                                & "|BAN_LOCALIDAD|BAN_SUCURSAL|BAN_CODIGO|Estado|BANDESCRI"
        grdBuscar.ColWidth(0) = 3500 'Banco
        grdBuscar.ColWidth(1) = 1200 'Cheuqe Nro
        grdBuscar.ColWidth(2) = 1100 'Fecha Vto
        grdBuscar.ColWidth(3) = 1100 'Importe
        grdBuscar.ColWidth(4) = 0    'BAN_CODINT
        grdBuscar.ColWidth(5) = 0    'BAN_BANCO
        grdBuscar.ColWidth(6) = 0    'BAN_LOCALIDAD
        grdBuscar.ColWidth(7) = 0    'BAN_SUCURSAL
        grdBuscar.ColWidth(8) = 0    'BAN_CODIGO
        grdBuscar.ColWidth(9) = 2000 'CES_DESCRI
        grdBuscar.ColWidth(10) = 0 'BAN_DESCRI
        cmdBuscaAprox_Click
    Case 7 'PRODUCTOS PRECIO COMPRA
        Me.Caption = "Buscar:[Productos]"
        grdBuscar.Width = 11200
        grdBuscar.Cols = 6
        grdBuscar.FormatString = ">C�digo|Descripci�n|>Precio|Rubro|Linea|Representada"
        grdBuscar.ColWidth(0) = 1000
        grdBuscar.ColWidth(1) = 5000
        grdBuscar.ColWidth(2) = 1000
        grdBuscar.ColWidth(3) = 2000
        grdBuscar.ColWidth(4) = 2000
        grdBuscar.ColWidth(5) = 2000
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
    Case 8 'CHEQUES entregados
        Me.Caption = "Buscar:[Cheques en Cartera (de Terceros)]"
        grdBuscar.Width = 9000
        grdBuscar.Cols = 11
        grdBuscar.FormatString = "Banco|^Cheuqe Nro|^Fecha Vto|>Importe|BAN_CODINT|BAN_BANCO" _
                                & "|BAN_LOCALIDAD|BAN_SUCURSAL|BAN_CODIGO|Estado|BANDESCRI"
        grdBuscar.ColWidth(0) = 3500 'Banco
        grdBuscar.ColWidth(1) = 1200 'Cheuqe Nro
        grdBuscar.ColWidth(2) = 1100 'Fecha Vto
        grdBuscar.ColWidth(3) = 1100 'Importe
        grdBuscar.ColWidth(4) = 0    'BAN_CODINT
        grdBuscar.ColWidth(5) = 0    'BAN_BANCO
        grdBuscar.ColWidth(6) = 0    'BAN_LOCALIDAD
        grdBuscar.ColWidth(7) = 0    'BAN_SUCURSAL
        grdBuscar.ColWidth(8) = 0    'BAN_CODIGO
        grdBuscar.ColWidth(9) = 2000 'CES_DESCRI
        grdBuscar.ColWidth(10) = 0 'BAN_DESCRI
        cmdBuscaAprox_Click
     Case 9 'CHEQUES PROPIOS ENTREGADOS
        Me.Caption = "Buscar:[Cheques Librados (Propios)]"
        grdBuscar.Width = 9000
        grdBuscar.Cols = 11
        grdBuscar.FormatString = "Banco|^Cheuqe Nro|^Fecha Vto|>Importe|BAN_CODINT|BAN_BANCO" _
                                & "|BAN_LOCALIDAD|BAN_SUCURSAL|BAN_CODIGO|Estado|BANDESCRI"
        grdBuscar.ColWidth(0) = 3500 'Banco
        grdBuscar.ColWidth(1) = 1200 'Cheuqe Nro
        grdBuscar.ColWidth(2) = 1100 'Fecha Vto
        grdBuscar.ColWidth(3) = 1100 'Importe
        grdBuscar.ColWidth(4) = 0    'BAN_CODINT
        grdBuscar.ColWidth(5) = 0    'BAN_BANCO
        grdBuscar.ColWidth(6) = 0    'BAN_LOCALIDAD
        grdBuscar.ColWidth(7) = 0    'BAN_SUCURSAL
        grdBuscar.ColWidth(8) = 0    'BAN_CODIGO
        grdBuscar.ColWidth(9) = 2000 'CES_DESCRI
        grdBuscar.ColWidth(10) = 0 'BAN_DESCRI
        cmdBuscaAprox_Click
    End Select
    Me.Width = grdBuscar.Width + 70
    Call Centrar_pantalla(Me)
    If grdBuscar.Rows > 1 Then
        grdBuscar.SetFocus
    Else
        TxtDescriB.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub grdBuscar_Click()
'    If grdBuscar.Rows <> 1 Then
'        OrdenarGrilla grdBuscar
'    End If
End Sub

Private Sub GrdBuscar_dblClick()
    If grdBuscar.Rows > 1 Then
        Select Case TipoBusqueda
        Case 1
            grdBuscar.Col = 0
        Case 2
            grdBuscar.Col = 0
        Case 3
            grdBuscar.Col = 0
        Case 4
            grdBuscar.Col = 0
'        Case 5
'            grdBuscar.Col = 0
'        Case 6
'            grdBuscar.Col = 0
'        Case 7
'            grdBuscar.Col = 0
'        Case 8
'            grdBuscar.Col = 0
'        Case 99
'            grdBuscar.Col = 0
        End Select
        Me.Hide
    End If
End Sub

Private Sub grdBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GrdBuscar_dblClick
    End If
End Sub

Private Sub TxtDescriB_GotFocus()
    SelecTexto TxtDescriB
End Sub

Private Sub txtDescriB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
        cmdBuscaAprox_Click
    End If
End Sub


Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
