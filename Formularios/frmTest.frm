VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Lista de Precios"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   855
      Left            =   3150
      Picture         =   "frmTest.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2450
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   5160
      Picture         =   "frmTest.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2450
      Width           =   975
   End
   Begin VB.CommandButton cmdReadXLS 
      Caption         =   "&Importar"
      Height          =   855
      Left            =   4155
      Picture         =   "frmTest.frx":1484
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2450
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6015
      Begin VB.Timer tmrCuenta 
         Left            =   5280
         Top             =   1920
      End
      Begin VB.PictureBox Progreso 
         Height          =   255
         Left            =   240
         ScaleHeight     =   195
         ScaleWidth      =   5475
         TabIndex        =   10
         Top             =   1560
         Width           =   5535
      End
      Begin VB.CommandButton cmdSelArch 
         Caption         =   "..."
         Height          =   325
         Left            =   5400
         TabIndex        =   1
         Top             =   900
         Width           =   425
      End
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   900
         Width           =   3615
      End
      Begin VB.ComboBox cboLPrecioRep 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   4185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2880
         TabIndex        =   13
         Top             =   1980
         Width           =   45
      End
      Begin VB.Label lblCuenta 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Restante:"
         Height          =   195
         Left            =   810
         TabIndex        =   11
         Top             =   1980
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Seleccionar Archivo:"
         Height          =   195
         Left            =   75
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lista de Precios:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdReadTXT 
      Caption         =   "&Read txt"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblestado 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   2640
      Width           =   465
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const m_intInicio As Integer = 100
Dim I As Integer, X As Integer

' ------------------------------------------------------------
'  Copyright ©2001 Mike G --> IvbNET.COM
'  All Rights Reserved, http://www.ivbnet.com
'  EMAIL : webmaster@ivbnet.com
' ------------------------------------------------------------
'  You are free to use this code within your own applications,
'  but you are forbidden from selling or distributing this
'  source code without prior written consent.
' ------------------------------------------------------------

Private Sub cmdExit_Click()
  Unload Me
End Sub

'Private Sub cmdReadTXT_Click()
'      Dim obj As Read_Files.CReadFile
'      Set obj = New Read_Files.CReadFile
'
'      Set dgData.DataSource = obj.Read_Text_File
'      Set obj = Nothing
'End Sub
Private Function ValidarImportacion() As Boolean
    If cboLPrecioRep.ListIndex = -1 Then
        MsgBox "No ha seleccionado una Lista de Precios", vbExclamation, TIT_MSGBOX
        cboLPrecioRep.SetFocus
        ValidarImportacion = False
        Exit Function
    End If
    If txtArchivo.Text = "" Then
        MsgBox "No ha seleccionado el Archivo a Importar", vbExclamation, TIT_MSGBOX
        cmdSelArch.SetFocus
        ValidarImportacion = False
        Exit Function
    End If
    ValidarImportacion = True
End Function

Private Sub CmdNuevo_Click()
    cboLPrecioRep.ListIndex = -1
    txtArchivo.Text = ""
    'Progreso.Value = 0
    Label4.Caption = ""
    lblEstado.Caption = ""
    
End Sub

Private Sub cmdReadXLS_Click()
'      Dim obj As Read_Files.CReadFile
'      Set obj = New Read_Files.CReadFile
      Dim I, j As Integer
      Dim sql As String
      Dim rec As New ADODB.Recordset
      Dim Rec1 As New ADODB.Recordset
      Dim nLinea As Integer
      Dim nRubro As Integer
      Dim nMarca As Integer
      Dim nDescri As String
      'Dim nPrecioVenta As String
      'Dim nPrecioIVA As String
      Dim resp As Integer
      Dim Codigo As String
      'Dim rec As New ADODB.Recordset
      
      Dim e As New Excel.Application ' declaración
      
    'de la variable que se enlazara con excel
    'App.Path & "\" &
    e.Workbooks.Open txtArchivo.Text 'habre
    'un libro llamado indicadores.xls
    e.Worksheets(1).Activate 'coloca la primera hoja del libro
    ' como la hoja activa
    'luego puede acceder a las celdas con
    ' e.Cells(fila,columna) ve el ejemplo de abajo
    ' que llena el equivalente
    ' al rango Range(A1:C5) con ceros
    'n = 1
    'For I = 1 To 5
    '    For J = 1 To 7
    '        MsgBox e.cells(7, J)
    '    Next J
        'n = n + 1
    '    'If IsNull(e.cells(I, 1)) Then
                
        'End If
    'Next I
    
    

      
      
      If ValidarImportacion = False Then Exit Sub
      resp = MsgBox("Confirma la Importacion de la Nueva Lista de Precios de : " & Trim(cboLPrecioRep.Text) & "?.   Este proceso puede demorar unos minutos ", 36, TIT_MSGBOX)
      If resp <> 6 Then Exit Sub
      
      'Set rec = obj.Read_Excel(txtArchivo.Text)
      'Set obj = Nothing
      
        'BARRA DE ESTADO
        'Establecemos el primer valor de la etiqueta.
        Me.lblCuenta.Caption = m_intInicio
        'Establece el intervalo al Timer para que
        'comience a descontar .
        'Me.tmrCuenta.Interval = 100    'milisegundos
        
       'ESTO LO HAGO PARA GUARDAR LA LINEA RUBRO Y MARCA DE LA LISTA DE PRECIO
        sql = "select * FROM PRODUCTO WHERE LIS_CODIGO = " & cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            nLinea = Rec1!LNA_CODIGO
            nRubro = Rec1!RUB_CODIGO
            nMarca = Rec1!TPRE_CODIGO
        End If
        Rec1.Close
        
        'BORRO LA LISTA A IMPORTAR
        'sql = "Delete from PRODUCTO WHERE (LIS_CODIGO = " & cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex) & ")"
        'sql = sql & " OR (LNA_CODIGO = " & nLinea & ""
        'sql = sql & " AND RUB_CODIGO = " & nRubro & ""
        'sql = sql & " AND TPRE_CODIGO = " & nMarca & ")"
        'DBConn.Execute sql
        I = 2
        Do While Trim(e.Cells(I, 1)) <> ""
        'If rec.EOF = False Then
            'If I = 890 Then
            '    MsgBox I
            'End If
            lblEstado.Caption = "Procesando...."
            Me.tmrCuenta.Interval = 100
        '    Do While rec.EOF = False
                
                'Calulos
                'nPrecioVenta = 0
                'nPrecioIVA = 0
                'nPrecioVenta = rec.Fields(2) + (rec.Fields(2) * 20) / 100
                'nPrecioIVA = Valido_Importe(nPrecioVenta) * 1.21
                'If IsNull(rec.Fields(0)) Then
                    'If IsNull(rec.Fields(2)) Then
                      ' rec.Close
                        'lblestado.Caption = "Fin de proceso"
                '        Exit Sub
                    'Else
                        'Codigo = rec.Fields(2)
                    'End If
                'Else
                '    Codigo = rec.Fields(0)
                'End If
                'pregunto si existe el producto
                'si existe actualizo los precios
                sql = "SELECT * FROM PRODUCTO WHERE "
                sql = sql & " PTO_CODIGO = '" & Trim(e.Cells(I, 1)) & "' " 'CODIGO
                'sql = sql & " PTO_CODIGO LIKE '" & rec.Fields(0) & "' "
                'sql = sql & " AND LNA_CODIGO = " & nLinea
                'sql = sql & " AND RUB_CODIGO = " & nRubro
                'sql = sql & " AND TPRE_CODIGO = " & nMarca
                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Rec1.EOF = False Then
                    'si existe ACTUALIZO LOS PRECIOS
                    sql = "UPDATE PRODUCTO SET"
'                    nDescri = ""
'                    nDescri = Replace(e.Cells(I, 2), "'", "")
'                    sql = sql & " PTO_DESCRI = " & XS(nDescri)
                    sql = sql & " PTO_PRECIO = " & XN(Valido_Importe(e.Cells(I, 7))) 'Precio
                    sql = sql & ",PTO_PRECIOC = " & XN(Valido_Importe(e.Cells(I, 8))) 'Precio Compra
                    'sql = sql & ",PTO_PRECIVA = " & XN(Valido_Importe(e.Cells(I, 6))) 'Precio Con IVA
                    sql = sql & ",LIS_CODIGO = " & XN(cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex))
                    'sql = sql & " WHERE PTO_CODIGO LIKE '" & rec.Fields(0) & "' "
                    sql = sql & " WHERE PTO_CODIGO = '" & Trim(e.Cells(I, 1)) & "' "
                    DBConn.Execute sql
                Else
                    'si no existe lo inserto
                    'Obtener el codigo de la marca
                    
                    sql = "INSERT INTO PRODUCTO (PTO_CODIGO,LNA_CODIGO,RUB_CODIGO,"
                    sql = sql & "TPRE_CODIGO,PTO_DESCRI,PTO_PRECIO,PTO_PRECIOC,PTO_PRECIVA,LIS_CODIGO)"
                    sql = sql & " VALUES ("
                    'txtDescri.Text = Replace(txtDescri, "'", "´")
                    'sql = sql & rec.Fields(0) & ","
                    sql = sql & XS(Trim(e.Cells(I, 1))) & ","
                    sql = sql & nLinea & ","
                    sql = sql & Trim(e.Cells(I, 3)) & ","
                    sql = sql & Trim(e.Cells(I, 4)) & ","
                    nDescri = ""
                    nDescri = Replace(Trim(e.Cells(I, 5)), "'", "")
                    sql = sql & XS(nDescri) & ","
                    sql = sql & XN(Valido_Importe(e.Cells(I, 7))) & ","
                    sql = sql & XN(Valido_Importe(e.Cells(I, 8))) & ","
                    sql = sql & XN("0") & ","
                    sql = sql & XN(cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex)) & ")"
                    DBConn.Execute sql
                End If
                Rec1.Close
                I = I + 1
                'rec.MoveNext
                
                
            Loop
         'Else
         '   MsgBox "Se ha producido un error en la importacion. Consulte con el programador!", vbExclamation, TIT_MSGBOX
         'End If
         'rec.Close
         lblEstado.Caption = "Fin de proceso"
        
         e.Visible = False ' hace visible la hoja de aclculo
         Set e = Nothing ' libera los recursos de memoria
      
      
      
      
      
      
      
      
      
      
      'Set rec = obj.Read_Excel(App.Path & "\" & "test.xls")
'      Set rec = obj.Read_Excel(txtArchivo.Text)
'      Set obj = Nothing
'
'        'BARRA DE ESTADO
'        'Establecemos el primer valor de la etiqueta.
'        Me.lblCuenta.Caption = m_intInicio
'        'Establece el intervalo al Timer para que
'        'comience a descontar .
'        'Me.tmrCuenta.Interval = 100    'milisegundos
'
'       'ESTO LO HAGO PARA GUARDAR LA LINEA RUBRO Y MARCA DE LA LISTA DE PRECIO
'        sql = "select * FROM PRODUCTO WHERE LIS_CODIGO = " & cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex)
'        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If Rec1.EOF = False Then
'            nLinea = Rec1!LNA_CODIGO
'            nRubro = Rec1!RUB_CODIGO
'            nMarca = Rec1!TPRE_CODIGO
'        End If
'        Rec1.Close
'
'        'BORRO LA LISTA A IMPORTAR
'        'sql = "Delete from PRODUCTO WHERE (LIS_CODIGO = " & cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex) & ")"
'        'sql = sql & " OR (LNA_CODIGO = " & nLinea & ""
'        'sql = sql & " AND RUB_CODIGO = " & nRubro & ""
'        'sql = sql & " AND TPRE_CODIGO = " & nMarca & ")"
'        'DBConn.Execute sql
'        I = 2
'        If rec.EOF = False Then
'            lblestado.Caption = "Procesando...."
'            Me.tmrCuenta.Interval = 100
'            Do While rec.EOF = False
'
'                'Calulos
'                'nPrecioVenta = 0
'                'nPrecioIVA = 0
'                'nPrecioVenta = rec.Fields(2) + (rec.Fields(2) * 20) / 100
'                'nPrecioIVA = Valido_Importe(nPrecioVenta) * 1.21
'                If IsNull(rec.Fields(0)) Then
'                    'If IsNull(rec.Fields(2)) Then
'                      ' rec.Close
'                        lblestado.Caption = "Fin de proceso"
'                        Exit Sub
'                    'Else
'                        'Codigo = rec.Fields(2)
'                    'End If
'                Else
'                    Codigo = rec.Fields(0)
'                End If
'                'pregunto si existe el producto
'                'si existe actualizo los precios
'                sql = "SELECT * FROM PRODUCTO WHERE "
'                sql = sql & " PTO_CODIGO LIKE '" & Codigo & "' "
'                'sql = sql & " PTO_CODIGO LIKE '" & rec.Fields(0) & "' "
'                'sql = sql & " AND LNA_CODIGO = " & nLinea
'                'sql = sql & " AND RUB_CODIGO = " & nRubro
'                'sql = sql & " AND TPRE_CODIGO = " & nMarca
'                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'                If Rec1.EOF = False Then
'                    'si existe ACTUALIZO LOS PRECIOS
'                    sql = "UPDATE PRODUCTO SET"
'                    sql = sql & " PTO_PRECIO = " & XN(Valido_Importe(rec.Fields(6))) 'Precio
'                    sql = sql & ",PTO_PRECIOC = " & XN(Valido_Importe(rec.Fields(7))) 'Precio Compra
'                    sql = sql & ",PTO_PRECIVA = " & XN(Valido_Importe(rec.Fields(5))) 'Precio Con IVA
'                    sql = sql & ",LIS_CODIGO = " & XN(cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex))
'                    'sql = sql & " WHERE PTO_CODIGO LIKE '" & rec.Fields(0) & "' "
'                    sql = sql & " WHERE PTO_CODIGO LIKE '" & Codigo & "' "
'                    DBConn.Execute sql
'                Else
'                    'si no existe lo inserto
'                    sql = "INSERT INTO PRODUCTO (PTO_CODIGO,LNA_CODIGO,RUB_CODIGO,"
'                    sql = sql & "TPRE_CODIGO,PTO_DESCRI,PTO_PRECIO,PTO_PRECIOC,PTO_PRECIVA,LIS_CODIGO)"
'                    sql = sql & " VALUES ("
'                    'txtDescri.Text = Replace(txtDescri, "'", "´")
'                    'sql = sql & rec.Fields(0) & ","
'                    sql = sql & XS(Codigo) & ","
'                    sql = sql & nLinea & ","
'                    sql = sql & nRubro & ","
'                    sql = sql & nMarca & ","
'                    nDescri = ""
'                    nDescri = Replace(rec.Fields(1), "'", "")
'                    sql = sql & XS(nDescri) & ","
'                    sql = sql & XN(Valido_Importe(rec.Fields(6))) & ","
'                    sql = sql & XN(Valido_Importe(rec.Fields(7))) & ","
'                    sql = sql & XN(Valido_Importe(rec.Fields(5))) & ","
'                    sql = sql & XN(cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex)) & ")"
'                    DBConn.Execute sql
'                End If
'                Rec1.Close
'                I = I + 1
'                rec.MoveNext
'
'
'            Loop
'         Else
'            MsgBox "Se ha producido un error en la importacion. Consulte con el programador!", vbExclamation, TIT_MSGBOX
'         End If
'         rec.Close
'         lblestado.Caption = "Fin de proceso"

End Sub



Private Sub tmrCuenta_Timer()

    
    'Cada décima de segundo, descontamos.
    'Progreso.Value = Progreso.Value + 1
'    Me.lblCuenta.Caption = 100 - (Progreso.Value \ 1)
'    Label4 = Progreso & " % Finalizado"
'    'Cuando llega a cero, paramos.
'    If CInt(Me.lblCuenta.Caption) = 0 Then
'        Me.tmrCuenta.Interval = 0
'    End If
    
End Sub



Private Sub CargoCboLPrecioRep()
    cboLPrecioRep.Clear
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    'sql = sql & " AND P.LNA_CODIGO = 7"   '6: Repuestos
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

Private Sub cmdSelArch_Click()
On Error Resume Next
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Seleccione un nombre de archivo"
CommonDialog1.Filter = "*.xls"

CommonDialog1.ShowOpen
If Err.Number = 0 Then
    If (CommonDialog1.FileName Like "*.xlsx") Or (CommonDialog1.FileName Like "*.XLSX") Or (CommonDialog1.FileName Like "*.xls") Or (CommonDialog1.FileName Like "*.XLS") Or (CommonDialog1.FileName Like "*.ods") Or (CommonDialog1.FileName Like "*.ODS") Then
         'Image1.Picture = LoadPicture(CommonDialog1.FileName)
        txtArchivo.Text = CommonDialog1.FileName
        On Error GoTo 0
    Else
        MsgBox "El Archivo seleccionado no es válido", vbExclamation, Me.Caption
    End If
End If
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)
    CargoCboLPrecioRep
    lblEstado.Caption = ""
    
    
    
End Sub
