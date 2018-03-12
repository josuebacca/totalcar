VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestBD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar BD"
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
      Picture         =   "frmTestBD.frx":0000
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
      Picture         =   "frmTestBD.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2450
      Width           =   975
   End
   Begin VB.CommandButton cmdReadXLS 
      Caption         =   "&Importar"
      Height          =   855
      Left            =   4155
      Picture         =   "frmTestBD.frx":1484
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
      Begin MSComctlLib.ProgressBar Progreso 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
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
Attribute VB_Name = "frmTestBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const m_intInicio As Integer = 100
Dim i As Integer, X As Integer

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
    Progreso.Value = 0
    Label4.Caption = ""
    lblestado.Caption = ""
    
End Sub

Private Sub cmdReadXLS_Click()
      Dim obj As Read_Files.CReadFile
      Set obj = New Read_Files.CReadFile
      Dim i, J As Integer
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
      Dim CELDA As String
      Dim Array_A(300) As String
          'Dim rec As New ADODB.Recordset
      
      Dim e As New Excel.Application ' declaración
          
    'de la variable que se enlazara con excel
    'App.Path & "\" &
    e.Workbooks.Open txtArchivo.Text 'habre
    'un libro llamado indicadores.xls
    e.Worksheets(1).Activate 'coloca la primera hoja del libro
    ' como la hoja activa
    'luego puede acceder a las celdas con
    CELDA = e.Cells(1, 1)
    
    Dim a() As String, s As String
    'Dim i As Integer
    s = CELDA
    a = Split(s, ",")
    For i = 0 To UBound(a)
        sql = "INSERT INTO CLICEN (CEN_DESCRI)"
        sql = sql & " VALUES ("
        sql = sql & XS(a(i)) & ")"
        DBConn.Execute sql
    'Me.Print a(i)
    Next i
    
'
'
'    i = 1
'    For J = 1 To Len(CELDA)
'        Array_A(i) = Mid(CELDA, J, J + 109)
'        J = J + 109
'        i = i + 1
'    Next J
'
'    For J = 1 To i
'        sql = "INSERT INTO CLICEN (CEN_DESCRI)"
'        sql = sql & " VALUES ("
'        sql = sql & XS(Array_A(J)) & ")"
'        DBConn.Execute sql
'    Next J

   lblestado.Caption = "Fin de proceso"
  
   e.Visible = False ' hace visible la hoja de aclculo
   Set e = Nothing ' libera los recursos de memoria

End Sub
Private Sub tmrCuenta_Timer()

    
    'Cada décima de segundo, descontamos.
    Progreso.Value = Progreso.Value + 1
    Me.lblCuenta.Caption = 100 - (Progreso.Value \ 1)
    Label4 = Progreso & " % Finalizado"
    'Cuando llega a cero, paramos.
    If CInt(Me.lblCuenta.Caption) = 0 Then
        Me.tmrCuenta.Interval = 0
    End If
    
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
    If (CommonDialog1.FileName Like "*.xls") Or (CommonDialog1.FileName Like "*.XLS") Then
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
    lblestado.Caption = ""
    
    
    
End Sub
Private Sub CortarCadena(COLUMNA As Double, Renglon As Double, Cadena As String)
    Dim salto, Max, inf, i, leer, leerb As Integer
    Dim salto1, salto2, salto3, salto4, salto5, salto6, salto7 As Integer
    Dim salto1b, salto2b, salto3b, salto4b, salto5b, salto6b, salto7b As Integer
    Dim cadena1 As String
    Dim cadena2 As String
    Dim cadena3 As String
    Dim cadena4 As String
    Dim cadena5 As String
    Dim cadena6 As String
    Dim cadena7 As String
    Dim cadena8 As String
    
    
    
    cadena1 = ""
    cadena2 = ""
    cadena3 = ""
    cadena4 = ""
    cadena5 = ""
    cadena6 = ""
    cadena7 = ""
    
    
    salto = 1
    Max = 130 * salto
    inf = Max - 5
    'falta = 0
    'If Len(cadena) > 35 Then
        
        
        
        For i = 1 To Len(Cadena)
            For J = 1 To 100
                Array_A(J) = Split(Cadena, 150)
            
            Next J
        
        
            If (Mid(Cadena, i, 1) = " ") And (i > inf) And (i < Max) Or (i > Max) Then
                
                    If salto = 1 Then
                    salto1 = i
                    Max = 36 + i
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        cadena1 = Left(Cadena, i)
                        cadena2 = Mid(Cadena, salto1, Max)
                        'Imprimir 3.2, renglon, False, Left(grdGrilla.TextMatrix(I, 1), 36) 'descripcion
                        'Imprimir 3.2, renglon + 0.5, False, Mid(grdGrilla.TextMatrix(I, 1), 37, 36) 'descripcion
                    
                    Else
                        cadena1 = Left(Cadena, i)
                    End If
                      'descripcion
                End If
                If salto = 2 Then
                    leer = i - salto1
                    salto2 = i
                    Max = 36 + i
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto1b = i
                        leerb = Len(Cadena) + 1 - salto1b
                        cadena2 = Mid(Cadena, salto1, leer)
                        cadena3 = Mid(Cadena, salto1b, leerb)  'descripcion
                    Else
                        cadena2 = Mid(Cadena, salto1, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 3 Then
                    Max = 36 + i
                    inf = Max - 10
                    leer = i - salto2
                    salto3 = i
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto2b = i
                        leerb = Len(Cadena) + 1 - salto2b
                        cadena3 = Mid(Cadena, salto2, leer)
                        cadena4 = Mid(Cadena, salto2b, leerb)
                    Else
                        cadena3 = Mid(Cadena, salto2, leer)  'descripcion
                    End If
                    
                End If
                If salto = 4 Then
                    leer = i - salto3
                    salto4 = i
                    Max = 36 + i
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto3b = i
                        leerb = Len(Cadena) + 1 - salto3b
                        cadena4 = Mid(Cadena, salto3, leer)
                        cadena5 = Mid(Cadena, salto3b, leerb)  'descripcion
                    Else
                         cadena4 = Mid(Cadena, salto3, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 5 Then
                    leer = i - salto4
                    salto5 = i
                    Max = 36 + i
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto4b = i
                        leerb = Len(Cadena) + 1 - salto4b
                        cadena5 = Mid(Cadena, salto4, leer)
                        cadena6 = Mid(Cadena, salto4b, leerb)  'descripcion
                    Else
                        cadena5 = Mid(Cadena, salto4, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 6 Then
                    leer = i - salto5
                    salto6 = i
                    Max = 36 + i
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto5b = i
                        leerb = Len(Cadena) + 1 - salto5b
                        cadena6 = Mid(Cadena, salto5, leer)
                        cadena7 = Mid(Cadena, salto5b, leerb)  'descripcion
                        
                    Else
                        cadena6 = Mid(Cadena, salto5, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 7 Then
                    leer = i - salto6
                    salto7 = i
                    Max = 36 + i
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto6b = i
                        leerb = Len(Cadena) + 1 - salto6b
                        cadena7 = Mid(Cadena, salto6, leer)
                        cadena8 = Mid(Cadena, salto6b, leerb)  'descripcion
                        
                    Else
                        cadena7 = Mid(Cadena, salto6, leer)  'descripcion
                    End If
                    
                End If
                
                salto = salto + 1
                'Max = valor * salto
                'inf = Max - 10
                
            End If
        Next
    
'        Imprimir COLUMNA, Renglon, False, cadena1
'        Imprimir COLUMNA, Renglon + 0.5, False, Trim(cadena2)
'        Imprimir COLUMNA, Renglon + 1, False, Trim(cadena3)
'        Imprimir COLUMNA, Renglon + 1.5, False, Trim(cadena4)
'        Imprimir COLUMNA, Renglon + 2, False, Trim(cadena5)
'        Imprimir COLUMNA, Renglon + 2.5, False, Trim(cadena6)
'        Imprimir COLUMNA, Renglon + 3, False, Trim(cadena7)
'        Imprimir COLUMNA, Renglon + 3.5, False, Trim(cadena8)
    'Else
    '    cadena1 = cadena
    '    MsgBox cadena1
    'End If
    
End Sub

