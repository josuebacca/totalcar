VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ABMGastos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM  de Gastos "
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMGastos.frx":0000
      Height          =   735
      Left            =   5370
      Picture         =   "ABMGastos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3450
      Width           =   885
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMGastos.frx":0614
      Height          =   735
      Left            =   6270
      Picture         =   "ABMGastos.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3450
      Width           =   885
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "ABMGastos.frx":0C28
      Height          =   735
      Left            =   3570
      Picture         =   "ABMGastos.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3450
      Width           =   885
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMGastos.frx":123C
      Height          =   735
      Left            =   4470
      Picture         =   "ABMGastos.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3450
      Width           =   885
   End
   Begin TabDlg.SSTab TabGastos 
      Height          =   3330
      Left            =   45
      TabIndex        =   6
      Top             =   60
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   5874
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "ABMGastos.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMGastos.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "framegastos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame framegastos 
         Caption         =   "   Gastos   "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         Left            =   -74895
         TabIndex        =   11
         Top             =   465
         Width           =   6915
         Begin VB.CommandButton cmdBuscaGto 
            Height          =   330
            Left            =   6270
            Picture         =   "ABMGastos.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   435
            Width           =   375
         End
         Begin VB.TextBox TxtBuscaDes 
            Height          =   315
            Left            =   1155
            MaxLength       =   60
            TabIndex        =   12
            Top             =   450
            Width           =   4980
         End
         Begin MSFlexGridLib.MSFlexGrid fgBuscaGto 
            Height          =   1710
            Left            =   165
            TabIndex        =   15
            Top             =   900
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   3016
            _Version        =   393216
            FixedCols       =   0
            BackColorSel    =   8388736
            AllowBigSelection=   -1  'True
            FocusRect       =   0
            SelectionMode   =   1
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   210
            TabIndex        =   14
            Top             =   465
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Gasto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2220
         Left            =   150
         TabIndex        =   7
         Top             =   630
         Width           =   6870
         Begin VB.TextBox txtDescri 
            Height          =   315
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            Top             =   1200
            Width           =   5505
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Left            =   1095
            TabIndex        =   0
            Top             =   630
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   135
            TabIndex        =   9
            Top             =   1215
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   420
            TabIndex        =   8
            Top             =   660
            Width           =   540
         End
      End
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
      Height          =   240
      Left            =   150
      TabIndex        =   10
      Top             =   3615
      Width           =   750
   End
End
Attribute VB_Name = "ABMGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim Suma As Double

Private Sub CmdBorrar_Click()
 On Error GoTo HayError
    If TxtCodigo.Text <> "" Then
        If MsgBox("Seguro que desea borrar el gasto?", vbQuestion + vbYesNo + vbDefaultButton2, "Borrar...") = vbYes Then
            DBConn.BeginTrans
             
            sql = "DELETE FROM GASTOS WHERE  GTS_CODIGO=" & XN(TxtCodigo)
            DBConn.Execute sql
            
            DBConn.CommitTrans
            cmdNuevo_Click
        End If
    End If
    Exit Sub
HayError:
    DBConn.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub cmdBuscaGto_Click()
    fgBuscaGto.Rows = 1
    sql = "SELECT * FROM GASTOS"
    sql = sql & " WHERE GTS_DESCRI LIKE '" & Trim(TxtBuscaDes) & "%'"
    sql = sql & " ORDER BY GTS_DESCRI"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
     rec.MoveFirst
     Do While rec.EOF = False
       fgBuscaGto.AddItem rec!GTS_CODIGO & Chr(9) & rec!GTS_DESCRI
       rec.MoveNext
     Loop
    Else
     MsgBox "No se encontraron Gastos!!!", vbExclamation, "Busqueda..."
     TxtBuscaDes.Text = ""
     TxtBuscaDes.SetFocus
     Exit Sub
    End If
    rec.Close
    fgBuscaGto.SetFocus
End Sub

Private Function ValidarDatos() As Boolean
  If txtdescri.Text = "" Then
   MsgBox "Debe ingresar la descripción del gasto", vbCritical, "Error..."
   txtdescri.SetFocus
   ValidarDatos = False
  Else
   ValidarDatos = True
  End If
End Function

Private Sub CmdGrabar_Click()

  On Error GoTo CUALQUIERA
  
  If ValidarDatos = False Then Exit Sub
  lblEstado.Caption = "Cargando Datos....."
  DBConn.BeginTrans
   
   If TxtCodigo.Text = "" Then 'POR ACA REALIZO UNA INSERTION
     sql = "SELECT MAX(GTS_CODIGO) AS MAXIMO FROM GASTOS"
     rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
     If IsNull(rec!Maximo) Then
      TxtCodigo.Text = "1"
     Else
      TxtCodigo.Text = (rec!Maximo) + 1
     End If
     rec.Close
     
     sql = "INSERT INTO GASTOS (GTS_CODIGO,GTS_DESCRI)"
     sql = sql & " VALUES (" & XN(TxtCodigo) & "," & XS(txtdescri) & ")"
     DBConn.Execute sql
      
     DBConn.CommitTrans
     cmdNuevo_Click
   Else 'ACA REALIZO LA MIDIFICACIÓN
     
     sql = "UPDATE GASTOS SET GTS_DESCRI=" & XS(txtdescri)
     sql = sql & " WHERE GTS_CODIGO=" & XN(TxtCodigo)
     DBConn.Execute sql
    
     DBConn.CommitTrans
     cmdNuevo_Click
   End If
    lblEstado.Caption = ""
    Exit Sub

CUALQUIERA:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    MsgBox Err.Description
End Sub

Private Sub cmdNuevo_Click()
    TabGastos.Tab = 0
    TxtCodigo.Text = ""
    txtdescri.Text = ""
    TxtCodigo.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set ABMGastos = Nothing
End Sub

Private Sub fgBuscaGto_DblClick()
      TxtCodigo.Text = fgBuscaGto.TextMatrix(fgBuscaGto.RowSel, 0)
      txtdescri.Text = fgBuscaGto.TextMatrix(fgBuscaGto.RowSel, 1)
      TabGastos.Tab = 0
      txtdescri.SetFocus
End Sub

Private Sub fgBuscaGto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then fgBuscaGto_DblClick
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF1 Then TabGastos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
 Set rec = New ADODB.Recordset
 
 'CONFIGURO LA GRILLA DE BUSQUEDA DE GASTOS
    fgBuscaGto.FormatString = "<Código|Descricpción"
    fgBuscaGto.ColWidth(0) = 1000
    fgBuscaGto.ColWidth(1) = 5000
    fgBuscaGto.Rows = 2
 '-----------------------------------------
   TabGastos.Tab = 0
   lblEstado.Caption = ""
   cmdGrabar.Enabled = False
   cmdBorrar.Enabled = False
End Sub

Private Sub TabGastos_Click(PreviousTab As Integer)
   If TabGastos.Tab = 0 And Me.Visible Then
    cmdGrabar.Enabled = True
    cmdBorrar.Enabled = True
    fgBuscaGto.Rows = 1
    txtdescri.SetFocus
   Else
    If Me.Visible = True Then TxtBuscaDes.SetFocus
    cmdGrabar.Enabled = False
    cmdBorrar.Enabled = False
   End If
End Sub

Private Sub TxtBuscaDes_GotFocus()
    SelecTexto TxtBuscaDes
End Sub

Private Sub TxtBuscaDes_KeyPress(KeyAscii As Integer)
    CarTexto KeyAscii
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If TxtCodigo <> "" Then
        sql = "SELECT * FROM GASTOS"
        sql = sql & " WHERE GTS_CODIGO=" & XN(TxtCodigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtdescri.Text = rec!GTS_DESCRI
            cmdBorrar.Enabled = True
        Else
            MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
            TxtCodigo.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtDescri_Change()
    If txtdescri.Text = "" Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

