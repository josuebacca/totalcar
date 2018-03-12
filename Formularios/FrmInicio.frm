VERSION 5.00
Begin VB.Form FrmInicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Control de Acceso"
   ClientHeight    =   2640
   ClientLeft      =   3870
   ClientTop       =   2415
   ClientWidth     =   4440
   Icon            =   "FrmInicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2768.897
   ScaleMode       =   0  'User
   ScaleWidth      =   2197.8
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   150
      TabIndex        =   4
      Top             =   60
      Width           =   4140
      Begin VB.TextBox TxtClave 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1770
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   2115
      End
      Begin VB.TextBox TxtUsuario 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   0
         Top             =   555
         Width           =   2115
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   315
         Picture         =   "FrmInicio.frx":27A2
         Top             =   765
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese su nombre de Usuario y Clave"
         Height          =   315
         Index           =   2
         Left            =   555
         TabIndex        =   8
         Top             =   255
         Width           =   2925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Left            =   1125
         TabIndex        =   6
         Top             =   1140
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   5
         Top             =   615
         Width           =   585
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "FrmInicio.frx":2AAC
      Height          =   750
      Left            =   3105
      Picture         =   "FrmInicio.frx":2DB6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "FrmInicio.frx":30C0
      Height          =   750
      Left            =   1890
      Picture         =   "FrmInicio.frx":33CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Index           =   1
      Left            =   105
      TabIndex        =   7
      Top             =   2055
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "FrmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CUANTAS_VECES As Integer
Dim rec As ADODB.Recordset
Dim sql As String

Public Sub Conexion()
Dim DSN_DEF As String
    Screen.MousePointer = vbHourglass
    CONECCION = False

    On Error GoTo ErrorIni
    LeoIni
    
    On Error GoTo ErrorTrans
    'ME CONECTO !
    Set DBConn = New ADODB.Connection
'    mNomUser = TxtUsuario.Text
'    mPassword = TxtClave.Text
'    DSN_DEF = "TOTALCAR"
   ' DBConn.ConnectionString = "ODBC;DATABASE=;UID=" & mNomUser & ";PWD=" & mPassword & ";DSN=" & DSN_DEF
    DBConn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    'DBConn.ConnectionString = "driver={SQL Server}; server=DANIEL;database=TOTALCAR"

    'DBConn.ConnectionTimeout = 30       'Default msado10.hlp => 15
    DBConn.CommandTimeout = 0          'Default msado10.hlp => 30
    'DBConn.Open , TxtUsuario, TxtClave
    DBConn.Open DBConn.ConnectionString, TxtUsuario, TxtClave
    
       
    If DBConn.State = adStateOpen Then CONECCION = True
        
    PERMISOS mNomUser
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrorTrans:
        Screen.MousePointer = vbNormal
        MsgBox "No se pudo establecer la conecci�n a la Base de Datos." & Chr(13) & Err.Description, vbExclamation, TIT_MSGBOX
        Exit Sub
ErrorIni:
        Screen.MousePointer = vbNormal
        MsgBox "No se pudo leer el archivo de configuraci�n del sistema." & Chr(13) & Err.Description, vbExclamation, TIT_MSGBOX
End Sub
Public Sub PERMISOS(USUARIO As String)

    Dim sql As String
    Dim r As ADODB.Recordset
    Dim I As Integer

    Set r = New ADODB.Recordset

    On Error Resume Next

    If Trim(USUARIO) = "A" Then
        'Este usuario tiene derecho a todo
        For I = 0 To MENU.Controls.Count - 1
            If TypeName(MENU.Controls(I)) = "Menu" Then
               MENU.Controls(I).Enabled = True
            End If
        Next
    Else
        For I = 0 To MENU.Controls.Count - 1
            If TypeName(MENU.Controls(I)) = "Menu" Then
               MENU.Controls(I).Enabled = False
            End If
        Next
    
        On Error GoTo 0
    
        sql = "SELECT * FROM PERMISOS WHERE " & _
        "USU_NOMBRE = '" & Trim(USUARIO) & "' AND " & _
        "PRM_SISTEMA = '" & Trim(App.Title) & "'"
        r.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If r.RecordCount > 0 Then
            r.MoveFirst
            Do While Not r.EOF
                For I = 0 To MENU.Controls.Count - 1
                    If TypeName(MENU.Controls(I)) = "Menu" Then
                        If UCase(Trim(MENU.Controls(I).Name)) = UCase(Trim(r!PRM_OPMENU)) Then
                            MENU.Controls(I).Enabled = True
                        End If
                    End If
                Next
                r.MoveNext
            Loop
        End If
        r.Close
    End If
End Sub


Private Sub CmdAceptar_Click()
    
    Set rec = New ADODB.Recordset
    mNomUser = Trim(TxtUsuario)
    
    Conexion
    
    If Not CONECCION Then
        If Err.Description <> "" Then
            MsgBox Err.Description
        End If
            
        CUANTAS_VECES = CUANTAS_VECES + 1
        If CUANTAS_VECES = 4 Then
            End
        End If
        TxtUsuario.SetFocus
        Exit Sub
    End If


    sql = "SELECT * FROM USUARIO WHERE " & _
          "USU_NOMBRE = '" & Trim(TxtUsuario) & "' AND " & _
           "USU_CLAVE = '" & Trim(TxtClave) & "'"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount <> 1 Then
        sql = "La contrase�a de usuario NO ES CORRECTA !" & Chr(13) & Chr(13)
        If CUANTAS_VECES = 3 Then
            sql = sql & "El sistema se cerrar�."
        Else
            sql = sql & "Por favor intentelo nuevamente."
        End If
        MsgBox sql, vbCritical, "Error:"
        If CUANTAS_VECES = 3 Then
            'si ya pifi� 3 veces salgo del Sistema
            CmdSalir_Click
        Else
            TxtClave.SelStart = 0
            TxtClave.SelLength = Len(TxtClave)
            TxtUsuario.SetFocus
            CUANTAS_VECES = CUANTAS_VECES + 1
        End If
    Else
        Label1(1).FontBold = True
        Label1(1).Caption = " Conectando ... "
        Label1(1).Refresh
        'muestro un figureti de coneccion
        mNomUser = Trim(TxtUsuario)
        mPassword = Trim(TxtClave)
        
        'BUSCO SUCURSAL---
        BuscoNroSucursal
        '-----------------
        Unload Me
        Set FrmInicio = Nothing
    End If
End Sub
Private Sub CmdAceptar_GotFocus()
    CmdAceptar.FontBold = True
End Sub

Private Sub CmdSalir_Click()
    End
End Sub

Private Sub CmdSalir_GotFocus()
    CmdSalir.FontBold = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    CUANTAS_VECES = 1
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        CmdAceptar_Click
    End If
End Sub

Private Sub TxtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    'If KeyAscii = vbKeyReturn Then TxtClave.SetFocus
End Sub
