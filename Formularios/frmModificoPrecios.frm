VERSION 5.00
Begin VB.Form frmModificoPrecios 
   Caption         =   "Modificar Precios"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   30
      TabIndex        =   29
      Top             =   3360
      Width           =   5055
      Begin VB.OptionButton optPreNeto 
         Caption         =   "Precio Neto"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3000
         TabIndex        =   32
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optPreIva 
         Caption         =   "Precio con IVA"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1080
         TabIndex        =   31
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox chkRedondeo 
         Caption         =   "Aplica Redondeo"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   3360
      TabIndex        =   26
      Top             =   720
      Width           =   1695
      Begin VB.OptionButton optmenos 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   28
         Top             =   200
         Width           =   615
      End
      Begin VB.OptionButton optmas 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   27
         Top             =   200
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   45
      TabIndex        =   25
      Top             =   0
      Width           =   5025
      Begin VB.OptionButton optPCompra 
         Caption         =   "Precio Compra"
         Height          =   240
         Left            =   2895
         TabIndex        =   1
         Top             =   330
         Width           =   1440
      End
      Begin VB.OptionButton optPVenta 
         Caption         =   "Precio Venta"
         Height          =   225
         Left            =   540
         TabIndex        =   0
         Top             =   330
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   720
      Left            =   3300
      Picture         =   "frmModificoPrecios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4125
      Width           =   840
   End
   Begin VB.Frame Frame2 
      Caption         =   "Por..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   45
      TabIndex        =   17
      Top             =   705
      Width           =   3225
      Begin VB.OptionButton OptPorc 
         Caption         =   "Porcentaje (%)"
         Height          =   225
         Left            =   540
         TabIndex        =   2
         Top             =   330
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptPesos 
         Caption         =   "Importe"
         Height          =   240
         Left            =   2160
         TabIndex        =   3
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   720
      Left            =   4155
      Picture         =   "frmModificoPrecios.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4125
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "A..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   30
      TabIndex        =   18
      Top             =   1365
      Width           =   5040
      Begin VB.TextBox txtTodos 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1530
         Width           =   840
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   270
         Left            =   105
         TabIndex        =   13
         Top             =   1560
         Width           =   840
      End
      Begin VB.OptionButton OptRep 
         Caption         =   "Marca"
         Height          =   240
         Left            =   105
         TabIndex        =   10
         Top             =   1155
         Width           =   1335
      End
      Begin VB.OptionButton OptRubro 
         Caption         =   "Rubro"
         Height          =   210
         Left            =   105
         TabIndex        =   7
         Top             =   765
         Width           =   795
      End
      Begin VB.OptionButton OptLinea 
         Caption         =   "Línea"
         Height          =   255
         Left            =   105
         TabIndex        =   4
         Top             =   345
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.TextBox txtRep 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3930
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox txtLinea 
         Height          =   315
         Left            =   3915
         MaxLength       =   10
         TabIndex        =   6
         Top             =   285
         Width           =   855
      End
      Begin VB.ComboBox cboRep 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1125
         Width           =   2160
      End
      Begin VB.ComboBox cboRubro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   705
         Width           =   2160
      End
      Begin VB.ComboBox cboLinea 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   285
         Width           =   2160
      End
      Begin VB.TextBox txtRubro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3930
         MaxLength       =   10
         TabIndex        =   9
         Top             =   690
         Width           =   855
      End
      Begin VB.Label lblTodos 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   1245
         TabIndex        =   22
         Top             =   1575
         Width           =   150
      End
      Begin VB.Label lblRep 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   3750
         TabIndex        =   21
         Top             =   1185
         Width           =   150
      End
      Begin VB.Label lblRub 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   3750
         TabIndex        =   20
         Top             =   735
         Width           =   150
      End
      Begin VB.Label lblLinea 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   3750
         TabIndex        =   19
         Top             =   345
         Width           =   150
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   720
      Left            =   2400
      Picture         =   "frmModificoPrecios.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4125
      Width           =   900
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
      Left            =   60
      TabIndex        =   23
      Top             =   4320
      Width           =   750
   End
End
Attribute VB_Name = "frmModificoPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim codlista As Integer

Private Sub chkRedondeo_Click()
    If chkRedondeo.Value = Checked Then
        optPreIva.Enabled = True
        optPreNeto.Enabled = True
    Else
        optPreIva.Enabled = False
        optPreNeto.Enabled = False
    End If
End Sub

Private Sub CmdAceptar_Click()
    Dim Porc As Double
    Dim TOTAL As String
    Dim PorcC As Double
    Dim TOTALC As String
    Dim I As Integer
    Dim signo As Double
    Dim lIVA As Double
    Dim PRECIVA As String
    
    On Error GoTo SeReclavose
    If MsgBox("¿Confirma la Modificación de los Precios?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Actualizando..."
    DBConn.BeginTrans
    
    
    sql = " SELECT P.PTO_DESCRI,L.LNA_DESCRI, "
    sql = sql & " R.RUB_DESCRI,RE.TPRE_DESCRI,P.PTO_PRECIO,P.PTO_PRECIOC,P.PTO_CODIGO,P.LNA_CODIGO,PTO_PRECIVA "
    sql = sql & " FROM PRODUCTO P,LINEAS L,RUBROS R,TIPO_PRESENTACION RE,LISTA_PRECIO LP"
    ',DETALLE_LISTA_PRECIO D "
    sql = sql & " WHERE P.LNA_CODIGO = L.LNA_CODIGO  "
    sql = sql & " AND P.LIS_CODIGO = LP.LIS_CODIGO "
    'AND D.LIS_CODIGO = LP.LIS_CODIGO"
    sql = sql & " AND P.RUB_CODIGO = R.RUB_CODIGO AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
    sql = sql & " AND P.LIS_CODIGO = " & codlista
    
    If cboLinea.ListIndex <> -1 Then
        sql = sql & " AND L.LNA_CODIGO = " & cboLinea.ItemData(cboLinea.ListIndex)
    ElseIf cboRubro.ListIndex <> -1 Then
        sql = sql & " AND R.RUB_CODIGO = " & cboRubro.ItemData(cboRubro.ListIndex)
    ElseIf cboRep.ListIndex <> -1 Then
        sql = sql & " AND RE.TPRE_CODIGO = " & cboRep.ItemData(cboRep.ListIndex)
    ElseIf OptTodos.Value = True Then
        
    End If
    
    sql = sql & " ORDER BY P.PTO_DESCRI "
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If optmas.Value = True Then
        signo = 1
    End If
    If optmenos.Value = True Then
        signo = -1
    End If
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            'VERIFICO SI ES PORCENTAJE O IMPORTE Y LO CALCULO
            If OptPorc.Value = True Then
                If OptLinea.Value = True Then
                    If optPVenta.Value = True Then
                        Porc = signo * ((CDbl(rec!PTO_PRECIO) * CDbl(txtLinea.Text)) / 100)
                        TOTAL = CDbl(rec!PTO_PRECIO) + Porc
                    Else
                        PorcC = signo * ((CDbl(rec!PTO_PRECIOC) * CDbl(txtLinea.Text)) / 100)
                        TOTALC = CDbl(rec!PTO_PRECIOC) + Porc
                    End If
                ElseIf OptRubro.Value = True Then
                    If optPVenta.Value = True Then
                        Porc = signo * ((CDbl(rec!PTO_PRECIO) * CDbl(txtRubro.Text)) / 100)
                        TOTAL = CDbl(rec!PTO_PRECIO) + Porc
                    Else
                        PorcC = signo * ((CDbl(rec!PTO_PRECIOC) * CDbl(txtRubro.Text)) / 100)
                        TOTALC = CDbl(rec!PTO_PRECIOC) + Porc
                    End If
                ElseIf OptRep.Value = True Then
                    If optPVenta.Value = True Then
                        Porc = signo * ((CDbl(rec!PTO_PRECIO) * CDbl(txtRep.Text)) / 100)
                        TOTAL = CDbl(rec!PTO_PRECIO) + Porc
                    Else
                        PorcC = signo * ((CDbl(rec!PTO_PRECIOC) * CDbl(txtRep.Text)) / 100)
                        TOTALC = CDbl(rec!PTO_PRECIOC) + Porc
                    End If
                ElseIf OptTodos.Value = True Then
                    If optPVenta.Value = True Then
                        Porc = signo * ((CDbl(rec!PTO_PRECIO) * CDbl(txtTodos.Text)) / 100)
                        TOTAL = CDbl(rec!PTO_PRECIO) + Porc
                    Else
                        PorcC = signo * ((CDbl(rec!PTO_PRECIOC) * CDbl(txtTodos.Text)) / 100)
                        TOTALC = CDbl(rec!PTO_PRECIOC) + PorcC
                    End If
                End If
            End If
            If OptPesos.Value = True Then
               If OptLinea.Value = True Then
                    If optPVenta.Value = True Then
                        TOTAL = CDbl(rec!PTO_PRECIO) + signo * (CDbl((txtLinea.Text)))
                    Else
                        TOTALC = CDbl(rec!PTO_PRECIOC) + signo * (CDbl((txtLinea.Text)))
                    End If
               ElseIf OptRubro.Value = True Then
                    If optPVenta.Value = True Then
                        TOTAL = CDbl(rec!PTO_PRECIO) + signo * (CDbl(txtRubro.Text))
                    Else
                        TOTALC = CDbl(rec!PTO_PRECIOC) + signo * (CDbl(txtRubro.Text))
                    End If
               ElseIf OptRep.Value = True Then
                    If optPVenta.Value = True Then
                        TOTAL = CDbl(rec!PTO_PRECIO) + signo * (CDbl(txtRep.Text))
                    Else
                        TOTALC = CDbl(rec!PTO_PRECIOC) + signo * (CDbl(txtRep.Text))
                    End If
               ElseIf OptTodos.Value = True Then
                    If optPVenta.Value = True Then
                        TOTAL = CDbl(rec!PTO_PRECIO) + signo * (CDbl(txtTodos.Text))
                    Else
                        TOTALC = CDbl(rec!PTO_PRECIOC) + signo * (CDbl(txtTodos.Text))
                    End If
               End If
            End If
            'Acá hago lo de los redondeos
            If rec!LNA_CODIGO = 7 Then 'MAQUINARIA
                lIVA = 1.21
            Else
                If rec!LNA_CODIGO = 6 Then
                    lIVA = 1.105
                End If
            End If
            If chkRedondeo.Value = Checked Then
                If optPreIva.Value = True Then
                    'Precio de Venta
                    If optPVenta.Value = True Then
                        TOTAL = TOTAL * lIVA
                        TOTAL = Round(TOTAL, 0)
                        TOTAL = TOTAL / lIVA
                        TOTAL = Valido_Importe(TOTAL)
                    End If
                    'Precio de Compra
                    If optPCompra.Value = True Then
                        TOTALC = TOTALC * lIVA
                        TOTALC = Round(TOTAL, 0)
                        TOTALC = TOTAL / lIVA
                        TOTALC = Valido_Importe(TOTALC)
                    End If
                Else
                    'Precio de Venta
                    If optPVenta.Value = True Then
                        TOTAL = Round(TOTAL, 0)
                        TOTAL = Valido_Importe(TOTAL)
                    End If
                    'Precio de Compra
                    If optPCompra.Value = True Then
                        TOTALC = Round(TOTAL, 0)
                        TOTALC = Valido_Importe(TOTALC)
                    End If
                End If
                
            End If
            
            If codlista <> 0 Then
                 ' GUARDO LOS CAMBIOS EN LA GRILLA Y EN LA TABLA
                 I = 1
                 For I = 1 To FrmListadePrecios.GrdModulos.Rows - 1
                    If FrmListadePrecios.GrdModulos.TextMatrix(I, 0) = rec!PTO_CODIGO Then
                        If optPVenta.Value = True Then
                            FrmListadePrecios.GrdModulos.TextMatrix(I, 5) = Valido_Importe(TOTAL)
                            If FrmListadePrecios.GrdModulos.TextMatrix(I, 2) = "MAQUINARIA" Then
                                PRECIVA = TOTAL * 1.105
                            Else
                                PRECIVA = TOTAL * 1.21
                            End If
                            FrmListadePrecios.GrdModulos.TextMatrix(I, 7) = Valido_Importe(PRECIVA)
                        Else
                            FrmListadePrecios.GrdModulos.TextMatrix(I, 6) = Valido_Importe(TOTALC)
                        End If
                        Exit For
                    End If
                 Next
                 'ACTUALIZO LA TABLA
                 sql = "UPDATE PRODUCTO "
                 If optPVenta.Value = True Then
                    sql = sql & " SET PTO_PRECIO=" & XN(TOTAL)
                    sql = sql & " ,PTO_PRECIVA = " & XN(PRECIVA)
                 Else
                    sql = sql & " SET PTO_PRECIOC=" & XN(TOTALC)
                 End If
                 sql = sql & " WHERE LIS_CODIGO=" & codlista
                 sql = sql & " AND PTO_CODIGO LIKE '" & rec!PTO_CODIGO & "'"
                 DBConn.Execute sql
                
            Else
                'GUARDO LOS CAMBIOS SOLO EN LA GRILLA
                I = 1
                 For I = 1 To FrmListadePrecios.GrdModulos.Rows - 1
                    If FrmListadePrecios.GrdModulos.TextMatrix(I, 0) = rec!PTO_CODIGO Then
                        If optPVenta.Value = True Then
                            FrmListadePrecios.GrdModulos.TextMatrix(I, 5) = Valido_Importe(TOTAL)
                            If FrmListadePrecios.GrdModulos.TextMatrix(I, 2) = "MAQUINARIA" Then
                                PRECIVA = TOTAL * 1.105
                            Else
                                PRECIVA = TOTAL * 1.21
                            End If
                            FrmListadePrecios.GrdModulos.TextMatrix(I, 7) = Valido_Importe(PRECIVA)
                        Else
                            FrmListadePrecios.GrdModulos.TextMatrix(I, 6) = Valido_Importe(TOTALC)
                        End If
                        Exit For
                    End If
                 Next
            End If
            
            rec.MoveNext
        Loop
    Else
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
    End If
    DBConn.CommitTrans
    rec.Close
    Screen.MousePointer = vbNormal
    CmdNuevo_Click
    Exit Sub
    
SeReclavose:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdNuevo_Click()
    cboLinea.ListIndex = -1
    cboRubro.ListIndex = -1
    cboRep.ListIndex = -1
    txtLinea.Text = "0,00"
    txtRubro.Text = "0,00"
    txtRep.Text = "0,00"
    txtTodos.Text = "0,00"
    lblEstado.Caption = ""
    'OptPorc.Value = True
    'OptLinea.Value = True
    If Me.Visible = True Then optPVenta.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set frmModificoPrecios = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Call Centrar_pantalla(Me)
    cargocboLinea
    cargocboRepres
    cargocboRubro
    If FrmListadePrecios.txtcodigo.Text = "" Then
        codlista = 0
    Else
        codlista = Int(FrmListadePrecios.txtcodigo.Text)
    End If
    CmdNuevo_Click
End Sub
Function cargocboLinea()
    cboLinea.Clear
    sql = "SELECT * FROM LINEAS  ORDER BY LNA_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboLinea.AddItem rec!LNA_DESCRI
            cboLinea.ItemData(cboLinea.NewIndex) = rec!LNA_CODIGO
            rec.MoveNext
        Loop
        cboLinea.ListIndex = -1
    End If
    rec.Close
End Function
Function cargocboRubro()
    
    cboRubro.Clear
    sql = "SELECT * FROM RUBROS "
    sql = sql & " ORDER BY RUB_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboRubro.AddItem rec!RUB_DESCRI
            cboRubro.ItemData(cboRubro.NewIndex) = rec!RUB_CODIGO
            rec.MoveNext
        Loop
        cboRubro.ListIndex = -1
    End If
    rec.Close
End Function
Function cargocboRepres()
    cboRep.Clear
    sql = "SELECT * FROM TIPO_PRESENTACION ORDER BY TPRE_DESCRI"
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

Function SeteosDos(num As Integer)
    If num = 1 Then
        cboLinea.Enabled = True
        txtLinea.Enabled = True
        cboRubro.Enabled = False
        txtRubro.Enabled = False
        cboRep.Enabled = False
        txtRep.Enabled = False
        txtTodos.Enabled = False
        txtLinea.SetFocus
        txtRubro.Text = "0,00"
        txtRep.Text = "0,00"
        txtTodos.Text = "0,00"
        
        cboLinea.ListIndex = 0
        cboRep.ListIndex = -1
        cboRubro.ListIndex = -1
        'OptLinea.Value = True
    End If
    If num = 2 Then
        cboLinea.Enabled = False
        txtLinea.Enabled = False
        cboRubro.Enabled = True
        txtRubro.Enabled = True
        cboRep.Enabled = False
        txtRep.Enabled = False
        txtTodos.Enabled = False
        txtRubro.SetFocus
        txtLinea.Text = "0,00"
        txtRep.Text = "0,00"
        txtTodos.Text = "0,00"
        
        cboRubro.ListIndex = 0
        cboRep.ListIndex = -1
        cboLinea.ListIndex = -1
        'OptRubro.Value = True
    End If
    If num = 3 Then
        cboLinea.Enabled = False
        txtLinea.Enabled = False
        cboRubro.Enabled = False
        txtRubro.Enabled = False
        cboRep.Enabled = True
        txtRep.Enabled = True
        txtTodos.Enabled = False
        txtRep.SetFocus
        txtRubro.Text = "0,00"
        txtLinea.Text = "0,00"
        txtTodos.Text = "0,00"
        
        cboRep.ListIndex = 0
        cboLinea.ListIndex = -1
        cboRubro.ListIndex = -1
        
        'OptRep.Value = True
    End If
    If num = 4 Then
        cboLinea.Enabled = False
        txtLinea.Enabled = False
        cboRubro.Enabled = False
        txtRubro.Enabled = False
        cboRep.Enabled = False
        txtRep.Enabled = False
        txtTodos.Enabled = True
        txtTodos.SetFocus
        txtRubro.Text = "0,00"
        txtRep.Text = "0,00"
        txtLinea.Text = "0,00"
        
        cboRep.ListIndex = -1
        cboLinea.ListIndex = -1
        cboRubro.ListIndex = -1
  End If
End Function

Private Sub OptLinea_Click()
    SeteosDos (1)
End Sub

Private Sub OptPesos_Click()
    lblLinea.Caption = "$"
    lblRub.Caption = "$"
    lblRep.Caption = "$"
    lblTodos.Caption = "$"
End Sub

Private Sub OptPorc_Click()
    lblLinea.Caption = "%"
    lblRub.Caption = "%"
    lblRep.Caption = "%"
    lblTodos.Caption = "%"
End Sub

Private Sub OptRep_Click()
    SeteosDos (3)
End Sub

Private Sub OptRubro_Click()
    SeteosDos (2)
End Sub

Private Sub OptTodos_Click()
    SeteosDos (4)
End Sub

Private Sub txtLinea_GotFocus()
    SelecTexto txtLinea
End Sub

Private Sub txtLinea_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtLinea, KeyAscii)
End Sub

Private Sub txtLinea_LostFocus()
    If txtLinea.Text <> "" Then
        If OptPesos.Value = True Then
            txtLinea.Text = Valido_Importe(txtLinea)
        Else
            If ValidarPorcentaje(txtLinea) = False Then
                txtLinea.SetFocus
            End If
        End If
    Else
        txtLinea.Text = "0,00"
    End If
End Sub

Private Sub txtRep_GotFocus()
    SelecTexto txtRep
End Sub

Private Sub txtRep_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtRep, KeyAscii)
End Sub

Private Sub txtRep_LostFocus()
    If txtRep.Text <> "" Then
        If OptPesos.Value = True Then
            txtRep.Text = Valido_Importe(txtRep)
        Else
            If ValidarPorcentaje(txtRep) = False Then
                txtRep.SetFocus
            End If
        End If
    Else
        txtRep.Text = "0,00"
    End If
End Sub

Private Sub txtRubro_GotFocus()
    SelecTexto txtRubro
End Sub

Private Sub txtRubro_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtRubro, KeyAscii)
End Sub

Private Sub txtRubro_LostFocus()
    If txtRubro.Text <> "" Then
        If OptPesos.Value = True Then
            txtRubro.Text = Valido_Importe(txtRubro)
        Else
            If ValidarPorcentaje(txtRubro) = False Then
                txtRubro.SetFocus
            End If
        End If
    Else
        txtRubro.Text = "0,00"
    End If
End Sub

Private Sub txtTodos_GotFocus()
    SelecTexto txtTodos
End Sub

Private Sub txtTodos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtTodos, KeyAscii)
End Sub

Private Sub txtTodos_LostFocus()
    If txtTodos.Text <> "" Then
        If OptPesos.Value = True Then
            txtTodos.Text = Valido_Importe(txtTodos)
        Else
            If ValidarPorcentaje(txtTodos) = False Then
                txtTodos.SetFocus
            End If
        End If
    Else
        txtTodos.Text = "0,00"
    End If
End Sub
