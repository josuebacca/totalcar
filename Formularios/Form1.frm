VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Remito"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   450
      Left            =   6000
      TabIndex        =   3
      Top             =   4260
      Width           =   930
   End
   Begin VB.Frame fraMedicamentos 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4110
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   7245
      Begin VB.TextBox txtEdit 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   270
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid grdGrilla 
         Height          =   3570
         Left            =   105
         TabIndex        =   2
         Top             =   225
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   6297
         _Version        =   393216
         Rows            =   3
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorSel    =   12583104
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         HighLight       =   0
         FormatString    =   ""
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If Me.ActiveControl.Name <> "grdGrilla" And _
'        KeyAscii = vbKeyEscape Then
'        cmdsalir_Click
'    End If
    If Me.ActiveControl.Name <> "grdGrilla" And _
        Me.ActiveControl.Name <> "txtEdit" And _
        KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()
    grdGrilla.FormatString = "Código|Descripción|Precio|Cantidad|Total"
    grdGrilla.ColWidth(0) = 1000
    grdGrilla.ColWidth(1) = 3000
    grdGrilla.ColWidth(2) = 1000
    grdGrilla.ColWidth(3) = 900
    grdGrilla.ColWidth(4) = 1000
    grdGrilla.Cols = 5
    grdGrilla.Rows = 9
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1
            LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
            grdGrilla.Col = 0
'        Case Else
'            grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, grdGrilla.Col)) = ""
        End Select
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
       (grdGrilla.Col = 2) Or _
       (grdGrilla.Col = 3) Then
        If KeyAscii = vbKeyReturn Then
            If grdGrilla.Col = 3 Then
                If grdGrilla.row < grdGrilla.Rows - 1 Then
                    grdGrilla.row = grdGrilla.row + 1
                    grdGrilla.Col = 0
                Else
                    SendKeys "{TAB}"
                End If
            Else
                grdGrilla.Col = grdGrilla.ColSel + 1 '3
            End If
        Else
            EDITAR grdGrilla, txtEdit, KeyAscii
        End If
    End If
End Sub

Private Sub grdGrilla_LeaveCell()
    If txtEdit.Visible = False Then Exit Sub
    'If Trim(TxtEdit) = "" Then TxtEdit = "0"
    If grdGrilla.Col = 2 Then
        grdGrilla = Format(txtEdit.Text, "0.00")
    Else
        grdGrilla = txtEdit.Text
    End If
    txtEdit.Visible = False
End Sub
Private Sub grdGrilla_GotFocus()
    If grdGrilla.Rows > 1 Then
        If txtEdit.Visible = False Then Exit Sub
        If grdGrilla.Col = 2 Then
            grdGrilla = Format(txtEdit.Text, "0.00")
        Else
            grdGrilla = txtEdit.Text
        End If
        txtEdit.Visible = False
    End If
End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    If grdGrilla.Col = 0 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 2 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    If grdGrilla.Col = 3 Then KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 Then
        frmBuscar.TipoBusqueda = 5
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        frmBuscar.Show vbModal
    End If
'
'    If KeyCode = vbKeyReturn Then
'        Select Case grdGrilla.Col
'        Case 0, 1
'            If Trim(txtEdit) = "" Then
'                txtEdit = ""
'                LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
'                grdGrilla.Col = 0
'                grdGrilla.SetFocus
'                Exit Sub
'            End If
'            Screen.MousePointer = vbHourglass
'            Set rec = New ADODB.Recordset
'        'CAMBIOS SOLICITADOS EL DIA 20-05-2002
'        'SE CREO UNA TABLA (PRECIO_MEDICAMENTO) QUE CONTIENE LOS PRECIOS Y FECHA
'        'DE CUATRO QUINCENAS PARA ATRAS (TOMANDO LA FECHA DE ULTIMA ACTUALIZACION DE MEDICA.)
'        'SE TIENE EN CUENTA LA FECHA DE DISPENSACION Y SEGUEN ESTA USO EL PRECIO DEL MEDICAMENTO
'        If fecDispensacion.Text = "" Then
'           MsgBox "Debe ingresar la fecha de dispensación", vbExclamation, "Medicamentos..."
'           fecDispensacion.SetFocus
'           Exit Sub
'        End If
'            sql = "SELECT MED.MED_TROQUEL, MED.MED_SECUEN, MED.MED_DESCRIP, MED.MED_PRECIO AS PRECIO, NULL AS CANTIDAD, MTV.MTV_CODIGO, MTV.MTV_DESCRIP, MED.OMO_CODIGO"
'            'AGREGADO DE LA TABLA PRECIO_MEDICAMENTO
'            sql = sql & ",PM.MED_FECHA1,PM.MED_PRECIO1,PM.MED_FECHA2,PM.MED_PRECIO2,PM.MED_FECHA3,PM.MED_PRECIO3,PM.MED_FECHA4,PM.MED_PRECIO4"
'            sql = sql & " FROM MEDICAMENTO MED, MED_TIPOVENTA MTV, PRECIO_MEDICAMENTO PM"
'            sql = sql & " WHERE MED.MTV_CODIGO = MTV.MTV_CODIGO "
'
'            sql = sql & " AND MED.MED_TROQUEL=PM.MED_TROQUEL AND MED.MED_SECUEN=PM.MED_SECUEN"
'
'            If grdGrilla.Col = 0 Then
'                If InStr(txtEdit, "-") > 0 Then
'                    sql = sql & " AND MED.MED_TROQUEL=" & Mid(txtEdit, 1, Len(txtEdit) - 2) & " AND MED.MED_SECUEN=" & Mid(txtEdit, Len(txtEdit), 1)
'                Else
'                    sql = sql & " AND MED.MED_TROQUEL=" & Trim(txtEdit)
'                End If
'            Else
'                sql = sql & " AND MED.MED_DESCRIP LIKE '" & Trim(txtEdit) & "%'"
'            End If
'            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'            If rec.RecordCount > 0 Then
'                If rec.RecordCount > 1 And LiquidacionAutomatica = False Then
'                    grdGrilla.SetFocus
'                    frmBuscar.TipoBusqueda = 5
'                    frmBuscar.TxtDescriB.Text = txtEdit.Text
'                    frmBuscar.FechaDispensa = CDate(fecDispensacion.Text)
'                    frmBuscar.Show vbModal
'                    grdGrilla.Col = 0
'                    EDITAR grdGrilla, txtEdit, 13
'                    If Trim(frmBuscar.grdBuscar.Text) <> "" Then txtEdit.Text = frmBuscar.grdBuscar.Text
'                    TxtEdit_KeyDown vbKeyReturn, 0
'                    Exit Sub
'                Else
'                    If grdGrilla.Col = 0 Then
'                        txtEdit = Trim(rec!MED_TROQUEL) & "-" & Trim(rec!MED_SECUEN)
'                    Else
'                        txtEdit = Trim(rec!MED_DESCRIP)
'                    End If
'                    BanderaTeli = True
'                    CargaMedicamentoEnGrilla rec, grdGrilla.row
'                    grdGrilla.Col = 3
'                End If
'            Else
'                If LiquidacionAutomatica Then
'                    txtEdit = ""
'                    'txtNroOrden.Tag = txtNroOrden.Tag Or 1
'                    'MOTIVORECHAZO = MOTIVORECHAZO & " MEDICAMENTO INEXISTENTE "
'                Else
'                    MsgBox "No se ha encontrado medicamento"
'                    txtEdit.Text = ""
'                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
'                    grdGrilla.Col = 0
'                End If
'            End If
'            rec.Close
'            Set rec = Nothing
'            Screen.MousePointer = vbNormal
'        Case 2
'            If Trim(txtEdit) = "" Then txtEdit = grdGrilla.Text
'            Set rec = New ADODB.Recordset
'
'            sql = "SELECT MED_PRECIO AS PRECIO FROM MEDICAMENTO WHERE "
'            sql = sql & " MED_TROQUEL=" & Mid(grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, 0)), 1, Len(grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, 0))) - 2)
'            sql = sql & " AND MED_SECUEN = " & Mid(grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, 0)), Len(grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, 0))), 1)
'
'            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'            If CDbl(txtEdit) > CDbl(Importe) Then
'                MarcarObservacionMedicamento 4, grdGrilla.row
'                grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, 4)) = Format(CDbl(txtEdit) * CDbl(grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, 3))), "0.00")
'            Else
'                QuitarObservacionMedicamento 4, grdGrilla.row
'                grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, 4)) = Format(CDbl(txtEdit) * CDbl(grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, 3))), "0.00")
'            End If
'            rec.Close
'            Set rec = Nothing
'        Case 3
'            If Trim(txtEdit) = "" Then txtEdit = grdGrilla.Text
'            grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, 4)) = Format(CDbl(txtEdit) * CDbl(grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, 2))), "0.00")
'        End Select
'        grdGrilla.SetFocus
'    End If
'    If KeyCode = vbKeyEscape Then
'       txtEdit.Visible = False
'       grdGrilla.SetFocus
'    End If
End Sub
