Attribute VB_Name = "Funciones"
Option Explicit

Sub DeshabilitarMenu(Frm As Form)
 Dim I As Integer
    For I = 0 To MENU.Controls.Count - 1
        If TypeName(MENU.Controls(I)) = "Menu" And UCase(Left(Frm.Controls(I).Name, 7)) <> "MNURAYA" Then
           MENU.Controls(I).Enabled = False
        End If
    Next
End Sub

Public Sub EliminarFilasDeGrilla(Grilla As MSFlexGrid)
'Borra totalmente una grilla sin dejar Rows
    Grilla.Rows = 1
    Do While Grilla.Rows > 1
        Grilla.RemoveItem 1
    Loop
End Sub

Public Sub LimpiarFilasDeGrilla(Grilla As MSFlexGrid, Optional Fila As Long)
   Dim I As Integer
   Dim j As Integer
'Borra datos de una o todas las filas
    Dim CantidadDeFilas As Integer
    Dim ActFila As Integer
    ActFila = Grilla.row
    If Fila <> 0 Then
        CantidadDeFilas = Fila + 1
    Else
        Fila = 1
        CantidadDeFilas = Grilla.Rows
    End If
    For I = Fila To CantidadDeFilas - 1
        Grilla.row = I
        For j = 0 To Grilla.Cols - 1
            Grilla.Col = j
            Grilla.Text = ""
        Next
    Next
    Grilla.row = ActFila
End Sub

Public Function ValidarPorcentaje(Control As Control) As Boolean
    If CDbl(Control.Text) <= 100 Then
       Control.Text = Format(Control, "0.00")
       ValidarPorcentaje = True
    Else
       MsgBox "Error, Porcentaje mayor al 100%", 16, TIT_MSGBOX
       Control.SetFocus
       SelecTexto Control
       ValidarPorcentaje = False
    End If
End Function

Public Function Valido_Importe(mTEXTO As String) As String
    Valido_Importe = IIf(Trim(mTEXTO) = "", "0,00", Format(mTEXTO, "#,##0.00"))
End Function

Public Function SeteoImpresora(Papel As Integer, Orientacion As Integer, Modo_Escala As Integer, Calidad_Impresion As Integer, Fuente As String, Fuente_Tamano As Integer, Fuente_Negrita As Boolean, Ancho_Impresion As Integer, Largo_Impresion As Double, Optional Ancho_Recuadro As Double) As Boolean
    '*************************************************************************************
    'A) CHEQUES:
    '   1) Llamada a la función: SeteoImpresora(256, 1, 7, -1, "Roman 10cpi", 10, False, 12220, 7950)
    '   2) Configuración de la impresora:
    '       * Driver              : Generic IBM Graphics 9pin
    '       * Paper Size          : Custom
    '               Width         : 1560
    '               Length        : 710
    '               Unit          : 0.1 milimeters
    '
    'B) OBLEAS
    '   1) LLamada a la función: SeteoImpresora(256, 2, 6, -4, "Times New Roman", 8, False, 0, 0, 3)
    '   2) Configuración de la impresora:
    '       * Paper Size          : CUSTOM
    '             Width           : 560
    '             Lenght          : 1000
    '             Unit            : 0,1 MILIMETERS
    '       * Orientation         : LANDSCAPE
    '       * Paper Source        : TOF BACKUP ENABLED
    '       * Media Choice        : SPEED 1.5 TIPS
    '       * Graphics Resolutions: 200 DPI
    '       * Dithering           : NONE
    '       * Intensity           : 100
    '       * Print Quality       : DENSITY 8
    '
    'C) TALON DE CONTROL
    '   1)LLamado a la funcion:Call SeteoImpresora(256, 1, 6, -1, "Roman 10cpi", 10, False, 220, 75)
    '   2) Configuración de la impresora:
    '       * Driver              : Generic IBM Graphics 9pin
    '       * Paper Size          : Custom
    '               Width         :
    '               Length        :
    '               Unit          : 0.1 milimeters
    '
    '***************************************************************************************
    On Error GoTo ErrorPrint
    Printer.PaperSize = Papel
    'Constante        Valor  Descripción
    'vbPRPSLetter       1    Carta, 216 x 279 mm
    'vbPRPSLegal        5    Oficio, 216 x 356 mm
    'vbPRPSA4           9    A4, 210 x 297 mm
    'vbPRPSUser        256   Definido por el usuario

    Printer.Orientation = Orientacion
    'Constante                   Descripción
    'VtOrientationHorizontal     El texto se muestra horizontalmente.
    'VtOrientationVertical       Las letras del texto se dibujan una encima de otra de arriba a abajo.
    'VtOrientationUp             El texto se rota para que se lea de abajo a arriba.
    'VtOrientationDown           El texto se rota para que se lea de arriba a abajo.
    
    Printer.ScaleMode = Modo_Escala
    'Constante    Valor   Descripción
    'vbUser         0      Indica que una o más de las propiedades ScaleHeight, ScaleWidth, ScaleLeft y ScaleTop tienen valores personalizados.
    'VbTwips        1      (Predeterminado) Twip (1440 twips por pulgada lógica; 567 twips por centímetro lógico).
    'VbPoints       2      Punto (72 puntos por pulgada lógica).
    'VbPixels       3      Píxel (la unidad mínima de la resolución del monitor o la impresora).
    'vbCharacters   4      Carácter (horizontal = 120 twips por unidad; vertical = 240 twips por unidad).
    'VbInches       5      Pulgada.
    'VbMillimeters  6      Milímetro.
    'VbCentimeters  7      Centímetro.
    
    Printer.PrintQuality = Calidad_Impresion
    'Constante     Valor   Descripción
    'vbPRPQDraft    -1      Resolución borrador
    'vbPRPQLow      -2      Resolución baja
    'vbPRPQMedium   -3      Resolución media
    'vbPRPQHigh     -4      Resolución alta

    Printer.Font = Fuente
    Printer.FontSize = Fuente_Tamano
    
    Printer.FontBold = Fuente_Negrita
    'Valor   Descripción
    'True    Activa el formato de negrita.
    'False   (Predeterminado) Desactiva el formato de negrita.
    
    If Largo_Impresion > 0 Then
       Printer.Height = Largo_Impresion
    End If
    If Ancho_Impresion > 0 Then
       Printer.Width = Ancho_Impresion
    End If
    If Not IsNull(Ancho_Recuadro) And Ancho_Recuadro > 0 Then
       Printer.DrawWidth = Ancho_Recuadro
    End If
    SeteoImpresora = True
    'AjI = 0
    On Error GoTo 0
    Exit Function
    
ErrorPrint:
    SeteoImpresora = False
    On Error GoTo 0
End Function

Public Function AgregoCtaCteCliente(Cliente As String, TipoCom As String, _
                                    NroComp As String, NroSuc As String, _
                                    FechaComp As String, _
                                    TotalCom As String, DebHab As String, FechaCtaCTe As String) As String
                                    
    'ACTUALIZO LA CUENTA CORRIENTE DEL CLIENTE
    sql = "INSERT INTO CTA_CTE_CLIENTE"
    sql = sql & "(CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,COM_FECHA,"
    sql = sql & "COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,"
    sql = sql & "CTA_CTE_DH,CTA_CTE_FECHA,COM_NUMEROTXT)"
    sql = sql & " VALUES ("
    sql = sql & XN(Cliente) & ","
    sql = sql & XN(TipoCom) & ","
    sql = sql & XN(NroComp) & ","
    sql = sql & XN(NroSuc) & ","
    sql = sql & XDQ(FechaComp) & ","
    sql = sql & XN(TotalCom) & ","
    If DebHab = "D" Then
        sql = sql & XN(TotalCom) & ","
        sql = sql & "0.00" & ","
    Else
        sql = sql & "0.00" & ","
        sql = sql & XN(TotalCom) & ","
    End If
    sql = sql & XS(DebHab) & ","
    sql = sql & XDQ(FechaCtaCTe) & ","
    sql = sql & XS(Format(NroComp, "00000000")) & ")"
    
    AgregoCtaCteCliente = sql
End Function

Public Function QuitoCtaCteCliente(Cliente As String, TipoCom As String, _
                                    NroComp As String, NroSuc As String) As String
    'BORO DE LA CUENTA CORRIENTE DEL CLIENTE
    sql = "DELETE FROM CTA_CTE_CLIENTE"
    sql = sql & " WHERE"
    sql = sql & " CLI_CODIGO=" & XN(Cliente)
    sql = sql & " AND TCO_CODIGO=" & XN(TipoCom)
    sql = sql & " AND COM_NUMERO=" & XN(NroComp)
    sql = sql & " AND COM_SUCURSAL=" & XN(NroSuc)
    QuitoCtaCteCliente = sql
End Function

Public Function AgregoCtaCteProveedores(TipoProv As String, Proveedor As String, TipoCom As String, _
                                    NroSuc As String, NroComp As String, FechaComp As String, _
                                    TotalCom As String, DebHab As String, FechaCtaCTe As String) As String
                                    
    'ACTUALIZO LA CUENTA CORRIENTE DEL PROVEEDOR
    sql = "INSERT INTO CTA_CTE_PROVEEDORES"
    sql = sql & "(TPR_CODIGO,PROV_CODIGO,TCO_CODIGO,COM_SUCURSAL,COM_NUMERO,"
    sql = sql & "COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,"
    sql = sql & "CTA_CTE_DH,CTA_CTE_FECHA)"
    sql = sql & " VALUES ("
    sql = sql & XN(TipoProv) & ","
    sql = sql & XN(Proveedor) & ","
    sql = sql & XN(TipoCom) & ","
    sql = sql & XS(NroSuc) & ","
    sql = sql & XS(NroComp) & ","
    sql = sql & XDQ(FechaComp) & ","
    sql = sql & XN(TotalCom) & ","
    If DebHab = "D" Then
        sql = sql & XN(TotalCom) & ","
        sql = sql & "0.00" & ","
    Else
        sql = sql & "0.00" & ","
        sql = sql & XN(TotalCom) & ","
    End If
    sql = sql & XS(DebHab) & ","
    sql = sql & XDQ(FechaCtaCTe) & ")"
    AgregoCtaCteProveedores = sql
End Function

Public Function QuitoCtaCteProveedores(TipoProv As String, Proveedor As String, TipoCom As String, _
                                    NroSuc As String, NroComp As String) As String
    'BORO DE LA CUENTA CORRIENTE DEL CLIENTE
    sql = "DELETE FROM CTA_CTE_PROVEEDORES"
    sql = sql & " WHERE"
    sql = sql & " TPR_CODIGO=" & XN(TipoProv)
    sql = sql & " AND PROV_CODIGO=" & XN(Proveedor)
    sql = sql & " AND TCO_CODIGO=" & XN(TipoCom)
    sql = sql & " AND COM_SUCURSAL=" & XS(NroSuc)
    sql = sql & " AND COM_NUMERO=" & XS(NroComp)
    QuitoCtaCteProveedores = sql
End Function

Public Sub BuscoEstado(Codigo As Integer, Control As Label)
    Set Rec4 = New ADODB.Recordset
    sql = "SELECT EST_DESCRI"
    sql = sql & " FROM ESTADO_DOCUMENTO"
    sql = sql & " WHERE"
    sql = sql & " EST_CODIGO=" & Codigo
    Rec4.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec4.EOF = False Then
        Control.Caption = Rec4!EST_DESCRI
        
        Select Case Codigo
        Case 2
            Control.ForeColor = &HFF&
        Case 1
            Control.ForeColor = &HFF0000
        Case 3
            Control.ForeColor = &H0&
        Case Else
            Control.ForeColor = &HFF0000
        End Select
    End If
    Rec4.Close
End Sub

Public Sub BuscoNroSucursal()
    Set rec = New ADODB.Recordset
    sql = "SELECT SUCURSAL FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Sucursal = Format(rec!Sucursal, "0000")
    End If
    rec.Close
End Sub
