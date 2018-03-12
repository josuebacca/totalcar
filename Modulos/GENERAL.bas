Attribute VB_Name = "GENERAL"
''para el matriculado
'Public mNroDoc As String
'Public mTipDoc As String
'Public mNombre As String
'Public mApellido As String
'
''para el familiar
'Public fNroDoc   As String
'Public fTipDoc   As String
'Public fNombre   As String
'Public fApellido As String
'
''para el Estudios Contables
'Public ecNroDoc   As String
'Public ecTipDoc   As String
'Public ecNombre   As String
'Public ecApellido As String
'Public ecCUIT     As String

''para el garante
'Public gNroDoc   As String
'Public gTipDoc   As String
'Public gNombre   As String
'Public gApellido As String

'Nombre de Usuario
Public mNomUser  As String
Public mPassword As String
Public mNivUser  As String

'Public mRemitoCli As Integer

''Parametro de Consulta
'Public SERVSOCI As Boolean
'
''Medicamentos
'Global m_medicamento As Boolean

'Constante para el titulo de los msgbox
Global Const TIT_MSGBOX  As String * 16 = "Sistema TOTALCAR"

Public Boton_Grabar As Boolean

Global buscobanco As Integer
Global Consulta As Integer ' para la Lista de Precios
Global Stock As Integer ' para el Control y Consulta Stock
Global Sucursal  As String
Global Item      As Integer
Global sql       As String
Global cSQL      As String
Global cSql1     As String
Global cSql2     As String
Global rec       As ADODB.Recordset
Global Rec1      As ADODB.Recordset
Global Rec2      As ADODB.Recordset
Global Rec3      As ADODB.Recordset
Global Rec4      As ADODB.Recordset
Global DBConn    As ADODB.Connection
Global Viene_Form   As Boolean
Global CONECCION    As Boolean
Global NumeroRecibo As Long
Global CLAVE_OK     As Boolean
Global FormLlamado  As String
'Estas variables son usadas en los sist. de la delegacion
'cuando se emiten recibos a matriculados de otra delegacion
'y se encuentran en la funcion ACETO_RECIBO del modulo cajero
Global dNoListo     As Boolean
Global v_frmingtramites As Boolean 'Esta variable es utilizada en el sistema de tecnica para la delegacion cuando llama a frmmatotrdel
Global v_frmenttramites As Boolean
Global viene_de_tecnica As Boolean

'-
'Global Delegacion As Integer
Public Delegacion  As Integer
'-
' Estas Declaraciones sirven para deshabilitar
' las teclas CTRL+ALT+DEL
' Efectos colaterales: posible formateo de disco,
' borrado de todos los archivos jpg, mp3, etc.
Public Const SPI_SCREENSAVERRUNNING = 97&
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
        (ByVal uAction As Long, ByVal uParam As Long, _
        lpvParam As Any, ByVal fuWinIni As Long) As Long
' Fin de la Declaraci�n



Public Function Consulto_FechaNacimiento(TIPO As String, NRO As String) As String

    Dim RecFEC As ADODB.Recordset
    Set RecFEC = New ADODB.Recordset
    Dim cSQLFEC As String

    cSQLFEC = "Select per_fecnac from persona " & _
              "Where tip_tipdoc = " & XS(TIPO) & " and per_nrodoc = " & XS(NRO)
    RecFEC.Open cSQLFEC, DBConn, adOpenStatic, adLockOptimistic
    If RecFEC.EOF = False Then
        Consulto_FechaNacimiento = Trim(RecFEC!per_fecnac)
    Else
        MsgBox "Fecha de Nacimiento Incorrecta.", vbExclamation, TIT_MSGBOX
        Exit Function
    End If
End Function
Public Function CONSULTO_DEUDA_SSociales(TIPO As String, NRO As String, _
                                          FECHADEUDA As String, _
                                          SISTEMA1 As Integer, _
                                          SUBSISTEMA1 As Integer, _
                                 Optional SISTEMA2 As Integer, _
                                 Optional SUBSISTEMA2 As Integer, _
                                 Optional SISTEMA3 As Integer, _
                                 Optional SUBSISTEMA3 As Integer, _
                                 Optional SISTEMA4 As Integer, _
                                 Optional SUBSISTEMA4 As Integer) As String
                                          
'Funci�n que me devuelve la Cantidad de meses de Adeudados
'Para utilizarla incluir en el evento
'Dim Cadena As String
'Dim Cant_meses As Integer
'Cadena = CONSULTO_DEUDA_SSociales(TIP_TIPDOC, PER_NRODOC, Date, SISTEMA1, SUBSISTEMA1, SISTEMA2, SUBSISTEMA2, SISTEMA3, SUBSISTEMA3, SISTEMA4, SUBSISTEMA4), DBConn, adOpenStatic, adLockOptimistic
'Cant_Meses = Cadena
'Autor: Eliana
        Dim RecDEUDA As ADODB.Recordset
        Set RecDEUDA = New ADODB.Recordset
        
        Dim Cant_Meses As String
        
        'Contador de los meses que debe
        Cant_Meses = "0"
        
        RecDEUDA.Open SELECT_DEUDA_SSociales(TIPO, NRO, FECHADEUDA, SISTEMA1, SUBSISTEMA1, SISTEMA2, SUBSISTEMA2, SISTEMA3, SUBSISTEMA3, SISTEMA4, SUBSISTEMA4), DBConn, adOpenStatic, adLockOptimistic
        If RecDEUDA.EOF = False Then
           'RecDEUDA.MoveFirst
           Do While Not RecDEUDA.EOF
              Cant_Meses = Cant_Meses + RecDEUDA.Fields(0)
           RecDEUDA.MoveNext
           Loop
           CONSULTO_DEUDA_SSociales = Cant_Meses
        Else
           'El Matriculado NO adeuda
           CONSULTO_DEUDA_SSociales = 0
        End If
       RecDEUDA.Close
End Function

Public Function SELECT_DEUDA_SSociales(TipoDoc As String, Nrodoc As String, _
                                        Fecha As String, _
                                        mSistema1 As Integer, mSubsistema1 As Integer, _
                               Optional mSistema2 As Integer, _
                               Optional mSubsistema2 As Integer, _
                               Optional mSistema3 As Integer, _
                               Optional mSubsistema3 As Integer, _
                               Optional mSistema4 As Integer, _
                               Optional mSubsistema4 As Integer) As String
        
        cSQL = " SELECT count(distinct(substring(Cta_Cte_SubSist.Cta_NumCed,1,6)))"
        cSQL = cSQL & " From Cta_Cte,Cta_Cte_SubSist,Periodo,Vencimiento V1"
        cSQL = cSQL & " Where Cta_Cte.Tip_TipDoc = " & XS(Trim(TipoDoc))
        cSQL = cSQL & " And   Cta_Cte.Per_NroDoc = " & XS(Trim(Nrodoc))
        cSQL = cSQL & " And   Cta_Cte.Cta_NumCed = Cta_Cte_SubSist.Cta_NumCed "
        cSQL = cSQL & " And   Cta_Cte_SubSist.Ccs_Pagado = 0 "
        cSQL = cSQL & " And   Cta_Cte_SubSist.RUB_CODIGO = 3"
        
        If Trim(mSistema1) <> "" And Trim(mSubsistema1) <> "" Then
            cSQL = cSQL & " And ((Cta_Cte_SubSist.SIS_CODIGO = " & mSistema1 & " And Cta_Cte_SubSist.SUB_CODIGO = " & mSubsistema1 & ")"
        End If
        If Trim(mSistema2) <> "" And Trim(mSubsistema2) <> "" Then
            cSQL = cSQL & " or (Cta_Cte_SubSist.SIS_CODIGO = " & mSistema2 & " And Cta_Cte_SubSist.SUB_CODIGO = " & mSubsistema2 & ")"
        End If
        If Trim(mSistema3) <> "" And Trim(mSubsistema3) <> "" Then
            cSQL = cSQL & " or (Cta_Cte_SubSist.SIS_CODIGO = " & mSistema3 & " And Cta_Cte_SubSist.SUB_CODIGO = " & mSubsistema3 & ")"
        End If
        If Trim(mSistema4) <> "" And Trim(mSubsistema4) <> "" Then
            cSQL = cSQL & " or (Cta_Cte_SubSist.SIS_CODIGO = " & mSistema4 & " And Cta_Cte_SubSist.SUB_CODIGO = " & mSubsistema4 & ")"
        End If
        cSQL = cSQL & ")"
        
        cSQL = cSQL & " And   Cta_Cte.Pdo_Period = Periodo.Pdo_Period "
        cSQL = cSQL & " And   Periodo.Pdo_Period = V1.Pdo_Period "
        cSQL = cSQL & " And   V1.Ven_FecVto = (Select Min(V2.Ven_FecVto) "
        cSQL = cSQL & "                          From Vencimiento V2 "
        cSQL = cSQL & "                         Where V2.Pdo_Period = Cta_Cte.Pdo_Period "
        cSQL = cSQL & "                           And V2.Ven_FecVto >= " & XD(Fecha) & ")"
        cSQL = cSQL & " And " & XD(Fecha) & " > (Select Min(V2.Ven_FecVto)  "
        cSQL = cSQL & "                            From Vencimiento V2 "
        cSQL = cSQL & "                           Where V2.Pdo_Period = Cta_Cte.Pdo_Period)"
        
        cSQL = cSQL & " Union All "
        
        cSQL = cSQL & " SELECT count(distinct(substring(Cta_Cte_SubSist.Cta_NumCed,1,6)))"
        cSQL = cSQL & " From Cta_Cte,Cta_Cte_SubSist,Periodo,Vencimiento V1"
        cSQL = cSQL & " Where Cta_Cte.Tip_TipDoc = " & XS(Trim(TipoDoc))
        cSQL = cSQL & "   And Cta_Cte.Per_NroDoc = " & XS(Trim(Nrodoc))
        cSQL = cSQL & "   And Cta_Cte.Cta_NumCed = Cta_Cte_SubSist.Cta_NumCed "
        cSQL = cSQL & "   And Cta_Cte_SubSist.Ccs_Pagado = 0 "
        cSQL = cSQL & "   And Cta_Cte_SubSist.RUB_CODIGO = 3"
        
        If Trim(mSistema1) <> "" And Trim(mSubsistema1) <> "" Then
            cSQL = cSQL & " And ((Cta_Cte_SubSist.SIS_CODIGO = " & mSistema1 & " And Cta_Cte_SubSist.SUB_CODIGO = " & mSubsistema1 & ")"
        End If
        If Trim(mSistema2) <> "" And Trim(mSubsistema2) <> "" Then
            cSQL = cSQL & " or (Cta_Cte_SubSist.SIS_CODIGO = " & mSistema2 & " And Cta_Cte_SubSist.SUB_CODIGO = " & mSubsistema2 & ")"
        End If
        If Trim(mSistema3) <> "" And Trim(mSubsistema3) <> "" Then
            cSQL = cSQL & " or (Cta_Cte_SubSist.SIS_CODIGO = " & mSistema3 & " And Cta_Cte_SubSist.SUB_CODIGO = " & mSubsistema3 & ")"
        End If
        If Trim(mSistema4) <> "" And Trim(mSubsistema4) <> "" Then
            cSQL = cSQL & " or (Cta_Cte_SubSist.SIS_CODIGO = " & mSistema4 & " And Cta_Cte_SubSist.SUB_CODIGO = " & mSubsistema4 & ")"
        End If
        cSQL = cSQL & ")"
        
        cSQL = cSQL & "   And Cta_Cte.Pdo_Period = Periodo.Pdo_Period "
        cSQL = cSQL & "   And Periodo.Pdo_Period = V1.Pdo_Period "
        cSQL = cSQL & "   And V1.Ven_FecVto = (Select Max(V2.Ven_FecVto) "
        cSQL = cSQL & "                          From Vencimiento V2 "
        cSQL = cSQL & "                         Where V2.Pdo_Period = Cta_Cte.Pdo_Period "
        cSQL = cSQL & "                           And V2.Ven_FecVto < " & XD(Fecha) & " ) "
        cSQL = cSQL & "   And " & XD(Fecha) & " > (Select Max(V2.Ven_FecVto) "
        cSQL = cSQL & "                              From Vencimiento V2 "
        cSQL = cSQL & "                             Where V2.Pdo_Period = Cta_Cte.Pdo_Period)"
        
        SELECT_DEUDA_SSociales = cSQL
End Function

Public Function Consulto_Afiliacion_FdoSolidario(TIPO As String, NRO As String) As Boolean

    Dim RecC As ADODB.Recordset
    Set RecC = New ADODB.Recordset
    Dim cSQLC As String

    cSQLC = "Select f1.fdo_fecdes " & _
            "  from fondo_solidario f1 " & _
            " Where f1.tip_tipdoc = " & XS(TIPO) & _
            "   and f1.per_nrodoc = " & XS(NRO) & _
            "   and f1.fdo_fecafi = (select max(fdo_fecafi) " & _
            "                          from fondo_solidario f2 " & _
            "                         Where f1.tip_tipdoc = f2.tip_tipdoc " & _
            "                           and f1.per_nrodoc = f2.per_nrodoc)"
    RecC.Open cSQLC, DBConn, adOpenForwardOnly, adLockReadOnly
    If RecC.EOF = False Then
       If IsNull(RecC!FDO_FECDES) Then
          Consulto_Afiliacion_FdoSolidario = True
       Else
          Consulto_Afiliacion_FdoSolidario = False
       End If
    Else
       Consulto_Afiliacion_FdoSolidario = False '"NO tiene Afiliaci�n al Fdo. Solidario"
    End If
    RecC.Close
    
End Function

Public Function CONSULTO_TIPO_NRO_DOCUMENTO(Titulo As String, Matricula As String) As String

     'Funci�n que me devuelve el Tipo y Nro de Documento
     'Para utilizarla incluir en el evento
     'Autor: Eliana

     'Dim Renglon As String
     'Dim Tipo    As Integer
     'Dim Nro     As Double

     '   Renglon = CONSULTO_TIPO_NRO_DOCUMENTO(10, 4444)
     '      Tipo = TextoPrevioAlGuion(Renglon)
     '       Nro = TextoPostAlGuion(Renglon)

        Dim RecTIPONRO As ADODB.Recordset
        Set RecTIPONRO = New ADODB.Recordset
        Dim mTitulo    As String
        Dim mMatricula As String

        mTitulo = Format(Titulo, "00")
        mMatricula = Format(Matricula, "00000")

        sql = "SELECT TIP_TIPDOC, PER_NRODOC FROM Titulo " & _
               "WHERE TTI_CODIGO = " & XS(mTitulo) & _
                " AND TIT_MATRIC = " & XS(mMatricula)
        RecTIPONRO.Open sql, DBConn, adOpenForwardOnly, adLockReadOnly
        If RecTIPONRO.EOF = False Then
           CONSULTO_TIPO_NRO_DOCUMENTO = RecTIPONRO!Tip_TipDoc + "-" + RecTIPONRO!Per_NroDoc
        Else
           CONSULTO_TIPO_NRO_DOCUMENTO = "0" + "-" + "0"
        End If
       RecTIPONRO.Close
End Function


Public Sub CentrarW(F As Form)
    Set F = F
    F.Left = (Screen.Width - F.Width) / 2
    F.Top = (Screen.Height - F.Height) / 2
End Sub

Function BuscarPunto(TEXTO As String) As Boolean
    BuscarPunto = False
    
    For I = 1 To Len(TEXTO)
        If Mid$(TEXTO, I, 1) = "." Then
            BuscarPunto = True
            Exit For
        End If
    Next I
                
End Function

Public Function ProximoNumeroRecibo() As Long
   
   Dim SQLRecibo As String
   Dim Recibo As ADODB.Recordset
   
   Set Recibo = New ADODB.Recordset
   
   'Esta Function va dentro de un Begin - Commit
   'Se debe invocar dentro de una transaccion
   SQLRecibo = "SELECT ULT_RECFIC FROM ULTIMOS"
   Recibo.Open SQLRecibo, DBConn, adOpenStatic, adLockOptimistic
   If Recibo.EOF = False Then
      ProximoNumeroRecibo = Recibo!ULT_RECFIC + 1
   Else
      ProximoNumeroRecibo = 1
   End If
   DBConn.Execute "UPDATE ULTIMOS SET ULT_RECFIC = " & ProximoNumeroRecibo
   Recibo.Close
   
End Function

Public Function NumeroEntero(ByRef KeyAscii As Integer) As Integer
    Dim Car As String * 1
    Car = Chr$(KeyAscii)
'    Values 8, 9, 10, and 13 convert to backspace, tab
'           , linefeed, and carriage return characters
'           , respectively
    If (Car < "0" Or Car > "9") And KeyAscii <> 8 And KeyAscii <> 9 And KeyAscii <> 10 And KeyAscii <> 13 Then
        Beep
        NumeroEntero = 0
    Else
        NumeroEntero = KeyAscii
    End If
End Function

Public Function CarNumeroDecimalComaPunto(ByRef TEXTO As String, ByRef KeyAscii As Integer, Optional NEG As Boolean) As Integer
    
    Dim Car As String * 1
        
    Car = Chr$(KeyAscii)
    
    If (Car = "/") Then
        Beep
        CarNumeroDecimalComaPunto = 0
    ElseIf (Car < "." Or Car > "9") And KeyAscii <> 8 And KeyAscii <> 44 And Not NEG Then
        Beep
        CarNumeroDecimalComaPunto = 0
    ElseIf (Car < "." Or Car > "9") And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 45 And NEG Then
        Beep
        CarNumeroDecimalComaPunto = 0
    ElseIf Car = "." Or Car = "," Then
        
        Car = "."
        
        ' verifico si hay un punto decimal ingresado
        If BuscarPunto(TEXTO) Then
            Beep
            CarNumeroDecimalComaPunto = 0
        Else
            HayPunto = True
            CarNumeroDecimalComaPunto = Asc(Car)
        End If
    Else
        CarNumeroDecimalComaPunto = (KeyAscii)
    End If
    
End Function

Public Function CarNumeroDecimal(ByRef TEXTO As String, ByRef KeyAscii As Integer, Optional NEG As Boolean) As Integer
    
    Dim Car As String * 1
        
    Car = Chr$(KeyAscii)
    
    If (Car = "/") Then
        Beep
        CarNumeroDecimal = 0
    ElseIf (Car < "." Or Car > "9") And KeyAscii <> 8 And KeyAscii <> 44 And Not NEG Then
        Beep
        CarNumeroDecimal = 0
    ElseIf (Car < "." Or Car > "9") And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 45 And NEG Then
        Beep
        CarNumeroDecimal = 0
    ElseIf Car = "." Or Car = "," Then
        
        Car = ","
         
        ' verifico si hay un punto decimal ingresado
        If BuscarPunto(TEXTO) Then
            Beep
            CarNumeroDecimal = 0
        Else
            HayPunto = True
            CarNumeroDecimal = Asc(Car)
        End If
    Else
        CarNumeroDecimal = (KeyAscii)
    End If
    
End Function

Public Function CarTexto(ByRef KeyAscii As Integer)



If KeyAscii = 34 Or Chr(KeyAscii) = "," Or Chr(KeyAscii) = "'" Then
        CarTexto = 0
    Else
        CarTexto = Asc(UCase(Chr(KeyAscii)))
    End If
    
End Function

Public Function CarTextoMinuscula(ByRef KeyAscii As Integer)
    If KeyAscii = 34 Or Chr(KeyAscii) = "," Or Chr(KeyAscii) = "'" Then
        CarTextoMinuscula = 0
    Else
        CarTextoMinuscula = Asc(Chr(KeyAscii))
    End If
End Function

Public Function ProximoNumeroCedulon(Periodo As String) As String

   'Genero el N�mero de Cedul�n
   'PERIODO ES STRING DE 6
   'nNroCedulon = Format(ProximoNumeroCedulon(PERIODO), "000000000000")
   
   Dim SQLCedulon As String
   Dim Cedulon As ADODB.Recordset
   Set Cedulon = New ADODB.Recordset
   
   'Esta Function va dentro de un Begin-Commit se debe invocar dentro de una transaccion
   
   SQLCedulon = " SELECT MAX(CTA_NUMCED)AS ULTIMO " & _
                " FROM CTA_CTE" & _
                " WHERE SUBSTRING(CTA_NUMCED,1,6) = " & XS(Periodo)
   Cedulon.Open SQLCedulon, DBConn, adOpenStatic, adLockOptimistic
   If Not IsNull(Cedulon!Ultimo) Then
      ProximoNumeroCedulon = Trim(Str(Val(Cedulon!Ultimo) + 1))
   Else
      ProximoNumeroCedulon = Periodo & "000001"
   End If
   Cedulon.Close
   Set Cedulon = Nothing
End Function

Public Function Mayuscula(ByRef KeyAscii As Integer) As Integer

'    Values 8, 9, 10, and 13 convert to backspace, tab
'           , linefeed, and carriage return characters
'           , respectively
'    If KeyAscii = 8 Then
'        MsgBox "tab"
'    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Or _
        (KeyAscii >= 65 And KeyAscii <= 90) Or _
        (KeyAscii >= 97 And KeyAscii <= 122) Or _
        KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii <> 9 Or KeyAscii <> 10 Or _
        (KeyAscii >= 160 And KeyAscii <= 165) Or _
        KeyAscii = 130 Or KeyAscii = 211 Or KeyAscii = 201 Or _
        KeyAscii = 193 Or KeyAscii = 205 Or KeyAscii = 209 Or KeyAscii = 218 _
        Then
            Mayuscula = Asc(UCase(Chr(KeyAscii)))
    Else
            Mayuscula = 0
    End If
End Function

Function DesHabilitoControles(frmGlobal As Form)
    DesHabilitoControles = False
    For Each Control In frmGlobal.Controls
        If TypeOf Control Is TextBox _
        Or TypeOf Control Is ComboBox _
        Or TypeOf Control Is CheckBox _
        Or TypeOf Control Is MaskEdBox _
        Or TypeOf Control Is ListBox Then
            Control.Enabled = False
            DesHabilitoControles = True
        End If
    Next Control
End Function

Function HabilitoControles(frmGlobal As Form)
    HabilitoControles = False
    For Each Control In frmGlobal.Controls
        If TypeOf Control Is TextBox _
        Or TypeOf Control Is ComboBox _
        Or TypeOf Control Is CheckBox _
        Or TypeOf Control Is MaskEdBox _
        Or TypeOf Control Is ListBox Then
                Control.Enabled = True
                HabilitoControles = True
        End If
    Next Control
End Function

Function ValidarIngreso(frmGlobal As Form) As Boolean
Dim MensCampos As String
ValidarIngreso = True
For Each Control In frmGlobal.Controls ' revisar los controles del form
    If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then    ' si el control es de carga o selecci�n de datos
        If Trim(Control.Tag) <> "" Then  'si el tag no est� vacio, es un campo necesario
            If Trim(Control.Text) = "" Then ' dejaron vacio un campo necesario
                MensCampos = MensCampos & Chr(13) & Control.Tag
                ValidarIngreso = False
            End If
        End If
    End If
Next Control

If MensCampos <> "" Then ' si hay mensaje es que hay campos incompletos
    Beep
    MsgBox "Debe completar los siguientes campos:" & MensCampos, vbOKOnly + vbInformation, TIT_MSGBOX
End If

End Function

Public Function FechaSQL(ByRef pFecha As String) As String
    If pFecha = "" Then 'si es un string blanco
        FechaSQL = "NULL"
    Else
        'si es una fecha valida
        If IsDate(CDate(pFecha)) Then
            FechaSQL = "'" & Format(Month(pFecha), "00") & "/" & Format(Day(pFecha), "00") & "/" & Format(Year(pFecha), "0000") & "'"
        End If
    End If
End Function

Public Function VerificarFecha(Fecha As String) As Boolean

Dim DIA As Integer
Dim MES As Integer
Dim a�o As Integer
Dim nombredelmes As String
    
    VerificarFecha = True
    
    If Val(Fecha) = 0 Then Exit Function
    
    If Len(Fecha) < 8 Then
        Beep
        MsgBox "Fecha Incompleta !", vbExclamation, TIT_MSGBOX
        VerificarFecha = False
        Exit Function
    End If
    
    DIA = Val(Mid(Trim(Fecha), 1, 2))
    MES = Val(Mid(Trim(Fecha), 3, 2))
    a�o = Val(Mid(Trim(Fecha), 5, 4))
    
    If MES = 1 Then nombredelmes = "Enero"
    If MES = 3 Then nombredelmes = "Marzo"
    If MES = 4 Then nombredelmes = "Abril"
    If MES = 5 Then nombredelmes = "Mayo"
    If MES = 6 Then nombredelmes = "Junio"
    If MES = 7 Then nombredelmes = "Julio"
    If MES = 8 Then nombredelmes = "Agosto"
    If MES = 9 Then nombredelmes = "Septiembre"
    If MES = 10 Then nombredelmes = "Octubre"
    If MES = 11 Then nombredelmes = "Noviembre"
    If MES = 12 Then nombredelmes = "Diciembre"
    
    If MES < 1 Then
        Beep
        MsgBox "El mes no puede ser menor de 1  !", vbExclamation, TIT_MSGBOX
        GoTo Error
    End If
    
    If MES > 12 Then
        Beep
        MsgBox "El mes no puede ser mayor de 12  !", vbExclamation, TIT_MSGBOX
        GoTo Error
    End If
    
    If ((MES = 1) Or (MES = 1) Or (MES = 3) Or (MES = 5) Or (MES = 6) Or (MES = 8) Or (MES = 10) Or (MES = 12)) And DIA > 31 Then
        Beep
        MsgBox Trim(nombredelmes) & " tiene s�lo 31 d�as !", vbExclamation, TIT_MSGBOX
        GoTo Error
    End If
    If MES = 2 And DIA > 28 Then
        Beep
        MsgBox "Febrero tiene s�lo 28 d�as !", vbExclamation, TIT_MSGBOX
        GoTo Error
    End If
    If ((MES = 4) Or (MES = 7) Or (MES = 9) Or (MES = 11)) And DIA > 30 Then
        Beep
        MsgBox Trim(nombredelmes) & " tiene s�lo 30 d�as !", vbExclamation, TIT_MSGBOX
        GoTo Error
    End If
    Exit Function
    
Error:
    VerificarFecha = False
End Function

Function MonedaSQL(ByVal TEXTO As String) As String

    Dim Caracter As String
    Dim AuxTexto As String
    Dim SinMiles As String
    Dim Resultado As String
    
    AuxTexto = Trim(TEXTO)
    
    ' si es un blanco
    If AuxTexto = "" Then
        MonedaSQL = "0"
        Exit Function
    End If
    
    'quito el s�mbolo de moneda
    If Not IsNumeric(Left$(AuxTexto, 1)) Then
        AuxTexto = Trim(Right$(AuxTexto, Len(AuxTexto) - 1))
    End If
    
    'elimino los puntos de miles
    SinMiles = ""
    For I = 1 To Len(AuxTexto)
        If Mid(AuxTexto, I, 1) <> vPuntoMiles Then
            SinMiles = SinMiles & Mid(AuxTexto, I, 1)
        End If
    Next I
        
    'reemplazo el punto decimal por punto (.)
    Resultado = ""
    For I = 1 To Len(SinMiles)
        If Mid(SinMiles, I, 1) = vPuntoDecimal Then
            Resultado = Resultado & "."
        Else
            Resultado = Resultado & Mid(SinMiles, I, 1)
        End If
    Next I
        
    'retorno el string resulato
    MonedaSQL = Resultado
    
End Function

Public Sub Buscar_Concepto(tabla As String, Campo As String, TEXTO As Object, combo As Object)

    Set rec = New ADODB.Recordset

    If Trim(TEXTO) = "" Then Exit Sub
    cSQL = "SELECT " & Trim(Campo) & " as concepto " & _
            " FROM " & Trim(tabla) & " WHERE " & _
        "substring(" & Trim(Campo) & ",1," & Len(TEXTO) & ") = '" & Trim(TEXTO) & "'"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    'si no encontro no permite continuar
    If rec.RecordCount = 0 Then
        Beep
        MsgBox "Concepto Inexistente !", vbExclamation, TIT_MSGBOX
        TEXTO = ""
        If TEXTO.Enabled Then TEXTO.SetFocus
        Exit Sub
    End If
    'si encontro uno solo lo pone en el texto
    TEXTO = Trim(rec!Concepto)
End Sub

Public Sub Descripcion_Combo(CboGral As Object, Descripcion As String)
                    
'******************As� se invoca a la funcion************************
'    Dim xSql As String                                             *
'    Dim snpDescri As ADODB.Recordset                               *
'                                                                   *
'     xSql = "Select pro_descri from provincia " & _                *
'            "where  pro_codigo =  " & snpProv!t_pro_codigo         *
'     If DBConn.GetRecordset(snpDescri, xSql) Then                  *
'        If snpDescri.EOF = False Then                              *
'            Descripcion_Combo CboProvincia, snpDescri!pro_descri   *
'        End If                                                     *
'        snpDescri.Close                                            *
'     End If                                                        *
'********************************************************************
      
      CboGral.ListIndex = 0
      Do While Trim(Descripcion) <> Trim(CboGral.List(CboGral.ListIndex))
        CboGral.ListIndex = CboGral.ListIndex + 1
      Loop
      CboGral.SetFocus
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
    
        sql = "SELECT * FROM PERMISOS WHERE USU_NOMBRE = '" & Trim(USUARIO) & "'"
        r.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If r.EOF = False Then
          'r.MoveFirst
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

Public Function ChkNull(Valor) As String
    If IsNull(Valor) Then
        ChkNull = ""
    Else
        ChkNull = Valor
    End If
End Function

Public Function ConvDate(Valor As Variant)
    If Valor = "__/__/____" Then
       ConvDate = "Null"
    Else
       ConvDate = "'" + Format(Valor, "mm/dd/yyyy") + "'"
    End If

End Function

Function GetFechaConBarras(Fecha As String)
  ' busca el DIA -----------------------
  POS1 = InStr(Fecha, "/")
  If POS1 = 0 Then
     GetFechaConBarras = ""
     Exit Function
  End If
  DIA = Left(Fecha, POS1 - 1)
  Select Case Len(DIA)
  Case Is < 1
    GetFechaConBarras = ""
    Exit Function
  Case 1
    DIA = "0" & DIA
  Case Is > 2
    GetFechaConBarras = ""
    Exit Function
  End Select
  
  ' busca el MES -----------------------
  POS2 = InStr(POS1 + 1, Fecha, "/")
  If POS2 = 0 Then
     GetFechaConBarras = ""
     Exit Function
  End If
  MES = Mid(Fecha, POS1 + 1, POS2 - POS1 - 1)
  Select Case Len(MES)
  Case Is < 1
    GetFechaConBarras = ""
    Exit Function
  Case 1
    MES = "0" & MES
  Case Is > 2
    GetFechaConBarras = ""
    Exit Function
  End Select
  
  ' busca el A�O -----------------------
  Anio = Mid(Fecha, POS2 + 1, 4)
  Select Case Len(Anio)
  Case Is < 1
    GetFechaConBarras = ""
    Exit Function
  Case 3
    GetFechaConBarras = ""
    Exit Function
  Case Is > 4
    GetFechaConBarras = ""
    Exit Function
  Case 1
    Anio = Left(Trim(Str(Year(Now))), 2) & "0" & Anio
  Case 2
    Anio = Left(Trim(Str(Year(Now))), 2) & Anio
  End Select
 
  GetFechaConBarras = DIA & MES & Anio
  
End Function

Function GetFechaSinBarras(Fecha As String)
  Select Case Len(Fecha)
  Case Is < 6
    GetFechaSinBarras = ""
  Case 6
    GetFechaSinBarras = Left(Fecha, 4) & Left(Year(Now), 2) & Right(Fecha, 2)
  Case 7
    GetFechaSinBarras = ""
  Case 8
    GetFechaSinBarras = Fecha
  Case Is > 8
    GetFechaSinBarras = ""
  End Select
End Function

Function ValidarIngresoFecha(Fecha As String) As String
    
 Dim FechaAValid As String

  POS1 = InStr(Fecha, "/")
  If POS1 <> 0 Then
    Fecha = GetFechaConBarras(Fecha)
  Else
    Fecha = GetFechaSinBarras(Fecha)
  End If
  If Fecha = "" Then
    ValidarIngresoFecha = ""
    Exit Function
  End If
  
'Controla el dia ----------------
  DIA = Left(Fecha, 2)
  If Not IsNumeric(DIA) Then
    ValidarIngresoFecha = ""
    Exit Function
  End If
  If DIA < "0" Or DIA > "31" Then
    ValidarIngresoFecha = ""
    Exit Function
  End If
  
'Controla el mes ----------------
  MES = Mid(Fecha, 3, 2)
  If Not IsNumeric(MES) Then
    ValidarIngresoFecha = ""
    Exit Function
  End If
  If MES < "0" Or MES > "12" Then
    ValidarIngresoFecha = ""
    Exit Function
  End If
  If MES = "2" And DIA > "28" Then 'Febrero (falta control para bisiesto)
    ValidarIngresoFecha = ""
    Exit Function
  End If
  If (MES = "4" Or MES = "6" Or MES = "9" Or MES = "11") And DIA > "30" Then
    ValidarIngresoFecha = ""
    Exit Function
  End If
  
  Anio = Right(Fecha, 4)
  FechaAValid = DIA & "/" & MES & "/" & Anio
  
  ValidarIngresoFecha = IIf(IsDate(FechaAValid), FechaAValid, "")

End Function

Public Function CarNumeroEntero(ByRef KeyAscii As Integer) As Integer
    Dim Car As String * 1
    Car = Chr$(KeyAscii)
    If (Car < "0" Or Car > "9") And KeyAscii <> 8 Then
        CarNumeroEntero = 0
    Else
        CarNumeroEntero = KeyAscii
    End If
End Function

Public Function CarNumeroTE(ByRef KeyAscii As Integer) As Integer
    Dim Car As String * 1
    Car = Chr$(KeyAscii)
    If (Car < "0" Or Car > "9") And KeyAscii <> 8 And Car <> "-" And Car <> "(" And Car <> ")" And Car <> " " Then
        CarNumeroTE = 0
    Else
        CarNumeroTE = KeyAscii
    End If
End Function

Public Function Carga_Fecha(Ctrl As MaskEdBox)
    Valor = Ctrl.ClipText
    Fec = (Ctrl.Text)
    num = Len(Valor)
    If num = 6 Then

        If Right(Valor, 2) > 50 Then
            a�o = Left(Year(Date), 2) + Right(Valor, 2)
        Else
            a�o = Left(Year(Date) + 1, 2) + Right(Valor, 2)
        End If

        Carga_Fecha = Left(Valor, 2) + "/" + Mid(Valor, 3, 2) + "/" + a�o

    ElseIf num = 4 Then
        Carga_Fecha = Left(Valor, 2) + "/" + Mid(Valor, 3, 2) + "/" + CStr(Year(Date))
    Else
        Carga_Fecha = Fec
    End If
    If IsDate(Format(Carga_Fecha, "mm/dd/yyyy")) = False And Fec <> "__/__/____" Then
        Carga_Fecha = "__/__/____"
        Ctrl.SetFocus
        MsgBox "ERROR EN FECHA", 48, TIT_MSGBOX
    End If
End Function
Public Function CarFecha(ByRef KeyAscii As Integer) As Integer
    Dim Car As String * 1
    Car = Chr$(KeyAscii)
    If (Car < "0" Or Car > "9") And KeyAscii <> 8 And Car <> "/" Then
        Beep
        CarFecha = 0
    Else
        CarFecha = KeyAscii
    End If
End Function

Public Function Centrar_pantalla(ventana As Form)
    ventana.Left = (Screen.Width - ventana.Width) / 2
    ventana.Top = (Screen.Height - ventana.Height) / 2
End Function


Sub seltxt(Optional txt As Variant)
    'If Not IsMissing(Txt) Then
        If TypeOf Screen.ActiveControl Is TextBox Then
            Dim TxtCtrol As TextBox
            Set TxtCtrol = Screen.ActiveControl
            TxtCtrol.SelStart = 0
            TxtCtrol.SelLength = Len(TxtCtrol.Text)
            Exit Sub
        End If
    'End If
End Sub

Public Function DateNull(Valor As Variant) As String
    If IsNull(Valor) Then
        DateNull = "__/__/____"
    Else
        'DateNull = MaskFecha_TXT(Valor)    Javier le coloco asterisco
        DateNull = Valor
    End If
End Function

Public Function XD(Valor As Variant, Optional ddmmyyyy As Boolean, Optional Formato As Boolean)
    
    If IsNull(Formato) Then Formato = False
    
    If Valor = "__/__/____" Or Valor = "" Then
        XD = "Null"
    Else
        If ddmmyyyy = False Then
            XD = "'" + Format(Valor, "mm/dd/yyyy ttttt") + "'"
        Else
            If Formato = False Then
                XD = "'" + Format(Valor, "dd/mm/yyyy ttttt") + "'"
            Else
                XD = "'" + Format(Valor, "mm/dd/yyyy ttttt") + "'"
            End If
        End If
    End If
End Function

Public Function XDQ(Valor As Variant)
    If Valor = "" Or IsNull(Valor) Then
        XDQ = "Null"
    Else
        XDQ = "#" & Format(Valor, "mm/dd/yyyy") & "#"
    End If
End Function

Public Function XD1(Valor As Variant, Optional ddmmyyyy As Boolean, Optional Formato As Boolean)
    
    If IsNull(Formato) Then Formato = False
    
    If Valor = "__/__/____" Or Valor = "" Then
        XD1 = "Null"
    Else
        If ddmmyyyy = False Then
            XD1 = "'" + Format(Valor, "mm/dd/yyyy") + "'"
        Else
            If Formato = False Then
                XD1 = "'" + Format(Valor, "dd/mm/yyyy") + "'"
            Else
                XD1 = "'" + Format(Valor, "mm/dd/yyyy") + "'"
            End If
        End If
    End If
End Function

Public Function XS(Valor As Variant, Optional minuscula As Boolean)
    If Valor = "" Then
        XS = "Null"
    Else
        If Not IsNull(minuscula) And minuscula = True Then
            XS = "'" + Trim(Valor) + "'"
        Else
            XS = "'" + UCase(Trim(Valor)) + "'"
        End If
    End If
End Function

Public Function XN(Valor As String, Optional decimales As Integer) As String
    Dim a As Integer, TEXTO As String, cad1 As String

    If Trim(Valor) = "" Or Trim(Valor) = "Null" Then
        XN = "Null"
    Else
        If decimales > 0 Then
            cad1 = "0."
            For a = 1 To decimales
                cad1 = cad1 & "0"
            Next
            XN = Format(Valor, cad1)
        Else
            XN = Trim(Valor)
        End If
        For a = 1 To Len(XN)
            If Mid(XN, a, 1) = "," Then
                TEXTO = TEXTO & "."
            ElseIf Mid(XN, a, 1) = "." Then
                TEXTO = TEXTO
            Else
                TEXTO = TEXTO & Mid(XN, a, 1)
            End If
        Next
        XN = TEXTO
    End If
End Function

Public Function XN1(Valor As Variant)
    Valor = Valor & ""
    If Trim(Valor) = "" Then
       XN1 = "Null"
    Else
       XN1 = Trim(UCase(Str(Valor)))
    End If
End Function

Public Sub Mensaje(Numero As Integer)
'Este procedimiento muestra un mensaje estandard segun el
'Numero que reciba:
'    1 - INSERT     4 - SELECT
'    2 - DELETE     5 -
'    3 - UPDATE     6 - IMPRESION
'    etc.
'Para agregar nuevos mensajes solo agregue el numero al final
'Autor: La Renga Borri
'(Ol raights reserbed).

    If Numero = 1 Then     ' INSERT
        MsgBox "Ha ocurrido un error al tratar de ingresar el registro !" & Chr(13) & Chr(13) & _
        Err.Description, vbCritical, "Error:"
    ElseIf Numero = 2 Then ' DELETE
        MsgBox "Ha ocurrido un error al tratar de eliminar el registro !" & Chr(13) & Chr(13) & _
        Err.Description, vbCritical, "Error:"
    ElseIf Numero = 3 Then ' UPDATE
        MsgBox "Ha ocurrido un error al tratar de actualizar el registro !" & Chr(13) & Chr(13) & _
        Err.Description, vbCritical, "Error:"
    ElseIf Numero = 4 Then ' SELECT
        MsgBox "Ha ocurrido un error al tratar de leer el registro !" & Chr(13) & _
        Err.Number & "  " & Err.Description, vbCritical, "Error:"
    ElseIf Numero = 5 Then 'Control de borrado y modificaci�n
        MsgBox " El c�digo seleccionado ya ha sido utilizado, transacci�n abortada. "
    ElseIf Numero = 6 Then 'IMPRESION
        MsgBox "Ha ocurrido un error al generar la impresion !" & Chr(13) & _
        Err.Number & "  " & Err.Description, vbCritical, "Error:"
    Else
        MsgBox "Ha ocurrido el siguiente error:" & Chr(13) & Chr(13) & _
        Err.Description, vbCritical, "Error:"
    End If
    
End Sub

Public Sub BorraFilaEnBlanco(G As MSFlexGrid, COLUMNA As Integer)
    'Recorre la grilla y elimina las filas en blanco
    'Una fila se considera en blanco cuando "Columna" est� en blanco
    Dim CantFila As Integer
    Dim FilaAct As Integer
    Dim j As Integer
    CantFila = G.Rows
    G.Col = COLUMNA
    If CantFila > 2 Then
        FilaAct = 1
        Do While FilaAct < CantFila And CantFila > 2
           G.row = FilaAct
           If G.Text = "" Then
              G.RemoveItem (FilaAct)
              CantFila = CantFila - 1
           Else
              FilaAct = FilaAct + 1
           End If
        Loop
    Else
       FilaAct = 1
       If G.Text = "" Then
          j = 0
          Do While j < G.Cols
             G.Col = j
             G.Text = ""
             j = j + 1
          Loop
       End If
    End If
    G.row = 1
End Sub

Public Function SELECT_DEUDA(TipoDoc As String, Nrodoc As String, Fecha As String, _
        completa As Boolean, Optional Sistema As String, Optional Subsistema As String) As String
        Dim Controlo As Boolean
        
        If Sistema <> "" And Subsistema <> "" Then
           Controlo = True
        Else
           Controlo = False
        End If
        
        cSQL = " Select Cta_Cte_SubSist.Cta_NumCed,"
        cSQL = cSQL & " Cta_Cte_SubSist.Sis_Codigo,Sistema.Sis_Descri,"
        cSQL = cSQL & " Cta_Cte_SubSist.Sub_Codigo,SubSistema.Sub_Descri,"
        cSQL = cSQL & " Cta_Cte_SubSist.Ccs_ValNom,V1.Ven_FecVto,"
        cSQL = cSQL & " Cta_Cte_SubSist.Rub_Codigo,Rubros.Rub_Descri,"
        cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,3,4),"
        cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,1,2)"
        cSQL = cSQL & " From Cta_Cte,Cta_Cte_SubSist,Periodo,Vencimiento V1,Sistema,SubSistema,Rubros"
        cSQL = cSQL & " Where Cta_Cte.Tip_TipDoc = " & XS(Trim(TipoDoc)) & _
                        " And Cta_Cte.Per_NroDoc=" & XS(Trim(Nrodoc))
        cSQL = cSQL & " And Cta_Cte.Cta_NumCed = Cta_Cte_SubSist.Cta_NumCed "
        
'        cSQL = cSQL + " And SUBSTRING(Cta_Cte_SubSist.CTA_NUMCED,1,6) NOT IN('122001','012002','022002','032002','042002','052002')"
        
        cSQL = cSQL & " And (Cta_Cte_SubSist.Ccs_Pagado = 0 OR (Cta_Cte_SubSist.Ccs_Pagado = 1 AND Cta_Cte_SubSist.CCS_FECPAG > " & XD(Fecha) & " ))  "
        cSQL = cSQL & " And Cta_Cte.Pdo_Period=Periodo.Pdo_Period "
        cSQL = cSQL & " And Periodo.Pdo_Period=V1.Pdo_Period "
        cSQL = cSQL & " And V1.Ven_FecVto = (Select Min(V2.Ven_FecVto) From Vencimiento V2 "
        cSQL = cSQL & " Where V2.Pdo_Period = Cta_Cte.Pdo_Period "
        cSQL = cSQL & " And V2.Ven_FecVto >= " & XD(Fecha) & ")"
        cSQL = cSQL & " And Rubros.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And Sistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And Sistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
        cSQL = cSQL & " And Subsistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And SubSistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
        cSQL = cSQL & " And SubSistema.Sub_Codigo=Cta_Cte_SubSist.Sub_Codigo"
        cSQL = cSQL & " And " & XD(Fecha) & " > "
        cSQL = cSQL & " (Select Min(V2.Ven_FecVto)  From Vencimiento V2 "
        cSQL = cSQL & " Where V2.Pdo_Period = Cta_Cte.Pdo_Period)"
        
        If Controlo = True Then
            cSQL = cSQL & " And SubSistema.Sis_Codigo = " & XN(Sistema)
            cSQL = cSQL & " And SubSistema.Sub_Codigo = " & XN(Subsistema)
        End If
        
        cSQL = cSQL & " Union All "
        
        cSQL = cSQL & " Select Cta_Cte_SubSist.Cta_NumCed,"
        cSQL = cSQL & " Cta_Cte_SubSist.Sis_Codigo,Sistema.Sis_Descri,"
        cSQL = cSQL & " Cta_Cte_SubSist.Sub_Codigo,SubSistema.Sub_Descri,"
        cSQL = cSQL & " Cta_Cte_SubSist.Ccs_ValNom, " & XD(Fecha) & ","
        cSQL = cSQL & " Cta_Cte_SubSist.Rub_Codigo,Rubros.Rub_Descri,"
        cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,3,4),"
        cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,1,2)"
        cSQL = cSQL & " From Cta_Cte,Cta_Cte_SubSist,Periodo,Vencimiento V1,Sistema,SubSistema,Rubros"
        cSQL = cSQL & " Where Cta_Cte.Tip_TipDoc=" & XS(Trim(TipoDoc)) & " And Cta_Cte.Per_NroDoc=" & XS(Trim(Nrodoc))
        cSQL = cSQL & " And Cta_Cte.Cta_NumCed=Cta_Cte_SubSist.Cta_NumCed "
        
 '       cSQL = cSQL + " And SUBSTRING(Cta_Cte_SubSist.CTA_NUMCED,1,6) NOT IN('122001','012002','022002','032002','042002','052002')"
        
        cSQL = cSQL & " And (Cta_Cte_SubSist.Ccs_Pagado = 0 OR (Cta_Cte_SubSist.Ccs_Pagado = 1 AND Cta_Cte_SubSist.CCS_FECPAG > " & XD(Fecha) & " ))  "
        cSQL = cSQL & " And Cta_Cte.Pdo_Period=Periodo.Pdo_Period "
        cSQL = cSQL & " And Rubros.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And Sistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And Sistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
        cSQL = cSQL & " And SubSistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And SubSistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
        cSQL = cSQL & " And SubSistema.Sub_Codigo=Cta_Cte_SubSist.Sub_Codigo "
        cSQL = cSQL & " And Cta_Cte.Pdo_Period=Periodo.Pdo_Period "
        cSQL = cSQL & " And Periodo.Pdo_Period=V1.Pdo_Period "
        cSQL = cSQL & " And V1.Ven_FecVto="
        cSQL = cSQL & " (Select Max(V2.Ven_FecVto) From Vencimiento V2 "
        cSQL = cSQL & " Where V2.Pdo_Period = Cta_Cte.Pdo_Period "
        cSQL = cSQL & " And V2.Ven_FecVto < " & XD(Fecha) & " ) "
        cSQL = cSQL & " And " & XD(Fecha) & " > (Select Max(V2.Ven_FecVto) "
        cSQL = cSQL & " From Vencimiento V2 Where V2.Pdo_Period = Cta_Cte.Pdo_Period)"
        
        If Controlo = True Then
            cSQL = cSQL & " And SubSistema.Sis_Codigo = " & XN(Sistema)
            cSQL = cSQL & " And SubSistema.Sub_Codigo = " & XN(Subsistema)
        End If
        
        If Not completa Then
           cSQL = cSQL & " Order By substring(Cta_Cte_SubSist.Cta_NumCed,3,4),"
           cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,1,2),"
           cSQL = cSQL & " Cta_Cte_SubSist.Rub_Codigo,"
           cSQL = cSQL & " Cta_Cte_SubSist.Sis_Codigo,"
           cSQL = cSQL & " Cta_Cte_SubSist.Sub_Codigo "
        Else
            cSQL = cSQL & " Union All "
            cSQL = cSQL & " Select Cta_Cte_SubSist.Cta_NumCed,"
            cSQL = cSQL & " Cta_Cte_SubSist.Sis_Codigo,Sistema.Sis_Descri,"
            cSQL = cSQL & " Cta_Cte_SubSist.Sub_Codigo,SubSistema.Sub_Descri,"
            cSQL = cSQL & " Cta_Cte_SubSist.Ccs_ValNom, V1.Ven_FecVto,"
            cSQL = cSQL & " Cta_Cte_SubSist.Rub_Codigo,Rubros.Rub_Descri,"
            cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,3,4),"
            cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,1,2)"
            cSQL = cSQL & " From Cta_Cte,Cta_Cte_SubSist,Periodo,Vencimiento V1,Sistema,SubSistema,Rubros"
            cSQL = cSQL & " Where Cta_Cte.Tip_TipDoc=" & XS(Trim(TipoDoc)) & " And Cta_Cte.Per_NroDoc=" & XS(Trim(Nrodoc)) & ""
            cSQL = cSQL & " And Cta_Cte.Cta_NumCed=Cta_Cte_SubSist.Cta_NumCed "
            
  '          cSQL = cSQL + " And SUBSTRING(Cta_Cte_SubSist.CTA_NUMCED,1,6) NOT IN('122001','012002','022002','032002','042002','052002')"
            
            cSQL = cSQL & " And (Cta_Cte_SubSist.Ccs_Pagado = 0 OR (Cta_Cte_SubSist.Ccs_Pagado = 1 AND Cta_Cte_SubSist.CCS_FECPAG > " & XD(Fecha) & " ))  "
            cSQL = cSQL & " And Cta_Cte.Pdo_Period=Periodo.Pdo_Period "
            cSQL = cSQL & " And Periodo.Pdo_Period=V1.Pdo_Period "
            cSQL = cSQL & " And V1.Ven_FecVto=(Select Min(V2.Ven_FecVto) From Vencimiento V2 "
            cSQL = cSQL & " Where V2.Pdo_Period=Cta_Cte.Pdo_Period) "
            cSQL = cSQL & " And V1.Ven_FecVto >=" & XD(Fecha) & ""
            cSQL = cSQL & " And Rubros.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
            cSQL = cSQL & " And Sistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
            cSQL = cSQL & " And Sistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
            cSQL = cSQL & " And SubSistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
            cSQL = cSQL & " And SubSistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
            cSQL = cSQL & " And SubSistema.Sub_Codigo=Cta_Cte_SubSist.Sub_Codigo"
            
            If Controlo = True Then
                cSQL = cSQL & " And SubSistema.Sis_Codigo = " & XN(Sistema)
                cSQL = cSQL & " And SubSistema.Sub_Codigo = " & XN(Subsistema)
            End If
            
            cSQL = cSQL & " Order By substring(Cta_Cte_SubSist.Cta_NumCed,3,4),"
            cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,1,2),"
            cSQL = cSQL & " Cta_Cte_SubSist.Rub_Codigo,"
            cSQL = cSQL & " Cta_Cte_SubSist.Sis_Codigo, "
            cSQL = cSQL & " Cta_Cte_SubSist.Sub_Codigo "
        End If
        
        SELECT_DEUDA = cSQL
End Function


Public Function GraboImpresiondeRecibo(NroRec As Long, Delegacion As String, Concepto As String, Optional Importe As String)

     Item = Item + 1
    
    'Grabo Impresi�n del Recibo
     cSQL = " Insert Into IMPRESION_RECIBO (REC_NUMERO,DEL_CODIGO, REC_ITEM,REC_DETALLE) " & _
            " Values ( " & NroRec & " ," & XN(Delegacion) & ", " & Item & _
            "," & XS(CompletarConEspacios(Left(Trim(Concepto), 27), 27) & Space(1) & _
                     CompletarConEspaciosIzq(ChkNull(Importe), 12) & Space(10) & _
                     CompletarConEspacios(Trim(Concepto), 68) & Space(1) & _
                     CompletarConEspaciosIzq(ChkNull(Importe), 12)) & ")"
    DBConn.Execute cSQL
     
End Function
Sub BuscaProx(Codigo As String, combo As Object)
    'Realiza una b�squeda aproximada en un combo
    'Autor: 'La Renga' Borri
    '(Ol raights reserved).
    combo.ListIndex = -1
    Do While combo.ListIndex < combo.ListCount - 1
        If Trim(Mid(Trim(combo.Text), 1, Len(Codigo))) = Trim(Codigo) Then Exit Do
        combo.ListIndex = combo.ListIndex + 1
    Loop
End Sub

Sub BuscaCodigoProx(Codigo As String, combo As Object)
    'Realiza una b�squeda aproximada en un combo por codigo
    'Autor: Trillado
    combo.ListIndex = -1
    Do While combo.ListIndex < combo.ListCount - 1
        If Trim(Mid(Trim(combo.Text), 100, Len(Codigo) + 2)) = Trim(Codigo) Then Exit Do
        combo.ListIndex = combo.ListIndex + 1
    Loop
End Sub

Sub BuscaCodigoProxItemData(Codigo As Integer, combo As Object)
    'Realiza una b�squeda aproximada en un combo por ItemData
    'Autor: Eliana & Trillado
    combo.ListIndex = 0
    Do While combo.ListIndex < combo.ListCount
        If combo.ItemData(combo.ListIndex) = Codigo Then Exit Do
        combo.ListIndex = combo.ListIndex + 1
    Loop
End Sub


Public Function Relleno_con_z(TEXTO As String)
    Largo = Len(TEXTO)
    For I = Largo To 30
        TEXTO = TEXTO & "Z"
    Next
    Relleno_con_z = XS(TEXTO)
End Function

Public Function ValidoCuit(cuit As String) As Boolean
    If Len(cuit) = 11 Then
       Dim I As Integer
       Dim VectorCuit(10) As Integer
       Dim VectorFactor As Variant
       VectorFactor = Array(5, 4, 3, 2, 7, 6, 5, 4, 3, 2, 0)
       Dim Sumatoria As Integer
       Dim Resto As Integer
       Dim DigVerif As Integer
       Sumatoria = 0
       I = 0
       Do While I < 11
          VectorCuit(I) = Val(Mid(cuit, I + 1, 1))
          I = I + 1
       Loop
       I = 0
       Do While I < 10
          Sumatoria = Sumatoria + (VectorCuit(I) * Val(VectorFactor(I)))
          I = I + 1
       Loop
       Resto = Sumatoria Mod 11
       If Resto = 0 Then
          DigVerif = 0
       Else
          DigVerif = 11 - Resto
          If DigVerif = 10 Then
             DigVerif = 9
          End If
       End If
       If DigVerif <> VectorCuit(10) Then
          MsgBox "C.U.I.T. inv�lido, d�gito verificador correcto: " & DigVerif & " ", vbExclamation, TIT_MSGBOX
          ValidoCuit = False
       Else
          ValidoCuit = True
       End If
    Else
       Dim Desc As String
       If Len(cuit) > 11 Then
          Desc = " m�s. "
       Else
          Desc = " menos. "
       End If
       MsgBox " Ha ingresado " & Abs(Len(cuit) - 11) & " d�gito(s) de" & Desc & " ", vbExclamation, TIT_MSGBOX
       ValidoCuit = False
    End If
End Function

Public Function CargaComboPais(combo As ComboBox) As Boolean
    
    Dim cSqlTmp2 As String
    
    Set RecTmp2 = New ADODB.Recordset
    Screen.MousePointer = 11
    cSqlTmp2 = " Select PAI_CODIGO,PAI_DESCRI From PAIS Order By PAI_DESCRI"
    combo.Clear
    RecTmp2.Open cSqlTmp2, DBConn, adOpenStatic, adLockOptimistic
    If RecTmp2.EOF = False Then
       Do While Not RecTmp2.EOF
          combo.AddItem CompletarConEspacios(RecTmp2.Fields(1), 100) & RecTmp2.Fields(0)
          RecTmp2.MoveNext
       Loop
       CargaComboPais = True
    Else
       Screen.MousePointer = 1
       MsgBox " La tabla PAIS est� vac�a. ", 16, TIT_MSGBOX
       CargaComboPais = False
    End If
    RecTmp2.Close
    Screen.MousePointer = 1
End Function

Public Function CargaComboProvincia(combo As ComboBox, Pais As String) As Boolean
    Dim cSqlTmp1 As String
    Set RecTmp1 = New ADODB.Recordset
    Screen.MousePointer = 11
    cSqlTmp1 = " Select PRO_CODIGO,PRO_DESCRI From PROVINCIA Where PAI_CODIGO = " & Pais & " Order By PRO_DESCRI"
    combo.Clear
    RecTmp1.Open cSqlTmp1, DBConn, adOpenStatic, adLockOptimistic
    If RecTmp1.EOF = False Then
       'RecTmp1.MoveFirst
       Do While Not RecTmp1.EOF
          combo.AddItem CompletarConEspacios(RecTmp1.Fields(1), 100) & RecTmp1.Fields(0)
          RecTmp1.MoveNext
       Loop
       CargaComboProvincia = True
    Else
       Screen.MousePointer = 1
       MsgBox " La tabla PROVINCIA est� vac�a. ", 16, TIT_MSGBOX
       CargaComboProvincia = False
    End If
    RecTmp1.Close
    Screen.MousePointer = 1
End Function

Public Function CargaComboLocalidad(combo As ComboBox, Pais As String, Provincia As String) As Boolean
    Dim cSqlTmp1 As String
    Set RecTmp1 = New ADODB.Recordset
    Screen.MousePointer = 11
    cSqlTmp1 = " Select LOC_CODIGO,LOC_DESCRI From LOCALIDAD Where PAI_CODIGO=" & Pais & " And PRO_CODIGO=" & Provincia & " Order By LOC_DESCRI"
    combo.Clear
    RecTmp1.Open cSqlTmp1, DBConn, adOpenStatic, adLockOptimistic
    If RecTmp1.EOF = False Then
       'RecTmp1.MoveFirst
       Do While Not RecTmp1.EOF
          combo.AddItem CompletarConEspacios(RecTmp1.Fields(1), 100) & RecTmp1.Fields(0)
          RecTmp1.MoveNext
       Loop
       CargaComboLocalidad = True
    Else
       Screen.MousePointer = 1
       MsgBox " La tabla LOCALIDAD est� vac�a. ", 16, TIT_MSGBOX
       CargaComboLocalidad = False
    End If
    combo.Refresh
    RecTmp1.Close
    Screen.MousePointer = 1
End Function

Public Function CompletarConEspacios(TEXTO As String, Cnt As Integer) As String
    'Devuelve una cadena compuesta por el texto pasado por par�metro y le anexa
    'blancos a la derecha, hasta que la longitud del texto y los blancos sea
    'igual a cnt
    CompletarConEspacios = TEXTO & Space(Cnt - Len(TEXTO))
End Function

Public Function CargaComboBarrio(combo As ComboBox, Pais As String, Provincia As String, Localidad As String) As Boolean
    Dim cSqlTmp1 As String
    Set RecTmp1 = New ADODB.Recordset
    Screen.MousePointer = 11
    cSqlTmp1 = " Select BAR_CODIGO,BAR_DESCRI From DIRECCION_COMPLETA Where PAI_CODIGO=" & Pais & " And PRO_CODIGO=" & Provincia & " And LOC_CODIGO=" & Localidad & " Order By BAR_DESCRI"
    combo.Clear
    RecTmp1.Open cSqlTmp1, DBConn, adOpenStatic, adLockOptimistic
    If RecTmp1.EOF = False Then
       'RecTmp1.MoveFirst
       Do While Not RecTmp1.EOF
          combo.AddItem CompletarConEspacios(RecTmp1.Fields(1), 100) & RecTmp1.Fields(0)
          RecTmp1.MoveNext
       Loop
       CargaComboBarrio = True
    Else
       Screen.MousePointer = 1
       MsgBox " La tabla BARRIO est� vac�a. ", 16, TIT_MSGBOX
       CargaComboBarrio = False
    End If
    RecTmp1.Close
    Screen.MousePointer = 1
End Function

Public Function SacoGuiones(TEXTO As String) As String
    'Analiza la cadena TEXTO y extrae cualquier gui�n que encuentre en ella
    Dim NueTexto As String
    Dim I As Integer
    NueTexto = ""
    If TEXTO <> "" And Not IsNull(TEXTO) Then
       I = 1
       Do While I <= Len(TEXTO)
          If Mid(TEXTO, I, 1) <> "-" Then
             NueTexto = NueTexto & Mid(TEXTO, I, 1)
          End If
          I = I + 1
       Loop
    End If
    SacoGuiones = NueTexto
End Function

Public Function CargaComboDelegacion(combo As ComboBox) As Boolean
    Dim cSqlTmp2 As String
    Set RecTmp2 = New ADODB.Recordset
    Screen.MousePointer = 11
    cSqlTmp2 = " Select DEL_CODIGO,DEL_DESCRI From DELEGACION WHERE DEL_FACTURA = 0 Order By DEL_DESCRI"
    combo.Clear
    RecTmp2.Open cSqlTmp2, DBConn, adOpenStatic, adLockOptimistic
    If RecTmp2.EOF = False Then
       'RecTmp2.MoveFirst
       Do While Not RecTmp2.EOF
          combo.AddItem CompletarConEspacios(RecTmp2.Fields(1), 100) & RecTmp2.Fields(0)
          RecTmp2.MoveNext
       Loop
       CargaComboDelegacion = True
    Else
       Screen.MousePointer = 1
       MsgBox " La tabla DELEGACION est� vac�a. ", 16, TIT_MSGBOX
       CargaComboDelegacion = False
    End If
    RecTmp2.Close
    Screen.MousePointer = 1
End Function

Public Function CargaComboTipoPago(combo As ComboBox) As Boolean
    Screen.MousePointer = 11
    combo.Clear
    combo.AddItem CompletarConEspacios("Boleta", 100) & "B"
    combo.AddItem CompletarConEspacios("Caja", 100) & "C"
    combo.AddItem CompletarConEspacios("Caja-Boleta", 100) & "A"
    Screen.MousePointer = 1
    CargaComboTipoPago = True
End Function
Public Sub CargaComboSubTipoOrden(combo As ComboBox, TIPO As Integer)
    sql = "SELECT * FROM SUBTIPO_ORDEN WHERE TOR_CODIGO = " & TIPO
    Set RecCombo = New ADODB.Recordset
    RecCombo.Open sql, DBConn, adOpenStatic, adLockOptimistic
    combo.Clear
    If RecCombo.EOF = False Then
        Do While Not RecCombo.EOF
            combo.AddItem CompletarConEspacios(Trim(RecCombo!SOR_DESCRI), 100) & RecCombo!SOR_CODIGO
            RecCombo.MoveNext
        Loop
    Else
        combo.AddItem ("Sin Datos")
    End If
    'Combo.ListIndex = 0
    RecCombo.Close
    Set rec = Nothing
End Sub
Public Sub CargaComboTipoOrden(combo As ComboBox, TIPO As Integer)
    Set RecCombo = New ADODB.Recordset
    sql = "SELECT * from TIPO_ORDEN WHERE TOR_CODIGO = " & TIPO
    RecCombo.Open sql, DBConn, adOpenStatic, adLockOptimistic
    combo.Clear
    If RecCombo.EOF = False Then
        Do While Not RecCombo.EOF
            combo.AddItem CompletarConEspacios(Trim(RecCombo!TOR_DESCRI), 100) & RecCombo!tor_codigo
            RecCombo.MoveNext
        Loop
    Else
        combo.AddItem ("Sin Datos")
    End If
    'Combo.ListIndex = 0
    RecCombo.Close
    Set RecCombo = Nothing
End Sub


Public Function ElementoNoRepetido(TEXTO As String, G As MSFlexGrid, Col As Integer) As Boolean
       'Busca en la grilla "G", si en la columna "Col" el string "texto" se repite
       Dim FilaActual As Integer
       Dim I As Integer
       Dim Repetido As Boolean
       FilaActual = G.row
       I = 1
       Repetido = False
       G.Col = Col
       Do While I < G.Rows
          G.row = I
          If G.Text = TEXTO Then
             Repetido = True
             Exit Do
          End If
          I = I + 1
       Loop
       G.row = FilaActual
       ElementoNoRepetido = Not Repetido
End Function

Public Sub CambiarColor(Matriz, CntObjetos As Integer, Color As ColorConstants, Detalle As String)
    '***************************************************************************************
    'Cambia el backColor del conjunto de controles pasados como par�metros
    'Matriz: Matriz variant que contiene los objetos a los que se les debe cambiar el fondo
    'CntObjetos: Cantidad de elementos de la matriz
    'Color: Color que debe darse al fondo
    'SoloDisable: "T"->se cambia el color a todos los objetos de la matriz
    '             "D"->se cambia el color solo a los que tengan Enabled=False
    '             "E"->se cambia el color solo a aquellos que tengan Enabled=True
    'Ejemplo para crear la matriz , asignarle los valores y llamar a la funci�n
    '       Dim MtrObjetos As Variant
    '       MtrObjetos = Array(Me.TxtCantidadCopias, Me.TxtCieEje, Me.TxtControlOriginal)
    '       Call CambiarColor(MtrObjetos, 3, &H808080, "D")
    '****************************************************************************************
    Dim I As Integer
    I = 0
    Do While I < CntObjetos
       If Detalle = "T" Then
          Matriz(I).BackColor = Color
       Else
          If Detalle = "D" Then
             If Not Matriz(I).Enabled Then
                Matriz(I).BackColor = Color
             End If
          Else
             If Matriz(I).Enabled Then
                Matriz(I).BackColor = Color
             End If
          End If
       End If
       I = I + 1
    Loop
End Sub

Sub EDITAR(griya As Control, edt As Control, Ki As Integer, Optional Fecha As Boolean)
'ESTE PROC. MUESTRA EL TXTEDIT EN LAS GRILLAS FELXGRID PARA HACERLAS
'EDITABLES ACOMODANDO EL TAMA�O DEL TXT AL DE LA CELDA DE LA GRILLA

    If Not Fecha Then 'el ocx fecha no necesita esto
        Select Case Ki
            Case 0 To 32
                edt.Text = griya
                edt.SelStart = 1000
            Case Else
                edt.Text = Chr(Ki)
                edt.SelStart = 1
        End Select
    End If
    edt.Move griya.CellLeft + griya.Left, griya.CellTop + griya.Top, griya.CellWidth, griya.CellHeight
    
    'edt.Text = ""
    edt.Visible = True
    edt.SetFocus
        
End Sub

Sub EditKeyCode(griya As Control, edt As Object, KeyCode As Integer, Shift As Integer)
'ESTE PROC. CONTROLA LAS TECLAS PRESIONADAS EN EL
'TXTEDIT EN LAS GRILLAS FELXGRID

    Select Case KeyCode
        Case 27
            edt.Visible = False
            griya.SetFocus
        Case 13
            griya.SetFocus
        Case 38
            griya.SetFocus
            DoEvents
            If griya.row > griya.FixedRows Then
                griya.row = griya.row - 1
            End If
        Case 40
            griya.SetFocus
            DoEvents
            If griya.row < griya.Rows - 1 Then
                griya.row = griya.row + 1
            End If
    End Select
End Sub

Function fgi(r As Integer, C As Integer) As Integer
    fgi = C + fg2.Cols * r
End Function


Public Function XM(Valor As Variant)
    If Valor = "__/__/____" Or Valor = "" Then
        XM = "Null"
    Else
        XM = "'" + Format(Valor, "dd/mm/yyyy") + "'"
    End If
End Function

Public Function DMY(Valor As Variant)
    If Valor = "__/__/____" Or Valor = "" Then
        DMY = "Null"
    Else
        DMY = Format(Valor, "dd/mm/yyyy")
    End If
End Function

Public Function MDY(Valor As Variant)
    If Valor = "__/__/____" Or Valor = "" Then
        MDY = "Null"
    Else
        MDY = Format(Valor, "mm/dd/yyyy")
    End If
End Function

Public Function CantidadDeElementosEnGrilla(G As MSFlexGrid, Col As Integer) As Integer
    'Devuelve la cantidad de elementos v�lidos en una MsFlexGrid
    'el par�metro col especifica la columna que debe evaluarse para
    'que la fila sea contada como v�lida
    Dim FilaAct As Integer
    Dim I As Integer
    Dim Cnt As Integer
    FilaAct = G.row
    Cnt = 0
    G.Col = Col
    I = 1
    Do While I < G.Rows
       G.row = I
       If G.Text <> "" Then
          Cnt = Cnt + 1
       End If
       I = I + 1
    Loop
    G.row = FilaAct
    CantidadDeElementosEnGrilla = Cnt
End Function


Public Function CalculoInteresMatricula(Cedulon As String, Saldo As Double, AFecha As String) As String
       'Calcula los Recargos correspondientes por Mart�cula (no por pr�stamo) a la fecha AFECHA
       'Par�metros: Cedulon: Nro. de Cedul�n
       '            Saldo  : Saldo del cedul�n sin incluir intereses
       '            AFecha : Fecha a la cual calcular los intereses
       
       'LEER: Como existen 3 vencimientos cuyos intereses se calculan al momento
       '      de generar la deuda, esta funci�n supone que no habr� cambios en
       '      la tasa de inter�s hasta que pase el 3� vencimiento. Por esto los cambios
       '      en la tasa de inter�s ocurridos hasta el 3� vto, no se tienen en cuenta
       '      (en el caso de cedulones generados en ese per�odo), utilizando la tasa
       '      correspondiente al 1� vto. Para cedulones correspondientes a per�odos
       '      anteriores, s� se respetan las variaciones en las tasas.
       
       Dim FecIni As String 'Primer Vencimiento del cedul�n
       Dim FecMed As String 'Segundo vto del cedul�n
       Dim FecFin As String 'Ultimo vto del cedul�n
       Dim FecRef As String 'Fecha de referencia que se utiliza en el c�lculo
       Dim cSqlInt As String
       Dim RecInt As ADODB.Recordset
       
       Dim FInferior As String 'Limite inferior de fecha por mes
       Dim FSuperior As String 'Limite superior de fecha por mes
       
       Dim CanDias As String   'Cantidad de d�as entre fechas
       Dim AcumTasa As Double  'Acumula las tasas mensuales
       Dim Tasa As Double
       Dim PunteroRaton As Integer
       Dim PeriodoCed As String
       
       PunteroRaton = Screen.MousePointer
       Screen.MousePointer = 11
       FirsTime = True
       AFecha = DMY(AFecha)
       AcumTasa = 0
       
       If Mid(Cedulon, 3, 4) < 1991 Then
          PeriodoCed = "041991"
          If Mid(Cedulon, 3, 4) = 1991 And Left(Cedulon, 2) < 4 Then
             PeriodoCed = "041991"
          End If
       Else
          PeriodoCed = Left(Cedulon, 6)
       End If
       
       Set RecInt = New ADODB.Recordset
       'cSqlInt = "Select distinct Ven_FecVto " & _
                 "From Vencimiento,Cta_Cte " & _
                 "Where substring(Cta_NumCed,1,6) = " & XS(PeriodoCed) & _
                  " And Cta_Cte.Pdo_Period = Vencimiento.Pdo_Period " & _
                  " Order By Ven_FecVto "
                  
       'Select Modificado por Optimizaci�n
       cSqlInt = "Select Ven_FecVto " & _
                 "From Vencimiento " & _
                 "Where Pdo_Period = " & XS(PeriodoCed) & _
                  " Order By Ven_FecVto "
       RecInt.Open cSqlInt, DBConn, adOpenStatic, adLockOptimistic
       If RecInt.EOF = False Then
          RecInt.MoveFirst
          FecIni = DMY(RecInt.Fields(0))
          RecInt.MoveNext
          FecMed = DMY(RecInt.Fields(0))
          RecInt.MoveNext
          FecFin = DMY(RecInt.Fields(0))
       End If
       RecInt.Close
       
       If CDate(FecIni) >= CDate(AFecha) Then     'Se solicita la deuda antes del 1� vto.(no lleva inter�s)
          AcumTasa = 0
       ElseIf CDate(FecMed) >= CDate(AFecha) Then 'Se solicita la deuda entre el 1� y 2� vto
          FInferior = GetFecPrevia(FecIni)
          CanDias = DateDiff("d", FecIni, FecMed)
          Tasa = GetTasa(FInferior)
          AcumTasa = AcumTasa + (Tasa * CanDias)
       ElseIf CDate(FecFin) >= CDate(AFecha) Then 'Se solicita la deuda entre el 2� y 3� vto.
          FInferior = GetFecPrevia(FecIni)
          CanDias = DateDiff("d", FecIni, FecFin)
          Tasa = GetTasa(FInferior)
          AcumTasa = AcumTasa + (Tasa * CanDias)
       Else 'Se solicita la deuda mas all� del 3� vto.
          FecRef = FecIni
          'Se utiliza la tasa correspondiente al 1� vto. para el per�odo 1�vto.-3�vto.
          If FecRef <= FecFin Then
             FInferior = GetFecPrevia(FecIni)
             CanDias = DateDiff("d", FecIni, FecFin)
             Tasa = GetTasa(FInferior)
             AcumTasa = AcumTasa + (Tasa * CanDias)
          End If
          FecRef = DateAdd("d", 1, FecFin)
          FecRef = DMY(FecRef)
          Do While CDate(FecRef) < CDate(AFecha)
             FInferior = GetFecPrevia(FecRef)
             FSuperior = GetFecPosterior(FecRef)
             'If CDate(FSuperior) > CDate(AFecha) Then
             '   FSuperior = AFecha
             'End If
             Tasa = GetTasa(FInferior)
             If CDate(FInferior) = CDate(FSuperior) Then
                CanDias = 1
             Else
                If CDate(FSuperior) >= CDate(AFecha) Then
                   CanDias = DateDiff("d", FecRef, AFecha) + 1
                Else
                   CanDias = DateDiff("d", FecRef, FSuperior)
                End If
             End If
             AcumTasa = AcumTasa + (Tasa * CanDias)
             FecRef = FSuperior
          Loop
       End If
       CalculoInteresMatricula = Format(CDbl(Saldo) * CDbl(AcumTasa), "###,###,##0.00")
       Screen.MousePointer = PunteroRaton
End Function

Function GetFecPrevia(Fecha As String) As String
    Dim RecFP As ADODB.Recordset
    Set RecFP = New ADODB.Recordset
    Dim cSqlFP As String
    cSqlFP = "Select TAS_FECHA From Tasa Where TAS_FECHA < = " & XD(Fecha) & " Order by TAS_FECHA desc"
    RecFP.Open cSqlFP, DBConn, adOpenStatic, adLockOptimistic
    If RecFP.EOF = False Then
       'RecFP.MoveLast
       GetFecPrevia = DMY(RecFP.Fields(0))
    Else
       GetFecPrevia = "01/01/1900"
    End If
    RecFP.Close
End Function

Function GetFecPosterior(Fecha As String) As String
    Dim RecFP As ADODB.Recordset
    Set RecFP = New ADODB.Recordset
    Dim cSqlFP As String
    cSqlFP = "Select TAS_FECHA From Tasa " & _
             " Where TAS_FECHA > " & XD(Fecha) & _
             " Order By TAS_FECHA"
    RecFP.Open cSqlFP, DBConn, adOpenStatic, adLockOptimistic
    If RecFP.EOF = False Then
       GetFecPosterior = DMY(RecFP.Fields(0))
    Else
       GetFecPosterior = "31/12/9999"
    End If
    RecFP.Close
End Function

Function GetFecPreviaPr(Fecha As String, TipoPr As String) As String
    Dim RecPr As ADODB.Recordset
    Set RecPr = New ADODB.Recordset
    Dim cSqlPr As String
    cSqlPr = "Select HIS_FECACTTA  From HIST_TASAS Where TPR_CODTIPPR=" & XS(TipoPr)
    cSqlPr = cSqlPr & " And  HIS_FECACTTA<=" & XD(Fecha) & " Order by HIS_FECACTTA"
    RecPr.Open cSqlPr, DBConn, adOpenStatic, adLockOptimistic
    If RecPr.EOF = False Then
       RecPr.MoveLast
       GetFecPreviaPr = DMY(RecPr.Fields(0))
    Else
       GetFecPreviaPr = "01/01/1900"
    End If
    RecPr.Close
End Function
Function GetFecPosteriorPr(Fecha As String, TipoPr As String) As String
    Dim RecPr As ADODB.Recordset
    Set RecPr = New ADODB.Recordset
    Dim cSqlPr As String
    cSqlPr = "Select HIS_FECACTTA  From HIST_TASAS Where TPR_CODTIPPR=" & XS(TipoPr)
    cSqlPr = cSqlPr & " And  HIS_FECACTTA>" & XD(Fecha) & " Order by HIS_FECACTTA"
    RecPr.Open cSqlPr, DBConn, adOpenStatic, adLockOptimistic
    If RecPr.EOF = False Then
       RecPr.MoveLast
       GetFecPosteriorPr = DMY(RecPr.Fields(0))
    Else
       GetFecPosteriorPr = "31/12/9999"
    End If
    RecPr.Close
End Function

Public Function GetTasaPr(Fecha As String, TIPO As String, TTasa As String) As Double
    Dim RecTasa As ADODB.Recordset
    Dim cSqlTasa As String
    Dim PunteroRaton As Integer
    PunteroRaton = Screen.MousePointer
    Screen.MousePointer = 11
    Set RecTasa = New ADODB.Recordset
    Select Case TTasa
           Case "N"
                cSqlTasa = " Select HIS_TASPRENO,HIS_PORPUNIT From HIST_TASAS "
           Case "I"
                cSqlTasa = " Select HIS_TASPREIN,HIS_PORPUNIT From HIST_TASAS "
           Case "S"
                cSqlTasa = " Select HIS_TASPRESU,HIS_PORPUNIT From HIST_TASAS "
    End Select
    cSqlTasa = cSqlTasa & " Where TPR_CODTIPPR=" & XS(TIPO)
    cSqlTasa = cSqlTasa & "   And HIS_FECACTTA=" & XD(Fecha)
    RecTasa.Open cSqlTasa, DBConn, adOpenStatic, adLockOptimistic
    If RecTasa.EOF = False Then
       GetTasaPr = ((Val(RecTasa.Fields(1)) + 1) * RecTasa.Fields(0)) / 12 'Tasa mensual
    Else
       GetTasaPr = "0"
       Screen.MousePointer = 1
       MsgBox " La tasa de inter�s para el per�odo " & XM(F) & " no est� cargada. El monto de la deuda consultada, es incorrecto ", vbExclamation, TIT_MSGBOX
       Screen.MousePointer = 11
    End If
    RecTasa.Close
    Screen.MousePointer = PunteroRaton
End Function

Public Function GetTasa(F) As Double
    Dim RecTasa As ADODB.Recordset
    Dim cSqlTasa As String
    Dim PunteroRaton As Integer
    PunteroRaton = Screen.MousePointer
    Screen.MousePointer = 11
    Set RecTasa = New ADODB.Recordset
    cSqlTasa = " Select TAS_TASA From TASA Where TAS_FECHA = " & XD(F)
    RecTasa.Open cSqlTasa, DBConn, adOpenStatic, adLockOptimistic
    If RecTasa.EOF = False Then
       GetTasa = ((RecTasa.Fields(0) / 30) / 100)
    Else
       GetTasa = "0"
       Screen.MousePointer = 1
       MsgBox " La tasa de inter�s para el per�odo " & XM(F) & " no est� cargada. El monto de la deuda consultada, es incorrecto ", vbExclamation, TIT_MSGBOX
       Screen.MousePointer = 11
    End If
    RecTasa.Close
    Screen.MousePointer = PunteroRaton
End Function

Public Function RecuperoParametro(NroParametro As String, Fecha As String) As Double
    Dim RecRPar As ADODB.Recordset
    Dim SqlRPar As String
    Dim Valor As String
    Set RecRPar = New ADODB.Recordset
    SqlRPar = " Select PTO_CODIGO,Max(PFE_FECHA) As MaxFecha,PFE_VALOR From PARAMETRO_FECHA "
    SqlRPar = SqlRPar & " Where PFE_FECHA<=" & XD(Fecha)
    SqlRPar = SqlRPar & " Group By PTO_CODIGO,PFE_VALOR "
    SqlRPar = SqlRPar & " Having PTO_CODIGO=" & XN(NroParametro)
    SqlRPar = SqlRPar & " Order By MaxFecha "
    RecRPar.Open SqlRPar, DBConn, adOpenStatic, adLockOptimistic
    If RecRPar.EOF = False Then
       RecRPar.MoveLast
       Valor = RecRPar.Fields(2)
    Else
       Valor = ""
    End If
    RecRPar.Close
    If Valor = "" Then
        RecuperoParametro = 0
    Else
        RecuperoParametro = CDbl(Valor)
    End If
End Function
Public Function Chk0(Valor) As String
    If IsNull(Valor) Or Valor = "" Then
        Chk0 = "0"
    ElseIf Valor = "" Then
        Chk0 = "0"
    ElseIf Valor = Space(Len(Valor)) Then
        Chk0 = "0"
    Else
        Chk0 = Valor
    End If
End Function

Public Function Calculo_Edad(Fec_Nac As Date, Optional FechaActual As Date) As Single
Dim FecAct As Date
If FechaActual = "00:00:00" Then      ' By Cismond Ado 2.0
    FecAct = Date
Else
    FecAct = FechaActual
End If

'ESTA FUNCION SE CREO PORQUE EL DATEDIFF ES UNA BOSTA!
'MARCOS BORRI
'22/06/1999 - ol raights reserved

'calculo los a�os contando los meses
Calculo_Edad = Int(DateDiff("m", CVDate(Fec_Nac), FecAct) / 12)

'Este aqui el problema:
'Por ejemplo: yo tengo 24 a�os y cumplo los a�os el 10 de diciembre.
'pero si hoy es 1 de diciembre para el datediff yo ya tengo 25 a�os !
'Entonces tengo que comparar los dias a mano

'PD: Esto no lo hago contando los dias y luego dividiendo por 365
'porque no se como tratar los a�os bisiestos asi que por las dudas lo hago a mano

If Month(FecAct) = Month(Fec_Nac) And Day(FecAct) < Day(Fec_Nac) Then Calculo_Edad = Calculo_Edad - 1

'Si tiene menos de un a�o devuelvo un decimal que indica la edad en meses
If Calculo_Edad = 0 Then
    Calculo_Edad = DateDiff("m", CVDate(Fec_Nac), FecAct)
    If Len(Trim(Str(Calculo_Edad))) = 1 Then
        Calculo_Edad = Calculo_Edad / 10
    ElseIf Len(Trim(Str(Calculo_Edad))) = 2 Then
        Calculo_Edad = Calculo_Edad / 100
    End If
End If

End Function

Public Function CalculoMontoCPS(nEdad As Integer, Cedulon As String) As Double
    Dim RecCPS As ADODB.Recordset
    Set RecCPS = New ADODB.Recordset
    Dim cSQLCPS As String
    
    ' Saco el Monto a Cobrar
    cSQLCPS = " SELECT MON_MONTO,MON_PORCEN " & _
              " FROM MONTO " & _
              " WHERE RUB_CODIGO = 3 " & _
                " AND SIS_CODIGO = 1 " & _
                " AND SUB_CODIGO = 22" & _
                " AND(" & nEdad & " >= MON_EDADES  " & _
                " AND " & nEdad & " <= MON_EDAHAS )" & _
                " AND MON_FECHIS <= " & XD("01/" & Left(Cedulon, 2) & " / " & Mid(Cedulon, 3, 4)) & " order by MON_FECHIS"
    RecCPS.Open cSQLCPS, DBConn, adOpenStatic, adLockOptimistic
    If (RecCPS.BOF And RecCPS.EOF) = 0 Then
        RecCPS.MoveLast
        ' Calculo el Importe a Aportar
        CalculoMontoCPS = (RecCPS!MON_MONTO * RecCPS!MON_PORCEN) / 100
    Else
        CalculoMontoCPS = 0
    End If
    RecCPS.Close
End Function

Public Function Matricula_Activa(Matricula As String) As Boolean
Dim r As ADODB.Recordset
Set r = New ADODB.Recordset

    Matricula_Activa = True

    If Trim(Matricula) = "" Then Exit Function

    sql = "SELECT CAR_CODIGO FROM TITULO WHERE " & _
    "TTI_CODIGO = " & Mid(Matricula, 1, 2) & " AND " & _
    "TIT_MATRIC = '" & Mid(Matricula, 3, 5) & "' AND " & _
    "TIT_DIGVER = '" & Mid(Matricula, 8, 1) & "' AND " & _
    "(CAR_CODIGO = 1 OR CAR_CODIGO = 13 OR CAR_CODIGO = 14 OR CAR_CODIGO = 15 OR CAR_CODIGO = 5  OR CAR_CODIGO = 4 OR CAR_CODIGO = 7  OR CAR_CODIGO = 10 OR titulo.car_codigo = 11)"
    r.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If r.RecordCount = 0 Then
        Matricula_Activa = False
        MsgBox "La matricula de esta persona no esta activa por lo tanto no se pueden modificar datos !", vbExclamation, TIT_MSGBOX
    End If
    r.Close

Set r = Nothing
End Function

Public Function UltimoDiadelMes(Fecha As Date) As Date
    Dim Fec As Date
    Fec = DateAdd("m", 1, Fecha)
    Fec = "01/" & Month(Fec) & "/" & Year(Fec)
    UltimoDiadelMes = DateAdd("d", -1, Fec)
End Function

Public Function TextoPrevioAlGuion(Cadena As String) As String
       Dim Largo As Integer
       Dim I As Integer
       Dim NueCadena As String
       I = 1
       NueCadena = ""
       Largo = Len(Cadena)
       Do While I <= Largo
          If Mid(Cadena, I, 1) = "-" Then
             Exit Do
          End If
          NueCadena = NueCadena & Mid(Cadena, I, 1)
          I = I + 1
       Loop
       TextoPrevioAlGuion = NueCadena
End Function

Public Function TextoPostAlGuion(Cadena As String) As String
       Dim Largo As Integer
       Dim I As Integer
       Dim NueCadena As String
       Dim Bandera As Boolean
       I = 1
       Bandera = False
       NueCadena = ""
       Largo = Len(Cadena)
       Do While I <= Largo
         If Not Bandera Then
          If Mid(Cadena, I, 1) = "-" Then
             Bandera = True
          End If
         Else
          NueCadena = NueCadena & Mid(Cadena, I, 1)
         End If
          I = I + 1
       Loop
       TextoPostAlGuion = NueCadena
End Function

Public Sub BuscoDescDelegacion(DELCOD As String)
    Dim SqlDel As String
    Dim RecDel As ADODB.Recordset
    Set RecDel = New ADODB.Recordset
    SqlDel = " Select DEL_DESCRI  From DELEGACION  Where DEL_CODIGO = " & XN(DELCOD)
    RecDel.Open SqlDel, DBConn, adOpenStatic, adLockOptimistic
    If RecDel.EOF = False Then
       DELDESC = Trim(RecDel.Fields!DEL_DESCRI)
    Else
       DELDESC = ""
    End If
    RecDel.Close
End Sub

Public Sub CambiaColorAFilaDeGrilla(G As MSFlexGrid, Fila As Integer, Optional ColorFuente As ColorConstants, Optional ColorFondo As ColorConstants)
    Dim j As Integer
    Dim ActFila, ActColu As Integer
    ActFila = G.row
    ActColu = G.Col
    G.row = Fila
    For j = 0 To G.Cols - 1
        G.Col = j
        'If Not IsNull(ColorFuente) Then
        '    G.CellForeColor = ColorFuente
        'End If
        'If Not IsNull(ColorFondo) Then
        '    G.CellBackColor = ColorFondo
        'End If
        'If ColorFuente <> "" Then
            G.CellForeColor = ColorFuente
        'End If
        'If ColorFondo <> "" Then
            G.CellBackColor = ColorFondo
        'End If
    Next
    G.row = ActFila
    G.Col = ActColu
End Sub

Public Sub Calculo_Interes_Grilla(G As MSFlexGrid, Fila As Integer, ColorFuente As ColorConstants)
    
    Interes = 0
    Actualizacion = 0
    
    Dim nEdad          As Integer
        
    'Si el Color es ROJO
    If ColorFuente = 255 Then
        'Recargo por el C.P.S.
        If G.TextArray(GRIDINDEX(G, G.row, 5)) = 1 And _
           G.TextArray(GRIDINDEX(G, G.row, 6)) = 22 Then
             'Para calcular...
             'Sacamos la EDAD Actual y la ESCALA Vigente a la fecha de la deuda
             'A la diferencia entre lo Generado y el resultado de la Escala
             'Se graba como una Actualizaci�n
             'Se calcula el Interes sobre la Nueva Base.
    
             ' Calculo la Actualizacion del CPS
             Dim PeriodoCed As String
             If G.TextArray(GRIDINDEX(G, G.row, 1)) < 1991 Then
                PeriodoCed = "041991"
                If G.TextArray(GRIDINDEX(G, G.row, 1)) = 1991 And G.TextArray(GRIDINDEX(G, G.row, 0)) < 4 Then
                   PeriodoCed = "041991"
                End If
             Else
                PeriodoCed = G.TextArray(GRIDINDEX(G, G.row, 0)) & G.TextArray(GRIDINDEX(G, G.row, 1))
             End If
    
             Dim UltDia As Date
             UltDia = UltimoDiadelMes("01" & "/" & G.TextArray(GRIDINDEX(G, G.row, 0)) & "/" & G.TextArray(GRIDINDEX(G, G.row, 1)))
             nEdad = Calculo_Edad(CVDate(FrmReciboContribuciones.TxtPerFecNac.Text), CVDate(UltDia))
    
             Dim FecAct As String
             Dim nMonto As Double
             
             FecAct = Format(Date, "MM") & Format(Date, "YYYY")
             nMonto = CalculoMontoCPS(nEdad, FecAct)
    
             'nMonto es el Valor Actual y Rec2.Fields(5) es lo grabado en la tabla
             Actualizacion = nMonto - G.TextArray(GRIDINDEX(G, G.row, 8))
             Interes = CalculoInteresMatricula(G.TextArray(GRIDINDEX(G, G.row, 0)) & G.TextArray(GRIDINDEX(G, G.row, 1)) & G.TextArray(GRIDINDEX(G, G.row, 2)), nMonto, FrmReciboContribuciones.txtRec_Fecha.Text)
        Else
             Interes = CalculoInteresMatricula(G.TextArray(GRIDINDEX(G, G.row, 0)) & G.TextArray(GRIDINDEX(G, G.row, 1)) & G.TextArray(GRIDINDEX(G, G.row, 2)), G.TextArray(GRIDINDEX(G, G.row, 8)), FrmReciboContribuciones.txtRec_Fecha.Text)
        End If
        'Actualizo las Celdas de la Grilla
        G.TextArray(GRIDINDEX(G, G.row, 9)) = Format(Actualizacion, "###,##0.00")
        G.TextArray(GRIDINDEX(G, G.row, 10)) = Format(Interes, "###,##0.00")
        G.TextArray(GRIDINDEX(G, G.row, 11)) = Format(CDbl(G.TextArray(GRIDINDEX(G, G.row, 8))) + CDbl(G.TextArray(GRIDINDEX(G, G.row, 9))) + CDbl(G.TextArray(GRIDINDEX(G, G.row, 10))), "###,##0.00")
        
        ''Actualizo los Caption de los Totales
        'FrmReciboContribuciones.LblActualizacionMatricula.Caption = Format(CDbl(FrmReciboContribuciones.LblActualizacionMatricula) + Actualizacion, "###,##0.00")
        'FrmReciboContribuciones.LblRecargoMatricula.Caption = Format(CDbl(FrmReciboContribuciones.LblRecargoMatricula) + Interes, "###,##0.00")
        'FrmReciboContribuciones.LblTotalMatricula.Caption = Format(CDbl(FrmReciboContribuciones.LblTotalMatricula) + Actualizacion + Interes, "###,##0.00")
        'FrmReciboContribuciones.LblTotal.Caption = Format(CDbl(FrmReciboContribuciones.LblTotal.Caption) + CDbl(G.TextArray(GRIDINDEX(G, G.row, 11))), "###,##0.00")
    Else
        ''Actualizo los Caption de los Totales
        'FrmReciboContribuciones.LblActualizacionMatricula.Caption = Format(CDbl(FrmReciboContribuciones.LblActualizacionMatricula) - CDbl(G.TextArray(GRIDINDEX(G, G.row, 9))), "###,##0.00")
        'FrmReciboContribuciones.LblRecargoMatricula.Caption = Format(CDbl(FrmReciboContribuciones.LblRecargoMatricula) - CDbl(G.TextArray(GRIDINDEX(G, G.row, 10))), "###,##0.00")
        'FrmReciboContribuciones.LblTotalMatricula.Caption = Format(CDbl(FrmReciboContribuciones.LblTotalMatricula) - CDbl(G.TextArray(GRIDINDEX(G, G.row, 9))) - CDbl(G.TextArray(GRIDINDEX(G, G.row, 10))), "###,##0.00")
        'FrmReciboContribuciones.LblTotal.Caption = Format(CDbl(FrmReciboContribuciones.LblTotal.Caption) - CDbl(G.TextArray(GRIDINDEX(G, G.row, 11))), "###,##0.00")
        
        'Actualizo las Celdas de la Grilla
        G.TextArray(GRIDINDEX(G, G.row, 9)) = Format(Actualizacion, "###,##0.00")
        G.TextArray(GRIDINDEX(G, G.row, 10)) = Format(Interes, "###,##0.00")
        G.TextArray(GRIDINDEX(G, G.row, 11)) = Format(CDbl(G.TextArray(GRIDINDEX(G, G.row, 8))) + CDbl(G.TextArray(GRIDINDEX(G, G.row, 9))) + CDbl(G.TextArray(GRIDINDEX(G, G.row, 10))), "###,##0.00")
   End If
   FrmReciboContribuciones.Refresh
End Sub

Public Sub CambiaColorAColumnaDeGrilla(G As MSFlexGrid, COLUMNA As Integer, ColorFuente)
    Dim j As Integer
    Dim ActColumna As Integer
    ActColumna = G.Col
    G.Col = ActColumna
    For j = 0 To G.Rows - 1
        G.row = j
        G.CellBackColor = ColorFuente
    Next
    G.Col = ActColumna
End Sub

Public Function CompletarConCeros(Tramite As String, Cnt As Integer) As String
    Dim I As Integer
    Dim NueCadena As String
    If Len(Trim(Tramite)) < Cnt Then
        I = 1
        Do While I <= Cnt - Len(Trim(Tramite))
           NueCadena = NueCadena & "0"
           I = I + 1
        Loop
        NueCadena = NueCadena & Trim(Tramite)
    Else
        NueCadena = Tramite
    End If
    CompletarConCeros = NueCadena
End Function

Public Function CompletarConEspaciosIzq(Tramite As String, Cnt As Integer) As String
    Dim I As Integer
    Dim NueCadena As String
    If Len(Trim(Tramite)) <= Cnt Then
        I = 1
        Do While I <= Cnt - Len(Trim(Tramite))
           NueCadena = NueCadena & " "
           I = I + 1
        Loop
        NueCadena = NueCadena & Trim(Tramite)
    End If
    CompletarConEspaciosIzq = NueCadena
End Function

Public Function CompletarConEspaciosIz(TEXTO As String, Cnt As Integer) As String
    'Devuelve una cadena compuesta por el texto pasado por par�metro y le anexa
    'blancos a la derecha, hasta que la longitud del texto y los blancos sea
    'igual a cnt
    CompletarConEspaciosIz = Space(Cnt - Len(TEXTO)) & TEXTO
End Function
Public Sub Set_Impresora()
    Printer.KillDoc
    Printer.Orientation = 1                 ' vbPRORLandscape '2 A lo ancho.
    Printer.PrintQuality = -3               ' -1=Borrador   -3=Media
    Printer.FontBold = False
    Printer.PaperSize = 1                   ' Carta 216 * 279 mm
    'Printer.PaperSize = 256                   ' Carta 216 * 279 mm
    Printer.ScaleMode = 7
    Printer.Height = 30700 '12190 '8642
    Printer.Width = 21000 '12220
    '-----------------------------------------------------------------
    ' EN W95 ANDA PERFECTO con (IBM Graphigs) es como DOS de una sola pasada Draft 10
    Printer.Font = "Roman 10cpi"
    Printer.FontSize = 10
    '-----------------------------------------------------------------
    ' EN W95 ANDA PERFECTO con (IBM Graphigs) es como DOS de una sola pasada Draft 15
'    Printer.Font = "Pica Compressed"
'    Printer.FontSize = 20
    '-----------------------------------------------------------------
    
End Sub

Public Sub Imprimir(ejex As Double, ejey As Double, remarcada As Boolean, TEXTO As String)
    If ejey >= 33 Then
        ejey = ejey
        Printer.NewPage
    End If
    'Printer.Font = "Pica Compressed"
    'Printer.Font = "Roman"
    'Printer.FontSize = 13
    Printer.FontBold = remarcada
    Printer.PSet (ejex + 0.2, ejey)
    Printer.Print TEXTO
End Sub

Function GRIDINDEX(Grid As MSFlexGrid, row As Integer, Col As Integer) As Long
'Devuelve el valor que apunta a una celda de una grilla
'para utilizarlo con el TextArray
'Autor: COLO
     GRIDINDEX = row * Grid.Cols + Col
End Function

Public Function Consulto_Cheque_Gestion_Judicial(TIPO As String, NRO As String) As Boolean

    Dim RecC As ADODB.Recordset
    Set RecC = New ADODB.Recordset
    Dim cSQLC As String
    'Consulto si tiene Cheques en G. Judicial o Rechazados por cualquier motivo
    cSQLC = "Select CHE_NUMERO " & _
            "  FROM CHEQUEESTADOVIGENTE " & _
            " WHERE TIP_TIPDOC = " & XS(TIPO) & _
            "   AND PER_NRODOC = " & XS(NRO) & _
            "   AND ECH_CODIGO BETWEEN 8 AND 24 "
    RecC.Open cSQLC, DBConn, adOpenStatic, adLockOptimistic
    If RecC.EOF = False Then
        Consulto_Cheque_Gestion_Judicial = True
    Else
        Consulto_Cheque_Gestion_Judicial = False
    End If
End Function

Public Function CONSULTO_DEUDA_PARTICULAR(TIPO As String, NRO As String, _
                                          FECHADEUDA As String, _
                                          RUBRO As Integer, _
                                          Sistema As Integer, _
                                          Subsistema As Integer) As String
                                          
'Funci�n que me devuelve los meses de deuda y el Importe de un Subsistema en Particular
'Para utilizarla incluir en el evento
'Dim Cadena As String
'Dim Cant_meses As Integer
'Dim Monto As Double
'Cadena = CONSULTO_DEUDA_PARTICULAR(TIP_TIPDOC, PER_NRODOC, Date, RUB_CODIGO, SIS_CODIGO, SUB_CODIGO)
'Cant_Meses = TextoPrevioAlGuion(Cadena)
'Monto = TextoPostAlGuion(Cadena)
'Autor: Eliana
        Dim RecDEUDA As ADODB.Recordset
        Set RecDEUDA = New ADODB.Recordset
        
        Dim Acumulado  As String
        Dim Cant_Meses As String
                
        'Acumulan el Importe del Subsistema
        Acumulado = "0"
        
        'Contador de los meses que debe
        Cant_Meses = "0"
        
        RecDEUDA.Open SELECT_DEUDA_PARTICULAR(TIPO, NRO, FECHADEUDA, False, RUBRO, Sistema, Subsistema), DBConn, adOpenStatic, adLockOptimistic
        If RecDEUDA.EOF = False Then
           'RecDEUDA.MoveFirst
           Do While Not RecDEUDA.EOF
              Cant_Meses = Cant_Meses + 1
              Acumulado = Acumulado + RecDEUDA.Fields(0)
           RecDEUDA.MoveNext
           Loop
           CONSULTO_DEUDA_PARTICULAR = Cant_Meses + "-" + Acumulado
        Else
           'El Matriculado NO adeuda
           CONSULTO_DEUDA_PARTICULAR = "0" + "-" + "0"
        End If
       RecDEUDA.Close
End Function

Public Function SELECT_DEUDA_PARTICULAR(TipoDoc As String, Nrodoc As String, _
                                        Fecha As String, completa As Boolean, _
                                        mRubro As Integer, _
                                        mSistema As Integer, _
                                        mSubsistema As Integer) As String
        
        cSQL = " Select Cta_Cte_SubSist.Ccs_ValNom,Cta_Cte_SubSist.Cta_NumCed,"
        cSQL = cSQL & " Cta_Cte_SubSist.Rub_Codigo,Rubros.Rub_Descri, "
        cSQL = cSQL & " Cta_Cte_SubSist.Sis_Codigo,Sistema.Sis_Descri,"
        cSQL = cSQL & " Cta_Cte_SubSist.Sub_Codigo,SubSistema.Sub_Descri,"
        cSQL = cSQL & " V1.Ven_FecVto,"
        cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,3,4),"
        cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,1,2)"
        cSQL = cSQL & " From Cta_Cte,Cta_Cte_SubSist,Periodo,Vencimiento V1,Sistema,SubSistema,Rubros"
        'cSQL = " Select Cta_Cte_SubSist.Ccs_ValNom"
        'cSQL = cSQL & " From Cta_Cte,Cta_Cte_SubSist,Periodo,Vencimiento V1"
        cSQL = cSQL & " Where Cta_Cte.Tip_TipDoc = " & XS(Trim(TipoDoc))
        cSQL = cSQL & " And   Cta_Cte.Per_NroDoc = " & XS(Trim(Nrodoc))
        cSQL = cSQL & " And   Cta_Cte.Cta_NumCed = Cta_Cte_SubSist.Cta_NumCed "
        'cSQL = cSQL & " And   Cta_Cte_SubSist.Ccs_Pagado = 0 "
        cSQL = cSQL & " And (Cta_Cte_SubSist.Ccs_Pagado = 0 OR (Cta_Cte_SubSist.Ccs_Pagado = 1 AND Cta_Cte_SubSist.CCS_FECPAG > " & XD(Fecha) & " ))  "
        cSQL = cSQL & " And   Cta_Cte_SubSist.RUB_CODIGO = " & mRubro
        cSQL = cSQL & " And   Cta_Cte_SubSist.SIS_CODIGO = " & mSistema
        cSQL = cSQL & " And   Cta_Cte_SubSist.SUB_CODIGO = " & mSubsistema
        cSQL = cSQL & " And   Cta_Cte.Pdo_Period = Periodo.Pdo_Period "
        cSQL = cSQL & " And   Periodo.Pdo_Period = V1.Pdo_Period "
        cSQL = cSQL & " And   V1.Ven_FecVto = (Select Min(V2.Ven_FecVto) "
        cSQL = cSQL & "                          From Vencimiento V2 "
        cSQL = cSQL & "                         Where V2.Pdo_Period = Cta_Cte.Pdo_Period "
        cSQL = cSQL & "                           And V2.Ven_FecVto >= " & XD(Fecha) & ")"
        cSQL = cSQL & " And " & XD(Fecha) & " > (Select Min(V2.Ven_FecVto)  "
        cSQL = cSQL & "                            From Vencimiento V2 "
        cSQL = cSQL & "                           Where V2.Pdo_Period = Cta_Cte.Pdo_Period)"
        cSQL = cSQL & " And Rubros.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And Sistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And Sistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
        cSQL = cSQL & " And Subsistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And SubSistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
        cSQL = cSQL & " And SubSistema.Sub_Codigo=Cta_Cte_SubSist.Sub_Codigo"
                
        cSQL = cSQL & " Union All "
        
        'cSQL = cSQL & " Select Cta_Cte_SubSist.Ccs_ValNom"
        'cSQL = cSQL & " From Cta_Cte,Cta_Cte_SubSist,Periodo,Vencimiento V1"
        cSQL = cSQL & " Select Cta_Cte_SubSist.Ccs_ValNom,Cta_Cte_SubSist.Cta_NumCed,"
        cSQL = cSQL & " Cta_Cte_SubSist.Rub_Codigo,Rubros.Rub_Descri, "
        cSQL = cSQL & " Cta_Cte_SubSist.Sis_Codigo,Sistema.Sis_Descri,"
        cSQL = cSQL & " Cta_Cte_SubSist.Sub_Codigo,SubSistema.Sub_Descri,"
        cSQL = cSQL & " V1.Ven_FecVto,"
        cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,3,4),"
        cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,1,2)"
        cSQL = cSQL & " From Cta_Cte,Cta_Cte_SubSist,Periodo,Vencimiento V1,Sistema,SubSistema,Rubros"
        
        cSQL = cSQL & " Where Cta_Cte.Tip_TipDoc = " & XS(Trim(TipoDoc))
        cSQL = cSQL & "   And Cta_Cte.Per_NroDoc = " & XS(Trim(Nrodoc))
        cSQL = cSQL & "   And Cta_Cte.Cta_NumCed = Cta_Cte_SubSist.Cta_NumCed "
        'cSQL = cSQL & "   And Cta_Cte_SubSist.Ccs_Pagado = 0 "
        cSQL = cSQL & "   And (Cta_Cte_SubSist.Ccs_Pagado = 0 OR (Cta_Cte_SubSist.Ccs_Pagado = 1 AND Cta_Cte_SubSist.CCS_FECPAG > " & XD(Fecha) & " ))  "
        cSQL = cSQL & "   And Cta_Cte_SubSist.RUB_CODIGO = " & mRubro
        cSQL = cSQL & "   And Cta_Cte_SubSist.SIS_CODIGO = " & mSistema
        cSQL = cSQL & "   And Cta_Cte_SubSist.SUB_CODIGO = " & mSubsistema
        cSQL = cSQL & "   And Cta_Cte.Pdo_Period = Periodo.Pdo_Period "
        cSQL = cSQL & "   And Periodo.Pdo_Period = V1.Pdo_Period "
        cSQL = cSQL & "   And V1.Ven_FecVto = (Select Max(V2.Ven_FecVto) "
        cSQL = cSQL & "                          From Vencimiento V2 "
        cSQL = cSQL & "                         Where V2.Pdo_Period = Cta_Cte.Pdo_Period "
        cSQL = cSQL & "                           And V2.Ven_FecVto < " & XD(Fecha) & " ) "
        cSQL = cSQL & "   And " & XD(Fecha) & " > (Select Max(V2.Ven_FecVto) "
        cSQL = cSQL & "                              From Vencimiento V2 "
        cSQL = cSQL & "                             Where V2.Pdo_Period = Cta_Cte.Pdo_Period)"
        cSQL = cSQL & " And Rubros.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And Sistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And Sistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
        cSQL = cSQL & " And Subsistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
        cSQL = cSQL & " And SubSistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
        cSQL = cSQL & " And SubSistema.Sub_Codigo=Cta_Cte_SubSist.Sub_Codigo"
        
        If completa Then
            cSQL = cSQL & " Union All "
            'cSQL = cSQL & " Select Cta_Cte_SubSist.Ccs_ValNom"
            'cSQL = cSQL & " From Cta_Cte,Cta_Cte_SubSist,Periodo,Vencimiento V1"
            
            cSQL = cSQL & " Select Cta_Cte_SubSist.Ccs_ValNom,Cta_Cte_SubSist.Cta_NumCed,"
            cSQL = cSQL & " Cta_Cte_SubSist.Rub_Codigo,Rubros.Rub_Descri, "
            cSQL = cSQL & " Cta_Cte_SubSist.Sis_Codigo,Sistema.Sis_Descri,"
            cSQL = cSQL & " Cta_Cte_SubSist.Sub_Codigo,SubSistema.Sub_Descri,"
            cSQL = cSQL & " V1.Ven_FecVto,"
            cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,3,4),"
            cSQL = cSQL & " substring(Cta_Cte_SubSist.Cta_NumCed,1,2)"
            cSQL = cSQL & " From Cta_Cte,Cta_Cte_SubSist,Periodo,Vencimiento V1,Sistema,SubSistema,Rubros"
        
            cSQL = cSQL & " Where Cta_Cte.Tip_TipDoc = " & XS(Trim(TipoDoc))
            cSQL = cSQL & "   And Cta_Cte.Per_NroDoc = " & XS(Trim(Nrodoc))
            cSQL = cSQL & "   And Cta_Cte.Cta_NumCed = Cta_Cte_SubSist.Cta_NumCed "
            'cSQL = cSQL & "   And Cta_Cte_SubSist.Ccs_Pagado = 0 "
            cSQL = cSQL & "   And (Cta_Cte_SubSist.Ccs_Pagado = 0 OR (Cta_Cte_SubSist.Ccs_Pagado = 1 AND Cta_Cte_SubSist.CCS_FECPAG > " & XD(Fecha) & " ))  "
            cSQL = cSQL & "   And Cta_Cte_SubSist.RUB_CODIGO = " & mRubro
            cSQL = cSQL & "   And Cta_Cte_SubSist.SIS_CODIGO = " & mSistema
            cSQL = cSQL & "   And Cta_Cte_SubSist.SUB_CODIGO = " & mSubsistema
            cSQL = cSQL & "   And Cta_Cte.Pdo_Period = Periodo.Pdo_Period "
            cSQL = cSQL & "   And Periodo.Pdo_Period = V1.Pdo_Period "
            cSQL = cSQL & "   And V1.Ven_FecVto =(Select Min(V2.Ven_FecVto) "
            cSQL = cSQL & "                         From Vencimiento V2 "
            cSQL = cSQL & "                        Where V2.Pdo_Period=Cta_Cte.Pdo_Period) "
            cSQL = cSQL & "   And V1.Ven_FecVto >=" & XD(Fecha)
            cSQL = cSQL & " And Rubros.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
            cSQL = cSQL & " And Sistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
            cSQL = cSQL & " And Sistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
            cSQL = cSQL & " And Subsistema.Rub_Codigo=Cta_Cte_SubSist.Rub_Codigo "
            cSQL = cSQL & " And SubSistema.Sis_Codigo=Cta_Cte_SubSist.Sis_Codigo "
            cSQL = cSQL & " And SubSistema.Sub_Codigo=Cta_Cte_SubSist.Sub_Codigo"
        
        End If
        cSQL = cSQL & " ORDER BY substring(Cta_Cte_SubSist.Cta_NumCed,3,4),"
        cSQL = cSQL & "          substring(Cta_Cte_SubSist.Cta_NumCed,1,2)"
        
        SELECT_DEUDA_PARTICULAR = cSQL
End Function

Public Sub OrdenarGrilla(Grilla As MSFlexGrid)
'Ordena la Grilla por la columna en donde se hizo click con el mouse
'Para utilizarla incluir en el evento click de la Grilla:
'OrdenarGrilla NombreDeLaGrilla
'Autor: COLO
    If Grilla.MouseRow = 0 Then
        Screen.MousePointer = vbHourglass
        Grilla.Col = Grilla.MouseCol
        Grilla.ColSel = Grilla.MouseCol
        Grilla.Sort = 1
        Screen.MousePointer = vbNormal
    End If
End Sub

Public Function ConfirmoReciboTURISMO(NroFicticio As String, NroReal As String, Estado As Integer) As Boolean
    ConfirmoReciboTURISMO = False
    On Error GoTo ErrorTrans
       
      Set Rec1 = New ADODB.Recordset
       
      DBConn.Execute "Update doc_cpce " & _
                        "set dcpce_nrorecibo = " & XS(NroReal) & ",dcpce_estado = " & Estado & _
                     " where dcpce_nrorecibo = " & XS(NroFicticio)
                   
      If Estado = 3 Then 'Recibo Cobrado
          'Actualizo el Saldo de la Reserva
          cSQL = "Select dcpce_monto,res_nro " & _
                 "  From doc_cpce " & _
                 " Where dcpce_nrorecibo = " & XS(NroReal)
          Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
          If Rec1.EOF = False Then
             sql = " UPDATE RESERVAS " & _
                      " SET RES_SALDO = RES_SALDO - " & XN(Rec1!dcpce_monto) & _
                    " WHERE RES_NRO = " & XN(Rec1!res_nro)
             DBConn.Execute sql, dbExecDirect
          End If
          Rec1.Close
      End If
      ConfirmoReciboTURISMO = True
      Exit Function
    
ErrorTrans:
    ConfirmoReciboTURISMO = False
End Function

Public Function ConfirmoReciboCURSO(NroFicticio As String, NroReal As String, Estado As Integer) As Boolean
    ConfirmoReciboCURSO = False
    On Error GoTo ErrorTrans
       
    Set rec0 = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
     
    DBConn.Execute "Update PAGOS_CURSOS " & _
                   "set pcu_nrorec = " & XS(NroReal) & ", pcu_estado = " & Estado & _
                   " where pcu_nrorec = " & XS(NroFicticio)
                 
    If Estado = 2 Then 'Recibo ANULADO
        'Busco el Detalle del Recibo
        cSQL = "SELECT REC_NUMERO,DET_NROITEM,RUB_CODIGO, SIS_CODIGO, SUB_CODIGO, DET_MONTO " & _
                " FROM DETALLE_RECIBO " & _
                "WHERE REC_NUMERO = " & XS(NroReal)
        rec0.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
        If rec0.EOF = False Then
           'Rec0.MoveFirst
           Do While rec0.EOF = False
              If Trim(rec0!SUB_CODIGO) = 99 Then 'Item de Material
                    
                    'Busco los Datos del Recibo
                    cSQL = "Select pcu_monto,del_codigo,tcu_codigo,tem_codigo,cur_fecdes," & _
                                  "iac_tipdoc,iac_nrodoc " & _
                           "  From PAGOS_CURSOS " & _
                           " Where pcu_nrorec = " & XS(NroReal)
                    Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                    If Rec1.EOF = False Then
                       sql = "UPDATE INSCRIPCIONES_A_CURSOS " & _
                               " SET IAC_SALTOT = IAC_SALTOT + " & XN(rec0!DET_MONTO) & _
                                 " , IAC_SALMAT = IAC_SALMAT + " & XN(rec0!DET_MONTO) & _
                             " WHERE del_codigo = " & XN(Rec1!DEL_CODIGO) & _
                               " AND tcu_codigo = " & XN(Rec1!TCU_CODIGO) & _
                               " AND tem_codigo = " & XN(Rec1!tem_codigo) & _
                               " AND cur_fecdes = " & XD(Rec1!cur_fecdes) & _
                               " AND iac_tipdoc = " & XS(Rec1!IAC_TIPDOC) & _
                               " AND iac_nrodoc = " & XS(Rec1!IAC_NRODOC)
                       DBConn.Execute sql, dbExecDirect
                    End If
                    Rec1.Close
                    
                ElseIf Trim(rec0!SUB_CODIGO) <> 99 Then 'Item del CURSO
                    
                    'Busco los Datos del Recibo
                    cSQL = "Select pcu_monto,del_codigo,tcu_codigo,tem_codigo,cur_fecdes," & _
                                  "iac_tipdoc,iac_nrodoc " & _
                           "  From PAGOS_CURSOS " & _
                           " Where pcu_nrorec = " & XS(NroReal)
                    Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                    If Rec1.EOF = False Then
                       sql = "UPDATE INSCRIPCIONES_A_CURSOS " & _
                               " SET IAC_SALTOT = IAC_SALTOT + " & XN(rec0!DET_MONTO) & _
                                 " , IAC_SALDO = IAC_SALDO + " & XN(rec0!DET_MONTO) & _
                             " WHERE del_codigo = " & XN(Rec1!DEL_CODIGO) & _
                               " AND tcu_codigo = " & XN(Rec1!TCU_CODIGO) & _
                               " AND tem_codigo = " & XN(Rec1!tem_codigo) & _
                               " AND cur_fecdes = " & XD(Rec1!cur_fecdes) & _
                               " AND iac_tipdoc = " & XS(Rec1!IAC_TIPDOC) & _
                               " AND iac_nrodoc = " & XS(Rec1!IAC_NRODOC)
                       DBConn.Execute sql, dbExecDirect
                    End If
                    Rec1.Close
                End If
                rec0.MoveNext
            Loop
        End If
        rec0.Close
    End If
    ConfirmoReciboCURSO = True
    Exit Function
    
ErrorTrans:
    ConfirmoReciboCURSO = False
End Function

Public Function ConfirmoReciboContribucion(NroReal As String, Estado As Integer, Delegacion As String) As Boolean

    ConfirmoReciboContribucion = False
    
    On Error GoTo ErrorTrans
    
    Set Rec1 = New ADODB.Recordset
    
    If Estado = 3 Then
 
        cSQL = " Select d.cta_numced,d.del_codigo,d.rub_codigo,d.sis_codigo,d.sub_codigo, " & _
               " r.REC_FECHA,r.TIP_TIPDOC, r.PER_NRODOC " & _
               "   From detalle_recibo d, recibo r " & _
               "  Where r.DEL_CODIGO = d.DEL_CODIGO " & _
               "    and r.REC_NUMERO = d.REC_NUMERO " & _
               "    and R.REC_NUMERO = " & NroReal
        Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                If IsNull(Rec1!CTA_NUMCED) Then Exit Do 'NO ES UN RECIBO DE CONTRIBUCIONES
                cSQL = "Update cta_cte_subsist " & _
                     " set ccs_pagado = 1, " & _
                     "     ccs_fecpag = " & XD(Rec1!REC_FECHA) & _
                     "     ,DEL_CODIGO = " & XN(Delegacion) & _
                     "     ,REC_NUMERO = " & XN(NroReal) & _
                     " where cta_numced = " & XS(Rec1!CTA_NUMCED) & _
                     "   and rub_codigo = " & XN(Rec1!RUB_CODIGO) & _
                     "   and sis_codigo = " & XN(Rec1!SIS_CODIGO) & _
                     "   and sub_codigo = " & XN(Rec1!SUB_CODIGO) & _
                     "   and tipo_docu  = " & XS(Rec1!Tip_TipDoc) & _
                     "   and nro_docu   = " & XS(Rec1!Per_NroDoc)
                DBConn.Execute cSQL
                Rec1.MoveNext
            Loop
        Else
            GoTo ErrorTrans
        End If
        Rec1.Close
    ElseIf Estado = 2 Then
        cSQL = " Select d.cta_numced,d.del_codigo,d.rub_codigo,d.sis_codigo,d.sub_codigo, " & _
               " r.REC_FECHA,r.TIP_TIPDOC, r.PER_NRODOC " & _
               "   From detalle_recibo d, recibo r " & _
               "  Where r.DEL_CODIGO = d.DEL_CODIGO " & _
               "    and r.REC_NUMERO = d.REC_NUMERO " & _
               "    and R.REC_NUMERO = " & NroReal
        Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                If IsNull(Rec1!CTA_NUMCED) Then Exit Do 'NO ES UN RECIBO DE CONTRIBUCIONES
                ' Si es (6 = Recargos) o (7 = Actualizaci�n) o (9 = AJUSTE)
                If Right(Rec1!SUB_CODIGO, 1) = 6 _
                Or Right(Rec1!SUB_CODIGO, 1) = 7 _
                Or Right(Rec1!SUB_CODIGO, 1) = 9 Then
                    cSQL = "delete from cta_cte_subsist " & _
                         "   where cta_numced = " & XS(Rec1!CTA_NUMCED) & _
                         "   and rub_codigo = " & XN(Rec1!RUB_CODIGO) & _
                         "   and sis_codigo = " & XN(Rec1!SIS_CODIGO) & _
                         "   and sub_codigo = " & XN(Rec1!SUB_CODIGO) & _
                         "   and tipo_docu  = " & XS(Rec1!Tip_TipDoc) & _
                         "   and nro_docu   = " & XS(Rec1!Per_NroDoc)
                    DBConn.Execute cSQL
                Else
                    cSQL = "Update cta_cte_subsist " & _
                         " set ccs_pagado = 0, " & _
                         "     ccs_fecpag = null, " & _
                         "     DEL_CODIGO = null, " & _
                         "     REC_NUMERO = null" & _
                         " where cta_numced = " & XS(Rec1!CTA_NUMCED) & _
                         "   and rub_codigo = " & XN(Rec1!RUB_CODIGO) & _
                         "   and sis_codigo = " & XN(Rec1!SIS_CODIGO) & _
                         "   and sub_codigo = " & XN(Rec1!SUB_CODIGO) & _
                         "   and tipo_docu  = " & XS(Rec1!Tip_TipDoc) & _
                         "   and nro_docu   = " & XS(Rec1!Per_NroDoc)
                    DBConn.Execute cSQL
                End If
                Rec1.MoveNext
            Loop
        Else
            GoTo ErrorTrans
        End If
        Rec1.Close
    End If
    ConfirmoReciboContribucion = True
    Exit Function
    
ErrorTrans:
    If Rec1.State = 1 Then Rec1.Close
    ConfirmoReciboContribucion = False
End Function

Public Function ConfirmoReciboTecnica(NroFicticio As String, NroReal As String, Estado As String, Dele As String) As Boolean
    Dim ReciboTramite As Boolean
    Dim ObleasEmitidas As Boolean
    Dim DelObl As String
    Dim NroObl As String
    Dim CD As String
    Dim CFo As String
    Dim CFc As String
    Dim CDR As String
    Dim NroTra As String
    Dim TicCod As String
    Dim FecIng As String
    Dim DELCOD As String
    Dim cExec As String
    Dim I As Integer
    Dim Ars As ADODB.Recordset
    Set Ars = New ADODB.Recordset
    ConfirmoReciboTecnica = False
    On Error GoTo MiError
    CD = "0"
    CFo = "0"
    CFc = "0"
    CDR = "0"
    'ACTUALIZO LAS TABLAS DE RECIBO EN TECNICA
    cExec = " Update RECIBO_TECNICA Set ERE_CODIGO=" & XN(Estado)
    cExec = cExec & " , REC_NUMERO=" & XN(NroReal)
    cExec = cExec & " Where REC_NUMERO=" & XN(NroFicticio)
    cExec = cExec & "   And DEL_RECIBO=" & XN(Dele)
    DBConn.Execute cExec
    cExec = " Update DETALLE_RECIBO_TECNICA Set REC_NUMERO=" & XN(NroReal)
    cExec = cExec & " Where REC_NUMERO=" & XN(NroFicticio)
    cExec = cExec & "   And DEL_RECIBO=" & XN(Dele)
    DBConn.Execute cExec
    'ACTUALIZO PAGO EN OBLEAS
    'Averiguo #Tramite
    If Estado = "3" Then
        cExec = " Select TIC_CODIGO,DEL_CODIGO,TRA_FECING,TRA_NUMERO From RECIBO_TECNICA "
        cExec = cExec & " Where REC_NUMERO=" & XN(NroReal)
        cExec = cExec & "   And DEL_RECIBO=" & XN(Dele)
        Ars.Open cExec, DBConn, adOpenStatic, adLockOptimistic
        If Ars.EOF = False Then
            'Ars.MoveFirst
            TicCod = Ars!TIC_CODIGO
            DELCOD = Ars!DEL_CODIGO
            NroTra = Ars!TRA_NUMERO
            FecIng = Ars!TRA_FECING
            ReciboTramite = True
        Else
            ReciboTramite = False
        End If
        Ars.Close
        If ReciboTramite Then
            'Verifico si ya se emitieron obleas
            cExec = " Select Count(*) As Cnt  From OBLEAS"
            cExec = cExec & " Where TIC_CODIGO=" & XN(TicCod)
            cExec = cExec & "   And DEL_CODIGO=" & XN(DELCOD)
            cExec = cExec & "   And TRA_FECING=" & XD(FecIng)
            cExec = cExec & "   And TRA_NUMERO=" & XN(NroTra)
            Ars.Open cExec, DBConn, adOpenStatic, adLockOptimistic
            If Ars!Cnt > 0 Then
                ObleasEmitidas = True
            Else
                ObleasEmitidas = False
            End If
            Ars.Close
            If ObleasEmitidas Then
                'Averiguo que se pag�
                cExec = " Select SUB_CODIGO,DRT_CANTID FROM DETALLE_RECIBO_TECNICA "
                cExec = cExec & " Where REC_NUMERO=" & XN(NroReal)
                cExec = cExec & "   And DEL_RECIBO=" & XN(Dele)
                Ars.Open cExec, DBConn, adOpenStatic, adLockOptimistic
                If Ars.EOF = False Then
                    'Ars.MoveFirst
                    Do While Not Ars.EOF
                        Select Case Trim(Ars!SUB_CODIGO)
                            Case "60"
                                CD = CDbl(CD) + CDbl(Ars!DRT_CANTID)
                            Case "64"
                                CFc = CDbl(CFc) + CDbl(Ars!DRT_CANTID)
                            Case "65"
                                CFo = CDbl(CFo) + CDbl(Ars!DRT_CANTID)
                            Case "66"
                                CDR = CDbl(CDR) + CDbl(Ars!DRT_CANTID)
                        End Select
                        Ars.MoveNext
                    Loop
                End If
                Ars.Close
                'Actualizo Obleas
                If CDR <> "0" Then
                    I = Val(CDR)
                    cExec = " Select DEL_CODIGO,OBL_NUMERO From OBLEAS "
                    cExec = cExec & " Where TIC_CODIGO=" & XN(TicCod)
                    cExec = cExec & "   And DEL_TRAMIT=" & XN(DELCOD)
                    cExec = cExec & "   And TRA_FECING=" & XD(FecIng)
                    cExec = cExec & "   And TRA_NUMERO=" & XN(NroTra)
                    cExec = cExec & "   And OBL_PAGO IS NULL "
                    cExec = cExec & "   And OBL_CODEST='OK' "
                    cExec = cExec & "   And (TOB_CODIGO=3 Or TOB_CODIGO=4)"
                    Ars.Open cExec, DBConn, adOpenStatic, adLockOptimistic
                    If Ars.EOF = False Then
                        'Ars.MoveFirst
                        Do While Not Ars.EOF And I > 0
                            cExec = " Update OBLEAS Set OBL_PAGO=" & XS(Trim(Dele) & "-" & Trim(NroReal))
                            cExec = cExec & " , OBL_TIPPAG='R' "
                            cExec = cExec & " Where DEL_CODIGO=" & XN(Ars!DEL_CODIGO)
                            cExec = cExec & "   And OBL_NUMERO=" & XN(Ars!OBL_NUMERO)
                            I = I - 1
                            DBConn.Execute cExec
                            Ars.MoveNext
                        Loop
                    End If
                    Ars.Close
                End If
                
                If CD <> "0" Then
                    I = Val(CD)
                    cExec = " Select DEL_CODIGO,OBL_NUMERO From OBLEAS "
                    cExec = cExec & " Where TIC_CODIGO=" & XN(TicCod)
                    cExec = cExec & "   And DEL_TRAMIT=" & XN(DELCOD)
                    cExec = cExec & "   And TRA_FECING=" & XD(FecIng)
                    cExec = cExec & "   And TRA_NUMERO=" & XN(NroTra)
                    cExec = cExec & "   And OBL_PAGO IS NULL "
                    cExec = cExec & "   And OBL_CODEST='OK' "
                    cExec = cExec & "   And (TOB_CODIGO=3 Or TOB_CODIGO=4)"
                    Ars.Open cExec, DBConn, adOpenStatic, adLockOptimistic
                    If Ars.EOF = False Then
                        'Ars.MoveFirst
                        Do While Not Ars.EOF And I > 0
                            cExec = " Update OBLEAS Set OBL_PAGO=" & XS(Trim(Dele) & "-" & Trim(NroReal))
                            cExec = cExec & " , OBL_TIPPAG='R' "
                            cExec = cExec & " Where DEL_CODIGO=" & XN(Ars!DEL_CODIGO)
                            cExec = cExec & "   And OBL_NUMERO=" & XN(Ars!OBL_NUMERO)
                            I = I - 1
                            DBConn.Execute cExec
                            Ars.MoveNext
                        Loop
                    End If
                    Ars.Close
                End If
                If CFc <> "0" Then
                    I = Val(CFc)
                    cExec = " Select DEL_CODIGO,OBL_NUMERO From OBLEAS "
                    cExec = cExec & " Where TIC_CODIGO=" & XN(TicCod)
                    cExec = cExec & "   And DEL_TRAMIT=" & XN(DELCOD)
                    cExec = cExec & "   And TRA_FECING=" & XD(FecIng)
                    cExec = cExec & "   And TRA_NUMERO=" & XN(NroTra)
                    cExec = cExec & "   And OBL_PAGO IS NULL "
                    cExec = cExec & "   And OBL_CODEST='OK' "
                    cExec = cExec & "   And (TOB_CODIGO=1 Or TOB_CODIGO=2)"
                    Ars.Open cExec, DBConn, adOpenStatic, adLockOptimistic
                    If Ars.EOF = False Then
                        'Ars.MoveFirst
                        Do While Not Ars.EOF And I > 0
                            cExec = " Update OBLEAS Set OBL_PAGO=" & XS(Trim(Dele) & "-" & Trim(NroReal))
                            cExec = cExec & " , OBL_TIPPAG='R' "
                            cExec = cExec & " Where DEL_CODIGO=" & XN(Ars!DEL_CODIGO)
                            cExec = cExec & "   And OBL_NUMERO=" & XN(Ars!OBL_NUMERO)
                            I = I - 1
                            DBConn.Execute cExec
                            Ars.MoveNext
                        Loop
                    End If
                    Ars.Close
                End If
                If CFo <> "0" Then
                    I = Val(CFo)
                    cExec = " Select DEL_CODIGO,OBL_NUMERO From OBLEAS "
                    cExec = cExec & " Where TIC_CODIGO=" & XN(TicCod)
                    cExec = cExec & "   And DEL_TRAMIT=" & XN(DELCOD)
                    cExec = cExec & "   And TRA_FECING=" & XD(FecIng)
                    cExec = cExec & "   And TRA_NUMERO=" & XN(NroTra)
                    cExec = cExec & "   And OBL_PAGO IS NULL "
                    cExec = cExec & "   And OBL_CODEST='OK' "
                    cExec = cExec & "   And (TOB_CODIGO=1 Or TOB_CODIGO=2)"
                    Ars.Open cExec, DBConn, adOpenStatic, adLockOptimistic
                    If Ars.EOF = False Then
                        'Ars.MoveFirst
                        Do While Not Ars.EOF And I > 0
                            cExec = " Update OBLEAS Set OBL_PAGO=" & XS(Trim(Dele) & "-" & Trim(NroReal))
                            cExec = cExec & " , OBL_TIPPAG='R' "
                            cExec = cExec & " Where DEL_CODIGO=" & XN(Ars!DEL_CODIGO)
                            cExec = cExec & "   And OBL_NUMERO=" & XN(Ars!OBL_NUMERO)
                            I = I - 1
                            DBConn.Execute cExec
                            Ars.MoveNext
                        Loop
                    End If
                    Ars.Close
                End If
                
            End If
        End If
    ElseIf Estado = "2" Then
        cExec = " Update OBLEAS Set OBL_PAGO = NULL "
        cExec = cExec & " , OBL_TIPPAG=NULL "
        cExec = cExec & " Where OBL_PAGO=" & XS(Trim(Dele) & "-" & Trim(NroReal))
        DBConn.Execute cExec
    End If
    ConfirmoReciboTecnica = True
    Exit Function
    
MiError:
    ConfirmoReciboTecnica = False
End Function

Public Function ConfirmaReciboPrestamo(mNroViejo As String, mNroNuevo As String, mEstado As Integer, mDelCob As String, mTDoc As String, mNDoc As String)
    Dim RecCuo As New ADODB.Recordset, RecPPA As New ADODB.Recordset, RecCon As New ADODB.Recordset
    Dim mTipPrest As String, mNroPrest As String, mCodConce As String, mNroCuota As Integer, mNroSubCu As Integer, mCanCuota As Integer, mEstPrest As String
    ConfirmaReciboPrestamo = True
    On Error GoTo ErrorGuardar
    
    sql = "SELECT * FROM servicios s,prestamos p "
    sql = sql & "WHERE s.tpr_codtippr=p.tpr_codtippr"
    sql = sql & " AND  s.pre_nroprest=p.pre_nroprest"
    sql = sql & " AND  p.per_nrodoc=" + XS(mNDoc)
    sql = sql & " AND  p.tip_tipdoc=" + XS(mTDoc)
    sql = sql & " AND  s.rec_numero= " + XS(mNroViejo)
    sql = sql + IIf(mEstado = 3, " AND  s.ser_feccobro IS NULL ", " AND  (s.ser_feccobro IS NOT NULL OR s.ser_feccobro IS NULL)")
    RecCuo.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If RecCuo.RecordCount = 1 Then
        mTipPrest = RecCuo.Fields!TPR_CODTIPPR
        mNroPrest = RecCuo.Fields!pre_nroprest
        mCodConce = RecCuo.Fields!cse_codconce
        mNroCuota = RecCuo.Fields!ser_nrocuota
        mNroSubCu = RecCuo.Fields!ser_nrosubcu
        mEstPrest = RecCuo.Fields!est_estprest
        '-------- Controlo que la CTA anterior a la que pretendo cobrar este paga.
        '---------- Fin del Control.
        If mEstado = 3 Then             ' Confirma
            sql = "SELECT * FROM servicios s "
            sql = sql & "WHERE s.tpr_codtippr=" + XS(mTipPrest)
            sql = sql & " AND  s.pre_nroprest=" + XS(mNroPrest)
            sql = sql & " AND  ser_nrocuota = " + XN(CDbl(mNroCuota - 1))
            sql = sql & " AND  ser_nrosubcu = " + XN(CDbl(mNroSubCu))
            sql = sql & " AND  cse_codconce = " + XS(mCodConce)
            RecCon.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If RecCon.RecordCount >= 1 Then
                If IsNull(RecCon!ser_feccobro) Then
                    GoTo ErrorControl
                End If
            End If
            RecCon.Close
            If mCodConce = "CTA" Then
                sql = "SELECT * FROM prest_plan " + _
                " WHERE pre_nroprest=" + XS(mNroPrest) + _
                " AND   tpr_codtippr=" + XS(mTipPrest) + _
                " AND   ppl_nroserie=(SELECT MAX(ppl_nroserie) FROM prest_plan " + _
                " WHERE pre_nroprest=" + XS(mNroPrest) + _
                " AND   tpr_codtippr=" + XS(mTipPrest) + ")"
                RecPPA.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Not RecPPA.EOF Then
                    mCanCuota = RecPPA.Fields!ppl_nropricu + RecPPA.Fields!ppl_cancuota - 1
                End If
                RecPPA.Close
            End If
            
            sql = "UPDATE servicios SET " + _
            "  ser_feccobro = " + XD(Date) + _
            " ,rec_numero = " + XS(mNroNuevo) + _
            " WHERE tpr_codtippr = " + XS(mTipPrest) + _
            " AND   pre_nroprest = " + XS(mNroPrest) + _
            " AND   ser_nrocuota = " + XN(CDbl(mNroCuota)) + _
            " AND   ser_nrosubcu = " + XN(CDbl(mNroSubCu)) + _
            " AND   cse_codconce = " + XS(mCodConce)
            DBConn.Execute sql$
            
            If mCanCuota = mNroCuota Or mCodConce = "CAN" Then
                sql = "UPDATE prestamos SET " + _
                "  est_estprest = 'CAN' " + _
                " WHERE tpr_codtippr = " + XS(mTipPrest) + _
                " AND   pre_nroprest = " + XS(mNroPrest)
                DBConn.Execute sql$
                ActualizaEstado mNroPrest, mTipPrest, "CAN", Date
            End If
        ElseIf mEstado = 2 Then             ' Anulado
            If mEstPrest = "CAN" Then       ' Caso pago ultima cuota y luego anulo.
                sql = "UPDATE prestamos SET " + _
                "  est_estprest = 'ACT'" + _
                " WHERE tpr_codtippr = " + XS(mTipPrest) + _
                " AND   pre_nroprest = " + XS(mNroPrest)
                DBConn.Execute sql$
                ActualizaEstado mNroPrest, mTipPrest, "ACT", Date
            End If
            sql = "UPDATE servicios SET " + _
            "  ser_imppunit = " + XN(CDbl(0)) + _
            " ,rec_numero = " + XN(CDbl(0)) + _
            " ,ser_feccobro = NULL  " + _
            " WHERE tpr_codtippr = " + XS(mTipPrest) + _
            " AND   pre_nroprest = " + XS(mNroPrest) + _
            " AND   ser_nrocuota = " + XN(CDbl(mNroCuota)) + _
            " AND   ser_nrosubcu = " + XN(CDbl(mNroSubCu)) + _
            " AND   cse_codconce = " + XS(mCodConce)
            DBConn.Execute sql$
        End If
    Else
        GoTo ErrorGuardar
    End If
    RecCuo.Close
    Exit Function
    
ErrorGuardar:
    MsgBox "Error al tratar de actualizar el Recibo de Prestamos "
    ConfirmaReciboPrestamo = False
    Exit Function
ErrorControl:
    MsgBox "Error: Existe una cuota anterior que no ha sido cobrada."
    ConfirmaReciboPrestamo = False
    Exit Function

End Function
    
Public Function ConfirmaReciboPUBLICACIONES(mNroViejo As String, mNroNuevo As String, mEstado As Integer, mDelega As String)

'Dim Rec2 As ADODB.Recordset
'Dim Rec3 As ADODB.Recordset
'Dim canti As Integer
'Set Rec2 = New ADODB.Recordset
'Set Rec3 = New ADODB.Recordset
    
    ConfirmaReciboPUBLICACIONES = False
    
    On Error GoTo ErrorTrans
        
    sql = "UPDATE PUBLICACIONES_VENTAS SET " + _
    "  ERE_CODIGO = " & mEstado & ", " & _
    "  REC_NUMERO = " & mNroNuevo & " " & _
    " WHERE " & _
    "  REC_NUMERO = '" & mNroViejo & "' AND " & _
    "  DEL_CODIGO = " & mDelega
    DBConn.Execute sql
    
'    sql = "select pub_cantidad, pub_codigo from publicaciones_ventas where" & _
'        "  REC_NUMERO = '" & mNroNuevo & "' AND " & _
'        "  DEL_CODIGO = " & mDelega
'    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec2.eof = false Then
'        If Val(mEstado) = 3 Then
'            DBConn.Execute "update stock_delega set sde_cantid = sde_cantid - " & Rec2!pub_cantidad & " where del_codigo = " & mDelega & _
'             " and pub_codigo = " & Rec2!PUB_CODIGO
'             'va restando desde la fecha mas vieja de compra lo vendido
'
'             canti = Rec2!pub_cantidad
'             Do While canti > 0
'                'BUSCO LA COMPRA MAS ANTIGUA QUE NO SE HAYA VENDIDO COMPLETAMENTE
'                sql = "select * from stock_publicacion where " & _
'                    "spu_vendido < spu_cantid and pub_codigo = " & Rec2!PUB_CODIGO & " and " & _
'                    "spu_feccom = (select min(spu_feccom) as minimo from stock_publicacion where " & _
'                    "              spu_vendido < spu_cantid and pub_codigo = " & Rec2!PUB_CODIGO & ")"
'                Rec3.Open sql, DBConn, adOpenStatic, adLockOptimistic
'                If Rec3.eof = false Then
'                    If canti <= Val(Rec3!spu_cantid) - Val(Rec3!spu_vendido) Then
'                        DBConn.Execute "update stock_publicacion set spu_vendido = spu_vendido + " & canti & " where spu_feccom = '" & Format(Rec3!spu_feccom, "mm/dd/yyyy") & "' " & _
'                         " and pub_codigo = " & Rec2!PUB_CODIGO
'                         Rec3.Close
'                         Exit Do
'                    Else
'                        canti = canti - (Val(Rec3!spu_cantid) - Val(Rec3!spu_vendido))
'                        DBConn.Execute "update stock_publicacion set spu_vendido = spu_cantid where spu_feccom = '" & Format(Rec3!spu_feccom, "mm/dd/yyyy") & "' " & _
'                         " and pub_codigo = " & Rec2!PUB_CODIGO
'                    End If
'                End If
'                Rec3.Close
'             Loop
'       'PARA ANULAR UN RECIBO
'        ElseIf Val(mEstado) = 2 Then
'            DBConn.Execute "update stock_delega set sde_cantid = sde_cantid + " & Rec2!pub_cantidad & " where del_codigo = " & mDelega & _
'             " and pub_codigo = " & Rec2!PUB_CODIGO
'             'va SUMANDO desde la fecha mas NUEVA de compra lo vendido
'             canti = Rec2!pub_cantidad
'             Do While canti > 0
'                'BUSCO LA COMPRA MAS ANTIGUA QUE NO SE HAYA VENDIDO COMPLETAMENTE
'                sql = "select * from stock_publicacion where " & _
'                    "spu_vendido > 0 and pub_codigo = " & Rec2!PUB_CODIGO & " and " & _
'                    "spu_feccom = (select max(spu_feccom) from stock_publicacion where " & _
'                    "              spu_vendido > 0 and pub_codigo = " & Rec2!PUB_CODIGO & ")"
'                Rec3.Open sql, DBConn, adOpenStatic, adLockOptimistic
'                If Rec3.eof = false Then
'                    If canti <= Val(Rec3!spu_vendido) Then
'                        DBConn.Execute "update stock_publicacion set spu_vendido = spu_vendido - " & canti & " where spu_feccom = '" & Format(Rec3!spu_feccom, "mm/dd/yyyy") & "' " & _
'                         " and pub_codigo = " & Rec2!PUB_CODIGO
'                         Rec3.Close
'                         Exit Do
'                    Else
'                        canti = canti - Val(Rec3!spu_vendido)
'                        DBConn.Execute "update stock_publicacion set spu_vendido = 0 where spu_feccom = '" & Format(Rec3!spu_feccom, "mm/dd/yyyy") & "' " & _
'                         " and pub_codigo = " & Rec2!PUB_CODIGO
'                    End If
'                Else
'                    MsgBox "Faltan compras para restar lo devuelto !"
'                End If
'                Rec3.Close
'             Loop
'
'        End If
'    End If
'    Rec2.Close
    
    ConfirmaReciboPUBLICACIONES = True
    Exit Function
    
ErrorTrans:
    ConfirmaReciboPUBLICACIONES = False
End Function
    
    
Public Function ConfirmoReciboConsulta(NroFicticio As String, NroReal As String, Estado As String, Dele As String) As Boolean
    Dim cExec As String
    On Error GoTo MiError
    'ACTUALIZO LAS TABLAS DE CONSULTA
    cExec = " Update CONSULTA Set ERE_CODIGO=" & XN(Estado)
    cExec = cExec & " , REC_NUMERO=" & XN(NroReal)
    cExec = cExec & " Where REC_NUMERO=" & XN(NroFicticio)
    cExec = cExec & "   And DEL_RECIBO=" & XN(Dele)
    DBConn.Execute cExec
    ConfirmoReciboConsulta = True
    Exit Function
    
MiError:
    ConfirmoReciboConsulta = False
End Function

Public Function ConfirmoReciboODONTO(NroFicticio As String, NroReal As String, Estado As Integer, Dele As String) As Boolean
    Set Rec1 = New ADODB.Recordset
    
    ConfirmoReciboODONTO = False
    On Error GoTo ErrorTrans
    
    If Estado = 3 Then 'Recibo Cobrado  'DEFINITIVO
          'Actualizo el estado de LA TABLA REINTEGRO_ODONTO
          'ROD_ESTADO
          cSQL = "Select * " & _
                 " From REINTEGRO_ODONTO " & _
                 " Where REC_NUMERO = " & XN(NroFicticio) & _
                 " AND DEL_CODIGO = " & XN(Dele)

          Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic

          If Rec1.EOF = False Then
             If Rec1!tor_codigo = 8 And (Rec1!SOR_CODIGO = 1 Or Rec1!SOR_CODIGO = 2) Then
             sql = " UPDATE REINTEGRO_ODONTO " & _
                    " SET ROD_ESTADO = " & XS("D") & _
                    " ,REC_NUMERO = " & XN(NroReal) & _
                    " ,REC_DELCOD = " & XN(Dele) & _
                    " Where REC_NUMERO = " & XN(NroFicticio) & _
                    " AND DEL_CODIGO = " & XN(Dele)
             End If
          End If
          Rec1.Close
    ElseIf Estado = 2 Then 'Recibo Anulado
          cSQL = "Select * " & _
                 " From REINTEGRO_ODONTO " & _
                 " Where REC_numero = " & XN(NroReal) & _
                 " AND DEL_CODIGO = " & XN(Dele)
          Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
          If Rec1.EOF = False Then
             If Rec1!tor_codigo = 8 And (Rec1!SOR_CODIGO = 1 Or Rec1!SOR_CODIGO = 2) Then
             sql = " UPDATE REINTEGRO_ODONTO " & _
                    " SET ROD_ESTADO = " & XS("A") & _
                    " Where REC_NUMERO = " & XN(NroReal) & _
                    " AND DEL_CODIGO = " & XN(Dele)
             End If
          End If
          Rec1.Close
    End If
    DBConn.Execute sql
      
    ConfirmoReciboODONTO = True
    Exit Function

ErrorTrans:
    ConfirmoReciboODONTO = False
End Function

Public Function ConfirmoReciboMedico(NroFicticio As String, NroReal As String, Estado As Integer, Dele As String) As Boolean
    ConfirmoReciboMedico = False
'    On Error GoTo Errortrans
'
'      Set Rec1 = New ADODB.Recordset
'
'      If Estado = 3 Then 'Recibo Cobrado  'DEFINITIVO
'          'Actualizo el estado de LA TABLA REINTEGRO_ODONTO
'          'ROD_ESTADO
'          cSQL = "Select * " & _
'                 " From REINTEGRO_ODONTO " & _
'                 " Where REC_NUMERO = " & XN(NroFicticio) & _
'                 " AND DEL_CODIGO = " & XN(Dele)
'
'          Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
'
'          If Rec1.eof = false Then
'             If Rec1!TOR_CODIGO = 8 And (Rec1!SOR_CODIGO = 1 Or Rec1!SOR_CODIGO = 2) Then
'             sql = " UPDATE REINTEGRO_ODONTO " & _
'                    " SET ROD_ESTADO = " & XS("D") & _
'                    " ,REC_NUMERO = " & XN(NroReal) & _
'                    " ,REC_DELCOD = " & XN(Dele) & _
'                    " Where REC_NUMERO = " & XN(NroFicticio) & _
'                    " AND DEL_CODIGO = " & XN(Dele)
'             End If
'          End If
'          Rec1.Close
'      Elseif Estado = 2 Then 'Recibo Anulado
'          cSQL = "Select * " & _
'                 " From REINTEGRO_ODONTO " & _
'                 " Where ROD_NRORECIBO = " & XN(NroReal) & _
'                 " AND DEL_CODIGO = " & XN(Dele)
'
'          Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
'
'          If Rec1.eof = false Then
'             If Rec1!TOR_CODIGO = 8 And (Rec1!SOR_CODIGO = 1 Or Rec1!SOR_CODIGO = 2) Then
'             sql = " UPDATE REINTEGRO_ODONTO " & _
'                    " SET ROD_ESTADO = " & XS("A") & _
'                    " Where REC_NUMERO = " & XN(NroReal) & _
'                    " AND DEL_CODIGO = " & XN(Dele)
'             End If
'          End If
'          Rec1.Close
'      End If
'      DBConn.Execute sql
'      ConfirmoReciboMedico = True
'      Exit Function
'
'Errortrans:
'    ConfirmoReciboMedico = False
End Function

Public Function ConfirmoReciboMEDICAMENTO(NroFicticio As String, NroReal As String, Estado As Integer, Dele As String) As Boolean
   ConfirmoReciboMEDICAMENTO = False
   On Error GoTo ErrorTrans
       
   If Estado = 3 Then 'Recibo Cobrado  'DEFINITIVO
        'Actualizo el estado de LA TABLA REINTEGRO_MEDICA
        'REC_ESTADO
        sql = "Update REINTEGRO_MEDICA set"
        sql = sql & " REC_NUMERO = " & XN(NroReal) & ", "
        sql = sql & " REC_ESTADO = " & Estado
        sql = sql & " where REC_NUMERO = " & XN(NroFicticio)
        sql = sql & " and   REC_DELCOD = " & XN(Dele)
   ElseIf Estado = 2 Then 'Recibo Anulado
        sql = "Update REINTEGRO_MEDICA set"
        sql = sql & " REC_ESTADO = " & Estado
        sql = sql & " where REC_NUMERO = " & XN(NroReal)
        sql = sql & " and   REC_DELCOD = " & XN(Dele)
   End If
   DBConn.Execute sql
   ConfirmoReciboMEDICAMENTO = True
   Exit Function
       
ErrorTrans:
   ConfirmoReciboMEDICAMENTO = False
End Function


Public Sub CargaComboEspecialidadMedico(combo As ComboBox)
    Set RecCombo = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    sql = " SELECT * FROM MEDICO_ESPECIALIDAD "
    combo.Clear
    RecCombo.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If RecCombo.EOF = False Then
       Do While Not RecCombo.EOF
          combo.AddItem CompletarConEspacios(RecCombo.Fields(1), 100) & RecCombo.Fields(0)
          RecCombo.MoveNext
       Loop
    Else
        combo.AddItem ("Sin Datos")
    End If
    RecCombo.Close
    Set RecCombo = Nothing
    Screen.MousePointer = vbNormal
End Sub

Public Sub SelecTexto(TEXTO As Control)
    'Marca el texto de un TextBox como seleccionado (pintado)
    TEXTO.SelStart = 0
    TEXTO.SelLength = Len(TEXTO)
End Sub

Public Sub CargaComboCategoriaMedico(combo As ComboBox)

    Set RecCombo = New ADODB.Recordset
    sql = "SELECT * FROM MEDICO_CATEGORIA"
    combo.Clear
    RecCombo.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If RecCombo.EOF = False Then
        Do While Not RecCombo.EOF
            combo.AddItem CompletarConEspacios(RecCombo.Fields(1), 100) & RecCombo.Fields(0)
            RecCombo.MoveNext
        Loop
    End If
    RecCombo.Close
    Set RecCombo = Nothing

End Sub

Public Function Limpiar_Puntos(TEXTO As String) As String
Dim a As Integer
    TEXTO = Trim(TEXTO)

    For a = 1 To Len(Trim(TEXTO))
        If Mid(TEXTO, a, 1) = "." Then
            TEXTO = Mid(TEXTO, 1, a - 1) & Mid(TEXTO, a + 1, Len(TEXTO))
            a = a - 1
        End If
    Next
    
    Limpiar_Puntos = TEXTO
End Function

Public Function RecuperoDescDelega(CodDel As String) As String
    Dim DescDelega As String
    Dim sDel As String
    Dim rDel As ADODB.Recordset
    Set rDel = New ADODB.Recordset
    sDel = " Select DEL_DESCRI From DELEGACION "
    sDel = sDel & " Where DEL_CODIGO=" & XN(CodDel)
    rDel.Open sDel, DBConn, adOpenStatic, adLockOptimistic
    If rDel.EOF = False Then
        'rDel.MoveFirst
        DescDelega = Trim(ChkNull(rDel!DEL_DESCRI))
    Else
        DescDelega = ""
    End If
    rDel.Close
    RecuperoDescDelega = DescDelega
End Function

Public Function KN(Parametro As Double) As String
    Dim Valor As String
    Dim a As Integer, TEXTO As String, cad1 As String
    Valor = Str(Parametro)
    If Trim(Valor) = "" Or Trim(Valor) = "Null" Then
        KN = "Null"
    Else
        KN = Trim(Valor)
        For a = 1 To Len(KN)
            If Mid(KN, a, 1) = "," Then
                TEXTO = TEXTO & "."
            ElseIf Mid(KN, a, 1) = "." Then
                TEXTO = TEXTO
            Else
                TEXTO = TEXTO & Mid(KN, a, 1)
            End If
        Next
        KN = TEXTO
    End If
End Function

Public Sub ConfiguroFrmClave(TitPrincipal As String, TitDos As String, TitTres As String, DescUsuario As String, USUARIO As String)
    frmClave.LblClave.Caption = TitPrincipal
    frmClave.LblTitulo1.Caption = TitDos
    frmClave.LblTitulo2.Caption = TitTres
    frmClave.LblUsuario.Caption = DescUsuario
    frmClave.TxtUsuario.Text = USUARIO
End Sub

Public Function ANCHO_CAMPO(Campo As String, ancho As Integer) As String
    If ancho > Len(Trim(Campo)) Then
        ANCHO_CAMPO = Trim(Campo) & Space(ancho - Len(Trim(Campo)))
    Else
        ANCHO_CAMPO = Trim(Mid(Trim(Campo), 1, ancho))
    End If
    
    If Len(ANCHO_CAMPO) < ancho Then ANCHO_CAMPO = ANCHO_CAMPO & Space(ancho - Len(ANCHO_CAMPO))
End Function

Public Function ConfirmoReciboTecnicaVta(NroFicticio As String, NroReal As String, Estado As String, Dele As String, Fecha As String) As Boolean

    ConfirmoReciboTecnicaVta = False
    On Error GoTo MiError
    
    'ACTUALIZO LAS TABLAS DE RECIBO EN TECNICA
    cExec = " Update RECIBO_TECNICA Set ERE_CODIGO=" & XN(Estado)
    cExec = cExec & " , REC_NUMERO=" & XN(NroReal)
    cExec = cExec & " Where REC_NUMERO=" & XN(NroFicticio)
    cExec = cExec & "   And DEL_RECIBO=" & XN(Dele)
    DBConn.Execute cExec
    cExec = " Update DETALLE_RECIBO_TECNICA Set REC_NUMERO=" & XN(NroReal)
    cExec = cExec & " Where REC_NUMERO=" & XN(NroFicticio)
    cExec = cExec & "   And DEL_RECIBO=" & XN(Dele)
    DBConn.Execute cExec
    If Estado = "3" Then
        cExec = " Update FORMULARIOS Set DEL_RECIBO = " & XN(Dele) & "  , REC_NUMERO = " & XN(NroReal) & ", FEC_RECIBO = " & XD(Fecha)
        cExec = cExec & " Where DEL_RECIBO=" & XN(Dele)
        cexe = cExec & "    And REC_NUMERO=" & XN(NroFicticio)
        DBConn.Execute cExec
    ElseIf Estado = "2" Then
        cExec = " Update FORMULARIOS Set DEL_RECIBO = NULL  , REC_NUMERO = NULL , FEC_RECIBO = NULL"
        cExec = cExec & " Where DEL_RECIBO=" & XN(Dele)
        cexe = cExec & "    And REC_NUMERO=" & XN(NroFicticio)
        DBConn.Execute cExec
    End If
    ConfirmoReciboTecnicaVta = True
    Exit Function
    
MiError:
    ConfirmoReciboTecnicaVta = False
End Function

Public Function ServicioMedico(TipoDoc As String, Nrodoc As String, Modulo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------'
    sql = "SELECT Persona.Per_Apelli, Persona.Per_Nombre, Persona.Tip_TipDoc, Persona.Per_nrodoc,'TITULAR' as PAR_DESCRI,servicio_medico.pla_codigo, Domicilio.Dom_Calle, Localidad.Loc_Descri, Titulo.Tit_Matric, Persona.Per_FecNac, "
    sql = sql & " '00' as PAR_CODIGO, '00' as FAM_ORDEN, SEX_CODIGO "
    sql = sql & " FROM Persona, Servicio_Medico, Plan_Cobertura, Domicilio, Localidad, Titulo "
    sql = sql & " WHERE persona.tip_tipdoc = " & XS(TipoDoc) & " "
    sql = sql & " AND persona.per_nrodoc=" & XS(Nrodoc) & " "
    sql = sql & " AND titulo.tip_tipdoc = persona.tip_tipdoc "
    sql = sql & " AND titulo.per_nrodoc = persona.per_nrodoc "
    sql = sql & " AND ( titulo.car_codigo = 1 OR titulo.car_codigo = 5 OR titulo.car_codigo = 7 OR titulo.car_codigo = 10 OR titulo.car_codigo = 13 OR titulo.car_codigo = 14 OR titulo.car_codigo = 11 ) "
    sql = sql & " AND persona.tip_tipdoc = domicilio.tip_tipdoc "
    sql = sql & " AND persona.per_nrodoc = domicilio.per_nrodoc "
    sql = sql & " AND Domicilio.Dom_Corres <> 'NO' "
    sql = sql & " AND Domicilio.Loc_Codigo = Localidad.Loc_Codigo "
    sql = sql & " AND Domicilio.Pro_Codigo = Localidad.Pro_Codigo "
    sql = sql & " AND Domicilio.Pai_Codigo = Localidad.Pai_Codigo "
    sql = sql & " AND persona.tip_tipdoc = servicio_medico.tip_tipdoc "
    sql = sql & " AND persona.per_nrodoc = servicio_medico.per_nrodoc "
    sql = sql & " AND Servicio_Medico.Mot_Codigo Is Null "
    sql = sql & " AND servicio_medico.ser_titopt = 'T' "
    sql = sql & " AND Servicio_Medico.Pla_Codigo = Plan_Cobertura.Pla_Codigo "
    sql = sql & " AND (SELECT COUNT(*) FROM MODULO_PLAN "
    sql = sql & " WHERE Plan_Cobertura.Pla_Codigo = Modulo_Plan.Pla_Codigo "
    sql = sql & " AND Modulo_Plan.Mod_Codigo = " & Modulo & ") > 0 "
    sql = sql & " AND (SELECT count(*) from servicio_medico_carencia "
    sql = sql & " WHERE servicio_medico.tip_tipdoc = servicio_medico_carencia.tip_tipdoc "
    sql = sql & " AND servicio_medico.per_nrodoc = servicio_medico_carencia.per_nrodoc "
    sql = sql & " AND servicio_medico.ser_fecafi = servicio_medico_carencia.ser_fecafi "
    sql = sql & " AND servicio_medico.pla_codigo = servicio_medico_carencia.pla_codigo "
    sql = sql & " AND servicio_medico_carencia.mod_codigo = " & Modulo
    sql = sql & " AND Servicio_Medico_carencia.Sec_Feccar >= '" & Format(Date, "mm/dd/yyyy") & "') = 0"
    
    sql = sql & " Union All "
    sql = sql & " Select persona.per_apelli, persona.per_nombre, persona.tip_tipdoc, persona.per_nrodoc, parentesco.par_descri, servicio_medico.pla_codigo, 'DOMICILIO','LOCALIDAD','MATRICULA', Persona.Per_FecNac, "
    sql = sql & " Familiar.PAR_CODIGO, Familiar.FAM_ORDEN, SEX_CODIGO "
    sql = sql & " FROM Persona, familiar, parentesco, servicio_medico, Plan_Cobertura "
    sql = sql & " WHERE familiar.mat_tip_tipdoc = " & XS(TipoDoc)
    sql = sql & " AND familiar.mat_per_nrodoc=" & XS(Nrodoc)
    sql = sql & " AND familiar.tip_tipdoc = persona.tip_tipdoc "
    sql = sql & " AND familiar.per_nrodoc = persona.per_nrodoc "
    sql = sql & " AND familiar.par_codigo = parentesco.par_codigo "
    sql = sql & " AND servicio_medico.ser_titopt = 'O' "
    sql = sql & " AND servicio_medico.tip_tipdoc = familiar.tip_tipdoc "
    sql = sql & " AND servicio_medico.per_nrodoc = familiar.per_nrodoc "
    sql = sql & " AND Servicio_Medico.Mot_Codigo Is Null "
    sql = sql & " AND Servicio_Medico.Pla_Codigo = Plan_Cobertura.Pla_Codigo "
    sql = sql & " AND (SELECT COUNT(*) FROM MODULO_PLAN "
    sql = sql & " WHERE Plan_Cobertura.Pla_Codigo = Modulo_Plan.Pla_Codigo "
    sql = sql & " AND Modulo_Plan.Mod_Codigo = " & Modulo & ") > 0 "
    sql = sql & " AND (SELECT COUNT(*) from servicio_medico_carencia "
    sql = sql & " WHERE servicio_medico.tip_tipdoc = servicio_medico_carencia.tip_tipdoc "
    sql = sql & " AND servicio_medico.per_nrodoc = servicio_medico_carencia.per_nrodoc "
    sql = sql & " AND servicio_medico.ser_fecafi = servicio_medico_carencia.ser_fecafi "
    sql = sql & " AND servicio_medico.pla_codigo = servicio_medico_carencia.pla_codigo "
    sql = sql & " AND servicio_medico_carencia.mod_codigo = " & Modulo
    sql = sql & " AND Servicio_Medico_carencia.Sec_Feccar >= '" & Format(Date, "mm/dd/yyyy") & "') = 0"
    sql = sql & " ORDER BY PAR_CODIGO, FAM_ORDEN"
    '---------------------------------------------------------------------------------------------------------------------------------------------'

    ServicioMedico = sql
    
End Function

Public Sub LISTADO_VIGENTES_SEGURO(TIPO_SEGURO As Integer, Fecha As Date)
       
    Set rec = New ADODB.Recordset
    
    Screen.MousePointer = 11
    
    On Error GoTo CLAVOSE
    
    If TIPO_SEGURO = 1 Then
        
        sql = "SELECT * FROM SYSOBJECTS WHERE NAME = 'LISTADO_SEGURO_VIDA'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then DBConn.Execute "DROP VIEW dbo.LISTADO_SEGURO_VIDA"
        rec.Close
        
        sql = "SELECT * FROM SYSOBJECTS WHERE NAME = 'LISTADO_SEGURO_VIDA_TOTAL'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then DBConn.Execute "DROP VIEW dbo.LISTADO_SEGURO_VIDA_TOTAL"
        rec.Close
        
        'Agrego los titulares con SEGURO ACTIVO
        sql = "CREATE VIEW dbo.LISTADO_SEGURO_VIDA AS " & _
        "SELECT  RTRIM(CONVERT(CHAR,TTI_CODIGO)) + '.' + TIT_MATRIC + '.' + TIT_DIGVER AS TITULO, PER_APELLI, PER_NOMBRE ,SEG_NROCER,SEG_CAPITA,SEG_PRIMA,SEG_TITOPT,'' AS PARENTESCO, P.TIP_TIPDOC,P.PER_NRODOC " & _
        "FROM seguro S, titulo T, persona P WHERE " & _
        "S.TIP_TIPDOC = P.TIP_TIPDOC AND S.PER_NRODOC = P.PER_NRODOC AND " & _
        "S.TIP_TIPDOC = T.TIP_TIPDOC AND S.PER_NRODOC = T.PER_NRODOC AND " & _
        "TIP_CODIGO = 1 AND TTI_CODIGO < 21 AND SEG_CAPITA > 0 AND " & _
        "CAR_CODIGO <> 3 AND CAR_CODIGO <> 4 AND S.SEG_TITOPT = 'T' AND " & _
        "SEG_FECAFI <= '" & Format(Fecha, "MM/DD/YYYY") & "' AND (SEG_FECDES IS NULL OR SEG_FECDES > '" & Format(Fecha, "MM/DD/YYYY") & "') "
       
        'Agrego los titulares FALLECIDOS con conyugue activa
        sql = sql & " UNION ALL " & _
        "SELECT  RTRIM(CONVERT(CHAR,TTI_CODIGO)) + '.' + TIT_MATRIC + '.' + TIT_DIGVER AS TITULO, PER_APELLI, PER_NOMBRE ,SEG_NROCER, 0 AS SEG_CAPITA, 0 AS SEG_PRIMA,SEG_TITOPT,'' AS PARENTESCO, P.TIP_TIPDOC,P.PER_NRODOC  " & _
        "FROM SEGURO S,TITULO T, FAMILIAR F, PERSONA P WHERE " & _
        "TIP_CODIGO = 1 And S.TIP_TIPDOC = T.TIP_TIPDOC " & _
        "AND S.PER_NRODOC = T.PER_NRODOC " & _
        "AND CAR_CODIGO = 4 AND TTI_CODIGO < 21 " & _
        "AND S.TIP_TIPDOC = F.MAT_TIP_TIPDOC " & _
        "AND S.PER_NRODOC = F.MAT_PER_NRODOC " & _
        "AND F.PAR_CODIGO = 1 " & _
        "AND S.TIP_TIPDOC = P.TIP_TIPDOC " & _
        "AND S.PER_NRODOC = P.PER_NRODOC " & _
        "AND (SELECT COUNT(*) FROM SEGURO WHERE " & _
        "      F.TIP_TIPDOC = TIP_TIPDOC AND " & _
        "      F.PER_NRODOC = PER_NRODOC AND " & _
        "      TIP_CODIGO = 1 AND SEG_TITOPT = 'O' AND " & _
        "      SEG_FECAFI <= '" & Format(Fecha, "MM/DD/YYYY") & "' AND " & _
        "      (SEG_FECDES IS NULL OR SEG_FECDES > '" & Format(Fecha, "MM/DD/YYYY") & "')) > 0 "
        
        'AND CAR_CODIGO <> 2
        'Agrego los conyugues activos SOLO LOS CONYUGUES
        sql = sql & "Union All " & _
        "SELECT  RTRIM(CONVERT(CHAR,TTI_CODIGO)) + '.' + TIT_MATRIC + '.' + TIT_DIGVER AS TITULO, PER_APELLI, PER_NOMBRE ,SEG_NROCER,SEG_CAPITA,SEG_PRIMA,SEG_TITOPT,'CONYUGUE' AS PARENTESCO, P.TIP_TIPDOC,P.PER_NRODOC  " & _
        "FROM SEGURO S,TITULO T, FAMILIAR F, PERSONA P WHERE " & _
        "      TIP_CODIGO = 1 " & _
        "AND S.TIP_TIPDOC = F.TIP_TIPDOC " & _
        "AND S.PER_NRODOC = F.PER_NRODOC " & _
        "AND S.SEG_CAPITA > 0 AND SEG_TITOPT = 'O'" & _
        "AND F.TIP_TIPDOC = P.TIP_TIPDOC " & _
        "AND F.PER_NRODOC = P.PER_NRODOC " & _
        "AND F.PAR_CODIGO = 1 " & _
        "AND T.TIP_TIPDOC = F.MAT_TIP_TIPDOC " & _
        "AND T.PER_NRODOC = F.MAT_PER_NRODOC " & _
        "AND CAR_CODIGO <> 3  " & _
        "AND SEG_FECAFI <= '" & Format(Fecha, "MM/DD/YYYY") & "' " & _
        "AND (SEG_FECDES IS NULL OR SEG_FECDES > '" & Format(Fecha, "MM/DD/YYYY") & "') "
        DBConn.Execute sql
        
    Else
    
        sql = "SELECT * FROM SYSOBJECTS WHERE NAME = 'LISTADO_SEGURO_SEPELIO'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then DBConn.Execute "DROP VIEW dbo.LISTADO_SEGURO_SEPELIO"
        rec.Close
        
        sql = "SELECT * FROM SYSOBJECTS WHERE NAME = 'LISTADO_SEGURO_SEPELIO_TOTAL'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then DBConn.Execute "DROP VIEW dbo.LISTADO_SEGURO_SEPELIO_TOTAL"
        rec.Close
        
        'Agrego los titulares AFILIADOS al seguro
        sql = "CREATE VIEW dbo.LISTADO_SEGURO_SEPELIO AS " & _
        "SELECT  RTRIM(CONVERT(CHAR,TTI_CODIGO)) + '.' + TIT_MATRIC + '.' + TIT_DIGVER AS TITULO, PER_APELLI, PER_NOMBRE ,SEG_NROCER,SEG_CAPITA,SEG_PRIMA,'' AS PARENTESCO, P.TIP_TIPDOC,P.PER_NRODOC,CAR_DESCRI,P.PER_FECNAC  " & _
        "FROM seguro S, titulo T, persona P, CARACTER C WHERE " & _
        "S.TIP_TIPDOC = P.TIP_TIPDOC AND S.PER_NRODOC = P.PER_NRODOC AND " & _
        "S.TIP_TIPDOC = T.TIP_TIPDOC AND S.PER_NRODOC = T.PER_NRODOC AND " & _
        "T.CAR_CODIGO = C.CAR_CODIGO AND " & _
        "TIP_CODIGO = 2 AND TTI_CODIGO < 21 AND " & _
        "T.CAR_CODIGO <> 3 AND T.CAR_CODIGO <> 4 AND S.SEG_TITOPT = 'T' AND " & _
        "SEG_FECAFI <= '" & Format(Fecha, "MM/DD/YYYY") & "' AND " & _
        "(SEG_FECDES IS NULL OR SEG_FECDES > '" & Format(Fecha, "MM/DD/YYYY") & "') "

        'Agrego los titulares FALLECIDOS con familiares activos
        sql = sql & " UNION ALL " & _
        "SELECT DISTINCT RTRIM(CONVERT(CHAR,TTI_CODIGO)) + '.' + TIT_MATRIC + '.' + TIT_DIGVER AS TITULO, PER_APELLI, PER_NOMBRE ,SEG_NROCER, 0 AS SEG_CAPITA, 0 AS SEG_PRIMA,'' AS PARENTESCO, P.TIP_TIPDOC,P.PER_NRODOC, CAR_DESCRI,P.PER_FECNAC  " & _
        "FROM SEGURO S,TITULO T, PERSONA P, CARACTER C WHERE " & _
        "     TIP_CODIGO = 2 " & _
        "And S.TIP_TIPDOC = T.TIP_TIPDOC AND S.PER_NRODOC = T.PER_NRODOC " & _
        "And S.TIP_TIPDOC = P.TIP_TIPDOC AND S.PER_NRODOC = P.PER_NRODOC " & _
        "AND T.CAR_CODIGO = C.CAR_CODIGO " & _
        "AND T.CAR_CODIGO = 4 AND TTI_CODIGO < 21 " & _
        "AND (SELECT COUNT(*) FROM SEGURO SS, FAMILIAR FF WHERE " & _
        "    FF.MAT_TIP_TIPDOC = T.TIP_TIPDOC AND " & _
        "    FF.MAT_PER_NRODOC = T.PER_NRODOC AND " & _
        "    FF.TIP_TIPDOC = SS.TIP_TIPDOC AND " & _
        "    FF.PER_NRODOC = SS.PER_NRODOC AND " & _
        "    SS.TIP_CODIGO = 2 AND SS.SEG_TITOPT = 'O' AND " & _
        "    SS.SEG_FECAFI <= '08/15/2000' AND " & _
        "    (SS.SEG_FECDES IS NULL OR SS.SEG_FECDES > '08/15/2000')) > 0 "
        
        'Agrego los familiares activos TODOS LOS FAMILIARES
        sql = sql & "Union All " & _
        "SELECT  RTRIM(CONVERT(CHAR,TTI_CODIGO)) + '.' + TIT_MATRIC + '.' + TIT_DIGVER AS TITULO, PER_APELLI, PER_NOMBRE ,SEG_NROCER,SEG_CAPITA,SEG_PRIMA,PAR_DESCRI AS PARENTESCO, P.TIP_TIPDOC,P.PER_NRODOC,'',P.PER_FECNAC " & _
        "FROM SEGURO S,TITULO T, FAMILIAR F, PERSONA P, PARENTESCO PA, CARACTER C WHERE " & _
        "      TIP_CODIGO = 2 " & _
        "AND S.TIP_TIPDOC = F.TIP_TIPDOC " & _
        "AND S.PER_NRODOC = F.PER_NRODOC " & _
        "AND F.PAR_CODIGO = PA.PAR_CODIGO " & _
        "AND SEG_TITOPT = 'O'" & _
        "AND F.TIP_TIPDOC = P.TIP_TIPDOC " & _
        "AND F.PER_NRODOC = P.PER_NRODOC " & _
        "AND T.TIP_TIPDOC = F.MAT_TIP_TIPDOC " & _
        "AND T.PER_NRODOC = F.MAT_PER_NRODOC " & _
        "AND T.CAR_CODIGO = C.CAR_CODIGO " & _
        "AND T.CAR_CODIGO <> 3 AND T.CAR_CODIGO <> 2 " & _
        "AND SEG_FECAFI <= '" & Format(Fecha, "MM/DD/YYYY") & "' " & _
        "AND (SEG_FECDES IS NULL OR SEG_FECDES > '" & Format(Fecha, "MM/DD/YYYY") & "') "
        
        DBConn.Execute sql
         
    
    End If
    
    Screen.MousePointer = 1
    
    Exit Sub
    
    
CLAVOSE:

    If rec.State = 1 Then rec.Close
    
    sql = "SELECT * FROM SYSOBJECTS WHERE NAME = 'LISTADO_SEGURO_VIDA'"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then DBConn.Execute "DROP VIEW dbo.LISTADO_SEGURO_VIDA"
    rec.Close

    sql = "SELECT * FROM SYSOBJECTS WHERE NAME = 'LISTADO_SEGURO_SEPELIO'"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then DBConn.Execute "DROP VIEW LISTADO_SEGURO_SEPELIO"
    rec.Close
    
    sql = "SELECT * FROM SYSOBJECTS WHERE NAME = 'LISTADO_SEGURO_SEPELIO_TOTAL'"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then DBConn.Execute "DROP VIEW LISTADO_SEGURO_SEPELIO_TOTAL"
    rec.Close
    
    Mensaje 4
    Screen.MousePointer = 1
End Sub


Public Sub CompararVersiones(origen As String, destino As String)
    'Autor: Gabriel Allasia
    On Error Resume Next
    If FileDateTime(origen) > FileDateTime(destino) Then
        MsgBox "Existe una nueva versi�n del Sistema del d�a " & Format(FileDateTime(origen), "dd/mm/yyyy") & ", se recomienda que" & Chr(13) & "abandone el Sistema y lo actualize." & Chr(13) & "Si Usted no tiene permiso para actualizarlo comun�quese con Soporte de Sistemas.", vbExclamation
    End If
End Sub

Public Sub InserPuntoNomenclador(TEXTO As Control)
 Dim a As String
 Dim B As String
 Dim C As String
 ' Agrega punto al codigo del nomenclador
  If Len(TEXTO) = 6 Then
   TEXTO.SelStart = 0
   TEXTO.SelLength = 2
   a = TEXTO.SelText
   TEXTO.SelStart = 2
   TEXTO.SelLength = 2
   B = TEXTO.SelText
   TEXTO.SelStart = 4
   TEXTO.SelLength = 2
   C = TEXTO.SelText
   TEXTO.Text = a + "." + B + "." + C
  End If
End Sub

Public Sub ActualizaEstado(vNroPre As String, vTipPre As String, vEstado As String, cFecha As Date)
    Dim RecAux As New ADODB.Recordset, Actualiza As Boolean
    Actualiza = False
    sql = "SELECT * FROM EstViejo e " + _
    " WHERE e.tpr_codtippr=" + XS(vTipPre) + _
    " AND   e.pre_nroprest=" + XS(vNroPre) + _
    " AND   e.est_fecestad IN (SELECT MAX(x.est_fecestad) " + _
    " FROM  EstViejo x " + _
    " WHERE e.tpr_codtippr=x.tpr_codtippr " + _
    " AND   e.pre_nroprest=x.pre_nroprest) " + _
    " ORDER BY e.est_fecestad "
    RecAux.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If RecAux.EOF = False Then
        If RecAux!est_estprest <> vEstado Then Actualiza = True
    Else
        Actualiza = True
    End If
    RecAux.Close
    If Actualiza = True Then
        sql = "INSERT INTO EstViejo (tpr_codtippr,pre_nroprest,est_estprest,est_fecestad) VALUES( " + _
        XS(vTipPre) + "," + _
        XS(vNroPre) + "," + _
        XS(vEstado) + "," + _
        XD(cFecha) + ")"
        DBConn.Execute sql
    End If
End Sub

Sub justifica_printer(x0, xf, y0, txt)
' x0, xf = posicion de los margenes izquierdo y derecho
' y0 = posicion vertical donde se desea empezar a escribir
' txt = texto a escribir

Dim X, Y, k, ancho
Dim s As String, ss As String
Dim x_spc

s = txt
X = x0
Y = y0
ancho = (xf - x0)

While s <> ""

  ss = ""
  While (s <> "") And (Printer.TextWidth(ss) <= ancho)
    ss = ss & Left$(s, 1)
    s = Right$(s, Len(s) - 1)
  Wend
  If (Printer.TextWidth(ss) > ancho) Then
    s = Right$(ss, 1) & s
    ss = Left$(ss, Len(ss) - 1)
  End If
  ' aqui tenemos en ss lo maximo que cabe en una linea
  If Right$(ss, 1) = " " Then
     ss = Left$(ss, Len(ss) - 1)
  Else
     If (InStr(ss, " ") > 0) And (Left$(s & " ", 1) <> " ") Then
      While Right$(ss, 1) <> " "
        s = Right$(ss, 1) & s
        ss = Left$(ss, Len(ss) - 1)
      Wend
      ss = Left$(ss, Len(ss) - 1)
     End If
  End If
  x_spc = 0
  X = x0
  If (Len(ss) > 1) And (s & "" <> "") Then
    x_spc = (ancho - Printer.TextWidth(ss)) / (Len(ss) - 1)
  End If
  Printer.CurrentX = X
  Printer.CurrentY = Y

  If x_spc = 0 Then
    Printer.Print ss;
  Else
    For k = 1 To Len(ss)
     Printer.CurrentX = X
     Printer.Print Mid$(ss, k, 1);
     X = X + Printer.TextWidth("*" & Mid$(ss, k, 1) & "*") - Printer.TextWidth("**")
     X = X + x_spc
    Next
  End If

  Y = Y + Printer.TextHeight(ss)
  While Left$(s, 1) = " "
    s = Right$(s, Len(s) - 1)
  Wend
Wend

End Sub

 
'Public Function EnLetras(Numero As String) As String
'    Dim B, paso As Integer
'    Dim expresion, entero, deci, flag As String
'
'    flag = "N"
'    For paso = 1 To Len(Numero)
'        If Mid(Numero, paso, 1) = "," Then
'            flag = "S"
'        Else
'            If flag = "N" Then
'                entero = entero + Mid(Numero, paso, 1) 'Extae la parte entera del numero
'            Else
'                deci = deci + Mid(Numero, paso, 1) 'Extrae la parte decimal del numero
'            End If
'        End If
'    Next paso
'
'    If Len(deci) = 1 Then
'        deci = deci & "0"
'    End If
'
'    flag = "N"
'    If Val(Numero) >= -999999999 And Val(Numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
'        For paso = Len(entero) To 1 Step -1
'            B = Len(entero) - (paso - 1)
'            Select Case paso
'            Case 3, 6, 9
'                Select Case Mid(entero, B, 1)
'                    Case "1"
'                        If Mid(entero, B + 1, 1) = "0" And Mid(entero, B + 2, 1) = "0" Then
'                            expresion = expresion & "cien "
'                        Else
'                            expresion = expresion & "ciento "
'                        End If
'                    Case "2"
'                        expresion = expresion & "doscientos "
'                    Case "3"
'                        expresion = expresion & "trescientos "
'                    Case "4"
'                        expresion = expresion & "cuatrocientos "
'                    Case "5"
'                        expresion = expresion & "quinientos "
'                    Case "6"
'                        expresion = expresion & "seiscientos "
'                    Case "7"
'                        expresion = expresion & "setecientos "
'                    Case "8"
'                        expresion = expresion & "ochocientos "
'                    Case "9"
'                        expresion = expresion & "novecientos "
'                End Select
'
'            Case 2, 5, 8
'                Select Case Mid(entero, B, 1)
'                    Case "1"
'                        If Mid(entero, B + 1, 1) = "0" Then
'                            flag = "S"
'                            expresion = expresion & "diez "
'                        End If
'                        If Mid(entero, B + 1, 1) = "1" Then
'                            flag = "S"
'                            expresion = expresion & "once "
'                        End If
'                        If Mid(entero, B + 1, 1) = "2" Then
'                            flag = "S"
'                            expresion = expresion & "doce "
'                        End If
'                        If Mid(entero, B + 1, 1) = "3" Then
'                            flag = "S"
'                            expresion = expresion & "trece "
'                        End If
'                        If Mid(entero, B + 1, 1) = "4" Then
'                            flag = "S"
'                            expresion = expresion & "catorce "
'                        End If
'                        If Mid(entero, B + 1, 1) = "5" Then
'                            flag = "S"
'                            expresion = expresion & "quince "
'                        End If
'                        If Mid(entero, B + 1, 1) > "5" Then
'                            flag = "N"
'                            expresion = expresion & "dieci"
'                        End If
'
'                    Case "2"
'                        If Mid(entero, B + 1, 1) = "0" Then
'                            expresion = expresion & "veinte "
'                            flag = "S"
'                        Else
'                            expresion = expresion & "veinti"
'                            flag = "N"
'                        End If
'
'                    Case "3"
'                        If Mid(entero, B + 1, 1) = "0" Then
'                            expresion = expresion & "treinta "
'                            flag = "S"
'                        Else
'                            expresion = expresion & "treinta y "
'                            flag = "N"
'                        End If
'
'                    Case "4"
'                        If Mid(entero, B + 1, 1) = "0" Then
'                            expresion = expresion & "cuarenta "
'                            flag = "S"
'                        Else
'                            expresion = expresion & "cuarenta y "
'                            flag = "N"
'                        End If
'
'                    Case "5"
'                        If Mid(entero, B + 1, 1) = "0" Then
'                            expresion = expresion & "cincuenta "
'                            flag = "S"
'                        Else
'                            expresion = expresion & "cincuenta y "
'                            flag = "N"
'                        End If
'
'                    Case "6"
'                        If Mid(entero, B + 1, 1) = "0" Then
'                            expresion = expresion & "sesenta "
'                            flag = "S"
'                        Else
'                            expresion = expresion & "sesenta y "
'                            flag = "N"
'                        End If
'
'                    Case "7"
'                        If Mid(entero, B + 1, 1) = "0" Then
'                            expresion = expresion & "setenta "
'                            flag = "S"
'                        Else
'                            expresion = expresion & "setenta y "
'                            flag = "N"
'                        End If
'
'                    Case "8"
'                        If Mid(entero, B + 1, 1) = "0" Then
'                            expresion = expresion & "ochenta "
'                            flag = "S"
'                        Else
'                            expresion = expresion & "ochenta y "
'                            flag = "N"
'                        End If
'
'                    Case "9"
'                        If Mid(entero, B + 1, 1) = "0" Then
'                            expresion = expresion & "noventa "
'                            flag = "S"
'                        Else
'                            expresion = expresion & "noventa y "
'                            flag = "N"
'                        End If
'                End Select
'
'            Case 1, 4, 7
'                Select Case Mid(entero, B, 1)
'                    Case "1"
'                        If flag = "N" Then
'                            If paso = 1 Then
'                                expresion = expresion & "uno "
'                            Else
'                                expresion = expresion & "un "
'                            End If
'                        End If
'                    Case "2"
'                        If flag = "N" Then
'                            expresion = expresion & "dos "
'                        End If
'                    Case "3"
'                        If flag = "N" Then
'                            expresion = expresion & "tres "
'                        End If
'                    Case "4"
'                        If flag = "N" Then
'                            expresion = expresion & "cuatro "
'                        End If
'                    Case "5"
'                        If flag = "N" Then
'                            expresion = expresion & "cinco "
'                        End If
'                    Case "6"
'                        If flag = "N" Then
'                            expresion = expresion & "seis "
'                        End If
'                    Case "7"
'                        If flag = "N" Then
'                            expresion = expresion & "siete "
'                        End If
'                    Case "8"
'                        If flag = "N" Then
'                            expresion = expresion & "ocho "
'                        End If
'                    Case "9"
'                        If flag = "N" Then
'                            expresion = expresion & "nueve "
'                        End If
'                End Select
'            End Select
'            If paso = 4 Then
'                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
'                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
'                   Len(entero) <= 6) Then
'                    expresion = expresion & "mil "
'                End If
'            End If
'            If paso = 7 Then
'                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
'                    expresion = expresion & "mill�n "
'                Else
'                    expresion = expresion & "millones "
'                End If
'            End If
'        Next paso
'
'        If deci <> "" Then
'            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
'                EnLetras = "menos " & expresion & "con " & deci ' & "/100"
'            Else
'                EnLetras = expresion & "con " & deci ' & "/100"
'            End If
'        Else
'            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
'                EnLetras = "menos " & expresion
'            Else
'                EnLetras = expresion
'            End If
'        End If
'    Else 'si el numero a convertir esta fuera del rango superior e inferior
'        EnLetras = ""
'    End If
'End Function

Public Function EnLetras(Numero As String) As String
    Dim B, paso As Integer
    Dim expresion, entero, deci, flag As String
       
    flag = "N"
    For paso = 1 To Len(Numero)
        If Mid(Numero, paso, 1) = "," Then
            flag = "S"
        Else
            If flag = "N" Then
                entero = entero + Mid(Numero, paso, 1) 'Extae la parte entera del numero
            Else
                deci = deci + Mid(Numero, paso, 1) 'Extrae la parte decimal del numero
            End If
        End If
    Next paso
   
    If Len(deci) = 1 Then
        deci = deci & "0"
    End If
   
    entero = Replace(entero, ".", "")
    flag = "N"
    If Val(Numero) >= -999999999 And Val(Numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
        For paso = Len(entero) To 1 Step -1
            B = Len(entero) - (paso - 1)
            Select Case paso
            Case 3, 6, 9
                Select Case Mid(entero, B, 1)
                    Case "1"
                        If Mid(entero, B + 1, 1) = "0" And Mid(entero, B + 2, 1) = "0" Then
                            expresion = expresion & "cien "
                        Else
                            expresion = expresion & "ciento "
                        End If
                    Case "2"
                        expresion = expresion & "doscientos "
                    Case "3"
                        expresion = expresion & "trescientos "
                    Case "4"
                        expresion = expresion & "cuatrocientos "
                    Case "5"
                        expresion = expresion & "quinientos "
                    Case "6"
                        expresion = expresion & "seiscientos "
                    Case "7"
                        expresion = expresion & "setecientos "
                    Case "8"
                        expresion = expresion & "ochocientos "
                    Case "9"
                        expresion = expresion & "novecientos "
                End Select
               
            Case 2, 5, 8
                Select Case Mid(entero, B, 1)
                    Case "1"
                        If Mid(entero, B + 1, 1) = "0" Then
                            flag = "S"
                            expresion = expresion & "diez "
                        End If
                        If Mid(entero, B + 1, 1) = "1" Then
                            flag = "S"
                            expresion = expresion & "once "
                        End If
                        If Mid(entero, B + 1, 1) = "2" Then
                            flag = "S"
                            expresion = expresion & "doce "
                        End If
                        If Mid(entero, B + 1, 1) = "3" Then
                            flag = "S"
                            expresion = expresion & "trece "
                        End If
                        If Mid(entero, B + 1, 1) = "4" Then
                            flag = "S"
                            expresion = expresion & "catorce "
                        End If
                        If Mid(entero, B + 1, 1) = "5" Then
                            flag = "S"
                            expresion = expresion & "quince "
                        End If
                        If Mid(entero, B + 1, 1) > "5" Then
                            flag = "N"
                            expresion = expresion & "dieci"
                        End If
               
                    Case "2"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "veinte "
                            flag = "S"
                        Else
                            expresion = expresion & "veinti"
                            flag = "N"
                        End If
                   
                    Case "3"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "treinta "
                            flag = "S"
                        Else
                            expresion = expresion & "treinta y "
                            flag = "N"
                        End If
               
                    Case "4"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "cuarenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cuarenta y "
                            flag = "N"
                        End If
               
                    Case "5"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "cincuenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cincuenta y "
                            flag = "N"
                        End If
               
                    Case "6"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "sesenta "
                            flag = "S"
                        Else
                            expresion = expresion & "sesenta y "
                            flag = "N"
                        End If
               
                    Case "7"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "setenta "
                            flag = "S"
                        Else
                            expresion = expresion & "setenta y "
                            flag = "N"
                        End If
               
                    Case "8"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "ochenta "
                            flag = "S"
                        Else
                            expresion = expresion & "ochenta y "
                            flag = "N"
                        End If
               
                    Case "9"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "noventa "
                            flag = "S"
                        Else
                            expresion = expresion & "noventa y "
                            flag = "N"
                        End If
                End Select
               
            Case 1, 4, 7
                If paso = 1 And Mid(entero, B, 1) <> 0 Then
                    flag = "N"
                End If
                Select Case Mid(entero, B, 1)
                    Case "1"
                        If flag = "N" Then
                            If paso = 1 Then
                                expresion = expresion & "uno "
                            Else
                                expresion = expresion & "un "
                            End If
                        End If
                    Case "2"
                        If flag = "N" Then
                            expresion = expresion & "dos "
                        End If
                    Case "3"
                        If flag = "N" Then
                            expresion = expresion & "tres "
                        End If
                    Case "4"
                        If flag = "N" Then
                            expresion = expresion & "cuatro "
                        End If
                    Case "5"
                        If flag = "N" Then
                            expresion = expresion & "cinco "
                        End If
                    Case "6"
                        If flag = "N" Then
                            expresion = expresion & "seis "
                        End If
                    Case "7"
                        If flag = "N" Then
                            expresion = expresion & "siete "
                        End If
                    Case "8"
                        If flag = "N" Then
                            expresion = expresion & "ocho "
                        End If
                    Case "9"
                        If flag = "N" Then
                            expresion = expresion & "nueve "
                        End If
                End Select
            End Select
            If paso = 4 Then
                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                   Len(entero) <= 7) Then
                   If Len(entero) = 7 Then
                    If Mid(entero, 2, 3) <> "000" Then
                        expresion = expresion & "mil "
                    End If
                   Else
                        expresion = expresion & "mil "
                   End If
                End If
            End If
            If paso = 7 Then
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    expresion = expresion & "mill�n "
                Else
                    expresion = expresion & "millones "
                End If
            End If
        Next paso
       
        If deci <> "" Then
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion & "con " & deci & "/100"
            Else
                EnLetras = expresion & "con " & deci & "/100"
            End If
        Else
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion
            Else
                EnLetras = expresion
            End If
        End If
    Else 'si el numero a convertir esta fuera del rango superior e inferior
        EnLetras = ""
    End If
End Function


