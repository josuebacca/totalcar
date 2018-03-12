VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReadFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ------------------------------------------------------------
'  Copyright ©2001 Mike G --> IvbNET.COM
'  All Rights Reserved, http://www.ivbnet.com
'  EMAIL : webmaster@ivbnet.com
' ------------------------------------------------------------
'  You are free to use this code within your own applications,
'  but you are forbidden from selling or distributing this
'  source code without prior written consent.
' ------------------------------------------------------------


'Read Excel File Using ADO
Public Function Read_Excel _
         (ByVal sFile _
          As String) As ADODB.Recordset

      On Error GoTo fix_err
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim sconn As String

      rs.CursorLocation = adUseClient
      rs.CursorType = adOpenKeyset
      rs.LockType = adLockBatchOptimistic

      sconn = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & sFile
      rs.Open "SELECT * FROM [EXELPREC$]", sconn
      Set Read_Excel = rs
      Set rs = Nothing
      Exit Function
fix_err:
      Debug.Print Err.Description + " " + _
                  Err.Source, vbCritical, "Import"
      Err.Clear
End Function

'**********************************************************************************
'Read Text files
'You can use Extended Properties='text;FMT=Delimited'"
'By adding a third argument we can tell ADO that the file doesn't contain headers.
'The argument named HDR takes YES or NO .
'connOpen "Provider=Microsoft.Jet" _
'         & ".OLEDB.4.0;Data Source=" & App.Path _
'         & ";Extended Properties='text;HDR=NO;" _
'         & "FMT=Delimited'"
'You can use Microsoft Text Driver or Microsoft.Jet
'**********************************************************************************

Public Function Read_Text_File() As ADODB.Recordset

      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim conn As ADODB.Connection
      Set conn = New ADODB.Connection
      conn.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
                  "DBQ=" & App.Path & ";", "", ""

      rs.Open "select * from [test#txt]", conn, adOpenStatic, _
                  adLockReadOnly ', adCmdText
      Set Read_Text_File = rs
      Set rs = Nothing
      Set conn = Nothing
End Function
