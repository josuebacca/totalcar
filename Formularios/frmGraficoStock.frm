VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGraficoStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grafico Stock"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   Icon            =   "frmGraficoStock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   10395
      Picture         =   "frmGraficoStock.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6615
      Width           =   840
   End
   Begin VB.Frame FrameGrafico 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   30
      TabIndex        =   1
      Top             =   15
      Width           =   11370
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   6240
         Left            =   90
         OleObjectBlob   =   "frmGraficoStock.frx":0614
         TabIndex        =   2
         Top             =   180
         Width           =   11235
      End
   End
End
Attribute VB_Name = "frmGraficoStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StockFisico As Integer
Dim StockPendiente As Integer
Dim StockDisponible As Integer
Dim StockCovit As Integer
Dim CantidadCol As Integer
Dim TipoGrafico As String
Dim Producto As String

Private Sub CmdSalir_Click()
    Set frmGraficoStock = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Call Centrar_pantalla(Me)
        StockFisico = CInt(frmControlStock.GrdModulos.TextMatrix(frmControlStock.GrdModulos.RowSel, 6))
        StockPendiente = CInt(frmControlStock.GrdModulos.TextMatrix(frmControlStock.GrdModulos.RowSel, 5))
        StockDisponible = 0 ' CInt(frmControlStock.GrdModulos.TextMatrix(frmControlStock.GrdModulos.RowSel, 6))
        If frmControlStock.GrdModulos.ColWidth(7) <> 0 Then
            StockCovit = 0 'CInt(frmControlStock.GrdModulos.TextMatrix(frmControlStock.GrdModulos.RowSel, 7))
            CantidadCol = 2
        Else
            CantidadCol = 2
            StockCovit = 0
        End If
        TipoGrafico = frmControlStock.cboTipoGrafico.Text
        Producto = frmControlStock.GrdModulos.TextMatrix(frmControlStock.GrdModulos.RowSel, 1)
    Call Grafico(StockPendiente, StockFisico, StockDisponible, StockCovit, TipoGrafico, Producto)
End Sub

Private Sub Grafico(StockPend As Integer, StockFis As Integer, StockDis As Integer, StockCov As Integer, TipoGraf As String, Prod As String)
 Dim column As Integer
 Dim row As Integer
    With MSChart1
      ' Muestra un gráfico 3d con 8 columnas y 8 filas
      ' de datos.
      If TipoGraf = "Gráfico 2D" Then
        .ChartType = VtChChartType2dBar
      Else
        .ChartType = VtChChartType3dBar
      End If
      .ColumnCount = CantidadCol
      .RowCount = 1
      
      .TextLengthType = VtTextLengthTypeVirtual
      .RowLabel = "Producto: " & Prod
      .ColumnLabelCount = CantidadCol
      For column = 1 To CantidadCol
         For row = 1 To 1
            .column = column
            .row = row
           Select Case column
            Case 1
                .ColumnLabel = "Stock Fisico (" & StockFis & ")"
                .Data = row * StockFis
           Case 2
                .ColumnLabel = "Stock Mínimo (" & StockPend & ")"
                .Data = row * StockPend
'           Case 3
'                .ColumnLabel = "Disponible (" & StockDis & ")"
'                .Data = row * StockDis
'           Case 4
'                .ColumnLabel = "Covit (" & StockCov & ")"
'                .Data = row * StockCov
           End Select
         Next row
      Next column
   End With
End Sub

