VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAnulaDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de ...."
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "frmAnulaDocumentos.frx":0000
      Height          =   720
      Left            =   7244
      Picture         =   "frmAnulaDocumentos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5295
      Width           =   870
   End
   Begin VB.CommandButton cmdRenumerar 
      Caption         =   "&Renumerar"
      Enabled         =   0   'False
      Height          =   720
      Left            =   8136
      Picture         =   "frmAnulaDocumentos.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5295
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmAnulaDocumentos.frx":099E
      Height          =   720
      Left            =   9030
      Picture         =   "frmAnulaDocumentos.frx":0CA8
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5295
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmAnulaDocumentos.frx":0FB2
      Height          =   720
      Left            =   5460
      Picture         =   "frmAnulaDocumentos.frx":12BC
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5295
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmAnulaDocumentos.frx":15C6
      Height          =   720
      Left            =   6352
      Picture         =   "frmAnulaDocumentos.frx":18D0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5295
      Width           =   870
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3405
      Left            =   60
      TabIndex        =   10
      ToolTipText     =   "Presione barra espaciadora para anular"
      Top             =   1815
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   6006
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   8388736
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.Frame frameBuscar 
      Caption         =   "xxx..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   90
      TabIndex        =   16
      Top             =   30
      Width           =   9825
      Begin VB.CommandButton cmdBuscarVendedor 
         Height          =   300
         Left            =   4170
         MaskColor       =   &H000000FF&
         Picture         =   "frmAnulaDocumentos.frx":1BDA
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Buscar Vendedor"
         Top             =   570
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton cmdBuscarCli 
         Height          =   315
         Left            =   4170
         MaskColor       =   &H000000FF&
         Picture         =   "frmAnulaDocumentos.frx":1EE4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Buscar"
         Top             =   255
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.TextBox txtVendedor 
         Height          =   300
         Left            =   3120
         TabIndex        =   5
         Top             =   585
         Width           =   990
      End
      Begin VB.CheckBox chkVendedor 
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   315
         TabIndex        =   1
         Top             =   693
         Width           =   1035
      End
      Begin VB.CommandButton CmdBuscAprox 
         Height          =   1425
         Left            =   9105
         MaskColor       =   &H000000FF&
         Picture         =   "frmAnulaDocumentos.frx":21EE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Buscar Nota de Pedido"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.TextBox txtDesCli 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4620
         MaxLength       =   50
         TabIndex        =   17
         Tag             =   "Descripción"
         Top             =   255
         Width           =   4320
      End
      Begin VB.TextBox txtCliente 
         Height          =   300
         Left            =   3120
         MaxLength       =   40
         TabIndex        =   4
         Top             =   255
         Width           =   990
      End
      Begin VB.CheckBox chkFecha 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   315
         TabIndex        =   3
         Top             =   1179
         Width           =   810
      End
      Begin VB.CheckBox chkCliente 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   315
         TabIndex        =   0
         Top             =   450
         Width           =   1335
      End
      Begin VB.ComboBox cboDocumento 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   915
         Width           =   2400
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Tipo"
         Height          =   195
         Left            =   315
         TabIndex        =   2
         Top             =   936
         Width           =   735
      End
      Begin VB.TextBox txtDesVen 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4605
         TabIndex        =   18
         Top             =   585
         Width           =   4320
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   3120
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   52559873
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   6000
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   52559873
         CurrentDate     =   41098
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         Caption         =   "     Cliente:"
         Height          =   195
         Left            =   2160
         TabIndex        =   30
         Top             =   360
         Width           =   750
      End
      Begin VB.Label lblVendedor 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   2175
         TabIndex        =   23
         Top             =   675
         Width           =   735
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   4935
         TabIndex        =   22
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   1905
         TabIndex        =   21
         Top             =   1305
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   2550
         TabIndex        =   20
         Top             =   990
         Width           =   360
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Presione barra espaciadora para anular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Left            =   3795
      TabIndex        =   28
      Top             =   5250
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Anulado"
      Height          =   195
      Left            =   4455
      TabIndex        =   27
      Top             =   5850
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pendiente"
      Height          =   195
      Left            =   4455
      TabIndex        =   26
      Top             =   5640
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Definitivo"
      Height          =   195
      Left            =   4455
      TabIndex        =   25
      Top             =   5445
      Width           =   660
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   3780
      Top             =   5880
      Width           =   540
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   3780
      Top             =   5685
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   150
      Left            =   3780
      Top             =   5490
      Width           =   540
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
      Left            =   210
      TabIndex        =   24
      Top             =   5775
      Width           =   750
   End
End
Attribute VB_Name = "frmAnulaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TipodeAnulacion As Integer
Dim I As Integer
Dim VStockPendiente As String

Private Sub cboDocumento_LostFocus()
    If chkFecha.Value = Unchecked And ActiveControl.Name <> "cmdNuevo" _
        And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub
Private Sub chkTipo_Click()
    If chkTipo.Value = Checked Then
        cboDocumento.Enabled = True
        cboDocumento.ListIndex = 0
    Else
        cboDocumento.Enabled = False
        cboDocumento.ListIndex = -1
    End If
End Sub
Private Sub BorrarRemito()
On Error GoTo CLAVOSE
    Dim resp As String
    If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
        resp = MsgBox("Seguro desea eliminar el Remito: " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Eliminando ..."
        
        'Borro el Detalle de la Factura
        sql = "DELETE FROM DETALLE_REMITO_CLIENTE "
        sql = sql & " WHERE RCL_SUCURSAL = " & Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4)
        sql = sql & " AND RCL_NUMERO =" & Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)
        DBConn.Execute sql
        'Borro la Factura
        sql = "DELETE FROM REMITO_CLIENTE "
        sql = sql & " WHERE RCL_SUCURSAL = " & Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4)
        sql = sql & " AND RCL_NUMERO =" & Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)
        DBConn.Execute sql
                
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        CmdNuevo_Click
    Else
        MsgBox "El Remito " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & " no puede eliminarse porque no esta ANULADA ", vbInformation, TIT_MSGBOX
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Sub BorrarFactura()
    Dim resp As String
    On Error GoTo CLAVOSE
    If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
        resp = MsgBox("Seguro desea eliminar la Factura: " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & " " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 1)) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Eliminando ..."
        
        'Borro el Detalle de la Factura
        sql = "DELETE FROM DETALLE_FACTURA_CLIENTE "
        sql = sql & " WHERE FCL_SUCURSAL = " & Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        sql = sql & " AND FCL_NUMERO =" & Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        DBConn.Execute sql
        'Borro la Factura
        sql = "DELETE FROM FACTURA_CLIENTE "
        sql = sql & " WHERE FCL_SUCURSAL = " & Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        sql = sql & " AND FCL_NUMERO =" & Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        DBConn.Execute sql
        
        'Borro la Factura de la cuenta corrientes
        sql = "DELETE FROM CTA_CTE_CLIENTE "
        sql = sql & " WHERE COM_SUCURSAL = " & Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        sql = sql & " AND COM_NUMERO =" & Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        sql = sql & " AND TCO_CODIGO =" & GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
        DBConn.Execute sql
                
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        CmdNuevo_Click
    Else
        MsgBox "La Factura " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & " " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & " no puede eliminarse porque no esta ANULADA ", vbCritical, TIT_MSGBOX
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Sub BorrarNC()
    Dim resp As String
    On Error GoTo CLAVOSE
    If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3 Then
        resp = MsgBox("Seguro desea eliminar la Nota de Credito: " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & " " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 1)) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Eliminando ..."
        
        'Borro el Detalle de la Factura
        sql = "DELETE FROM DETALLE_NOTA_CREDITO_PROVEEDOR "
        sql = sql & " WHERE CPR_SUCURSAL = " & CInt(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
        sql = sql & " AND CPR_NUMERO =" & CInt(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8))
        DBConn.Execute sql
        'Borro la Factura
        sql = "DELETE FROM NOTA_CREDITO_PROVEEDOR "
        sql = sql & " WHERE CPR_NROSUC = " & CInt(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
        sql = sql & " AND CPR_NUMERO =" & CInt(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8))
        DBConn.Execute sql
                
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        CmdNuevo_Click
    Else
        'MsgBox "La Factura " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & " " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & " no puede eliminarse porque no esta ANULADA ", vbCritical, TIT_MSGBOX
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Sub BorrarRecibo()
'
End Sub
Private Sub BorrarOrdenCompra()
'
End Sub
Private Sub BorrarRemitoProveedor()
'
End Sub
Private Sub BorrarFacturaProveedor()
Dim resp As String
    On Error GoTo CLAVOSE
    If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
        resp = MsgBox("Seguro desea eliminar la Factura: " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & " " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 1)) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub

        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Eliminando ..."

        'Borro el Detalle de la Factura
        sql = "DELETE FROM DETALLE_FACTURA_PROVEEDOR "
        sql = sql & " WHERE FPR_NROSUC = " & Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        sql = sql & " AND FPR_NUMERO =" & Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        DBConn.Execute sql
        'Borro la Factura
        sql = "DELETE FROM FACTURA_PROVEEDOR "
        sql = sql & " WHERE FPR_NROSUC = " & Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        sql = sql & " AND FPR_NUMERO =" & Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        DBConn.Execute sql

        'Borro la Factura de la cuenta corrientes
        sql = "DELETE FROM CTA_CTE_PROVEEDORES "
        sql = sql & " WHERE COM_SUCURSAL = " & XS(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
        sql = sql & " AND COM_NUMERO =" & XS(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8))
        sql = sql & " AND TCO_CODIGO =" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 6))
        DBConn.Execute sql

        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        CmdNuevo_Click
    Else
        MsgBox "La Factura " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & " " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & " no puede eliminarse porque no esta ANULADA ", vbCritical, TIT_MSGBOX
    End If
    Exit Sub

CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Sub BorrarOrdenPago()
'
End Sub

Private Sub CmdBorrar_Click()
    If MsgBox("¿Confirma Eliminar?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo SeClavo
    lblEstado.Caption = "Actualizando..."
    Screen.MousePointer = vbHourglass
    DBConn.BeginTrans
    
    Select Case TipodeAnulacion
        Case 1 'Presupuestos
            BorrarPedido
        Case 2 'REMITOS
            BorrarRemito
        Case 3 'FACTURAS
            BorrarFactura
        Case 4 'RECIBOS
            BorrarRecibo
        Case 5 'ORDEN COMPRA
            BorrarOrdenCompra
        Case 6 'REMITO PROVEEDOR
            BorrarRemitoProveedor
        Case 7 'FACTURAS PROVEEDOR
            BorrarFacturaProveedor
        Case 8 'ORDEN PAGO
            BorrarOrdenPago
        Case 10 'NC
            BorrarNC
    End Select
    
    DBConn.CommitTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    CmdNuevo_Click
    Exit Sub

SeClavo:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX

End Sub

Private Sub CmdBuscAprox_Click()
    Select Case TipodeAnulacion
        Case 1 'Presupuestos
            GrdModulos.Rows = 1
            BuscoPedidos
        Case 2 'REMITOS
            GrdModulos.Rows = 1
            BuscoRemitos
        Case 3 'FACTURAS
            GrdModulos.Rows = 1
            BuscoFacturas
        Case 4 'RECIBOS
            GrdModulos.Rows = 1
            BuscoRecibos
        Case 5 'ORDEN COMPRA
            GrdModulos.Rows = 1
            BuscoOrdenesCompra
        Case 6 'REMITOS
            GrdModulos.Rows = 1
            BuscoRemitosProveedor
        Case 7 'FACTURAS
            GrdModulos.Rows = 1
            BuscoFacturasProveedor
        Case 8 'ORDEN DE PAGO
            GrdModulos.Rows = 1
            BuscoOrdenPago
        Case 9 'ORDEN DE PAGO
            GrdModulos.Rows = 1
            BuscoGastosGrales
        Case 10 'NOTA DE CREDITO
            GrdModulos.Rows = 1
            BuscoNC
            
    End Select
End Sub

Private Sub BuscoPedidos()
    lblEstado.Caption = "Buscando Presupuestos..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT NP.NPE_NUMERO, NP.NPE_FECHA,NP.NPE_TOTAL, NP.EST_CODIGO, C.CLI_CODIGO, C.CLI_RAZSOC,E.EST_DESCRI"
    sql = sql & " FROM NOTA_PEDIDO NP, CLIENTE C, ESTADO_DOCUMENTO E"
    sql = sql & " WHERE"
    sql = sql & " NP.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND NP.EST_CODIGO=E.EST_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND NP.NPE_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND NP.NPE_FECHA<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY NP.NPE_NUMERO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem Format(rec!NPE_NUMERO, "00000000") & Chr(9) & rec!NPE_FECHA _
                            & Chr(9) & Format(rec!NPE_TOTAL, "#0.00") & Chr(9) & Trim(rec!CLI_RAZSOC) & Chr(9) & Trim(rec!EST_DESCRI) _
                            & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!CLI_CODIGO
                            
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Presupuestos", vbExclamation, TIT_MSGBOX
        chkCliente.SetFocus
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub
Private Sub BuscoOrdenesCompra()
    lblEstado.Caption = "Buscando Presupuestos..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT NP.OC_NUMERO, NP.OC_FECHA, NP.EST_CODIGO, C.PROV_CODIGO, C.PROV_RAZSOC,E.EST_DESCRI"
    sql = sql & " FROM ORDEN_COMPRA NP, PROVEEDOR C, ESTADO_DOCUMENTO E"
    sql = sql & " WHERE"
    sql = sql & " NP.PROV_CODIGO=C.PROV_CODIGO"
    sql = sql & " AND NP.EST_CODIGO=E.EST_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND NP.PROV_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND NP.OC_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND NP.OC_FECHA<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY NP.OC_NUMERO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem Format(rec!OC_NUMERO, "00000000") & Chr(9) & rec!OC_FECHA _
                            & Chr(9) & "" & Chr(9) & Trim(rec!PROV_RAZSOC) & Chr(9) & Trim(rec!EST_DESCRI) _
                            & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!PROV_CODIGO
                            
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Presupuestos", vbExclamation, TIT_MSGBOX
        chkCliente.SetFocus
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub BuscoRemitos()
    lblEstado.Caption = "Buscando Remitos..."
    Screen.MousePointer = vbHourglass
    sql = "SELECT RC.RCL_NUMERO, RC.RCL_SUCURSAL, RC.RCL_FECHA,RC.RCL_TOTAL, RC.EST_CODIGO,"
    sql = sql & " C.CLI_CODIGO, C.CLI_RAZSOC, E.EST_DESCRI,"
    sql = sql & " RC.NPE_NUMERO, RC.NPE_FECHA, RC.RCL_SINFAC, RC.STK_CODIGO"
    sql = sql & " FROM REMITO_CLIENTE RC,CLIENTE C,"
    sql = sql & " ESTADO_DOCUMENTO E"
    sql = sql & " WHERE"
    sql = sql & " RC.EST_CODIGO=E.EST_CODIGO"
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND RC.RCL_NUMERO<=90000000" ' NO MUESTRA LOS REMITOS MULTIPLES
    If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
    'If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND RC.RCL_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND RC.RCL_FECHA<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY RC.RCL_SUCURSAL, RC.RCL_NUMERO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem Format(rec!RCL_SUCURSAL, "0000") & "-" & Format(rec!RCL_NUMERO, "00000000") _
                            & Chr(9) & rec!RCL_FECHA & Chr(9) & Format(rec!RCL_TOTAL, "#0.00") & Chr(9) & rec!CLI_RAZSOC _
                            & Chr(9) & rec!EST_DESCRI & Chr(9) & rec!EST_CODIGO _
                            & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!CLI_CODIGO _
                            & Chr(9) & IIf(IsNull(rec!NPE_NUMERO), "", rec!NPE_NUMERO) _
                            & Chr(9) & IIf(IsNull(rec!NPE_FECHA), "", rec!NPE_FECHA) _
                            & Chr(9) & rec!RCL_SINFAC & Chr(9) & rec!STK_CODIGO
                            
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Remitos...", vbExclamation, TIT_MSGBOX
        chkCliente.SetFocus
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub
Private Sub BuscoRemitosProveedor()
    lblEstado.Caption = "Buscando Remitos..."
    Screen.MousePointer = vbHourglass
    sql = "SELECT RC.RPR_NUMERO, RC.RPR_SUCURSAL, RC.RPR_FECHA, RC.EST_CODIGO,"
    sql = sql & " C.PROV_CODIGO, C.PROV_RAZSOC, E.EST_DESCRI,"
    sql = sql & " RC.OC_NUMERO, RC.OC_FECHA, RC.RPR_SINFAC, RC.STK_CODIGO"
    sql = sql & " FROM REMITO_PROVEEDOR RC,PROVEEDOR C,"
    sql = sql & " ESTADO_DOCUMENTO E"
    sql = sql & " WHERE"
    sql = sql & " RC.EST_CODIGO=E.EST_CODIGO"
    sql = sql & " AND RC.PROV_CODIGO=C.PROV_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.PROV_CODIGO=" & XN(txtCliente)
    'If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND RC.RPR_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND RC.RPR_FECHA<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY RC.RPR_SUCURSAL, RC.RPR_NUMERO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem Format(rec!RPR_SUCURSAL, "0000") & "-" & Format(rec!RPR_NUMERO, "00000000") _
                            & Chr(9) & rec!RPR_FECHA & Chr(9) & "" & Chr(9) & rec!PROV_RAZSOC _
                            & Chr(9) & rec!EST_DESCRI & Chr(9) & rec!EST_CODIGO _
                            & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!PROV_CODIGO _
                            & Chr(9) & IIf(IsNull(rec!OC_NUMERO), "", rec!OC_NUMERO) _
                            & Chr(9) & IIf(IsNull(rec!OC_FECHA), "", rec!OC_FECHA) _
                            & Chr(9) & rec!RPR_SINFAC & Chr(9) & rec!STK_CODIGO
                            
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Remitos...", vbExclamation, TIT_MSGBOX
        chkCliente.SetFocus
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub
'BuscoFacturasProveedor
Private Sub BuscoFacturasProveedor()
    lblEstado.Caption = "Buscando Facturas..."
    Screen.MousePointer = vbHourglass
    sql = "SELECT FC.FPR_NUMERO, FC.FPR_NROSUC, FC.FPR_FECHA,FC.FPR_TOTAL,FC.EST_CODIGO, E.EST_DESCRI,"
    sql = sql & " C.PROV_CODIGO, C.PROV_RAZSOC, TC.TCO_ABREVIA, FC.TCO_CODIGO,"
    sql = sql & " RC.RPR_NUMERO, RC.RPR_SUCURSAL, RC.RPR_FECHA, RC.STK_CODIGO,C.TPR_CODIGO"
    sql = sql & " FROM FACTURA_PROVEEDOR FC, REMITO_PROVEEDOR RC, PROVEEDOR C,"
    sql = sql & " TIPO_COMPROBANTE TC, ESTADO_DOCUMENTO E"
    sql = sql & " WHERE"
    sql = sql & " FC.RPR_NUMERO=RC.RPR_NUMERO"
    sql = sql & " AND FC.RPR_SUCURSAL=RC.RPR_SUCURSAL"
    sql = sql & " AND FC.RPR_FECHA=RC.RPR_FECHA"
    sql = sql & " AND FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND FC.EST_CODIGO=E.EST_CODIGO"
    sql = sql & " AND FC.PROV_CODIGO=C.PROV_CODIGO"
    'sql = sql & " AND FC.TPR_CODIGO=C.TPR_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.PROV_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND FC.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND FC.FPR_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND FC.FPR_FECHA<=" & XDQ(FechaHasta)
    If chkTipo.Value = Checked Then sql = sql & " AND FC.TCO_CODIGO=" & XN(cboDocumento.ItemData(cboDocumento.ListIndex))
    
    sql = sql & " ORDER BY FC.FPR_NROSUC, FC.FPR_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FPR_NROSUC, "0000") & "-" & Format(rec!FPR_NUMERO, "00000000") _
                            & Chr(9) & rec!FPR_FECHA & Chr(9) & Format(rec!FPR_TOTAL, "#0.00") & Chr(9) & rec!PROV_RAZSOC _
                            & Chr(9) & rec!EST_DESCRI & Chr(9) & rec!EST_CODIGO _
                            & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!TCO_CODIGO _
                            & Chr(9) & rec!PROV_CODIGO & Chr(9) & Format(rec!RPR_SUCURSAL, "0000") & "-" & Format(rec!RPR_NUMERO, "00000000") _
                            & Chr(9) & rec!RPR_FECHA & Chr(9) & rec!STK_CODIGO _
                            & Chr(9) & rec!TPR_CODIGO
                            
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Facturas...", vbExclamation, TIT_MSGBOX
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub


Private Sub BuscoFacturas()
    lblEstado.Caption = "Buscando Facturas..."
    Screen.MousePointer = vbHourglass
    sql = "SELECT FC.FCL_NUMERO, FC.FCL_SUCURSAL, FC.FCL_FECHA, FC.FCL_TOTAL, FC.EST_CODIGO, E.EST_DESCRI,"
    sql = sql & " C.CLI_CODIGO, C.CLI_RAZSOC, TC.TCO_ABREVIA, FC.TCO_CODIGO,"
    sql = sql & " RC.RCL_NUMERO, RC.RCL_SUCURSAL, RC.RCL_FECHA, RC.STK_CODIGO"
    sql = sql & " FROM FACTURA_CLIENTE FC, REMITO_CLIENTE RC, CLIENTE C,"
    sql = sql & " TIPO_COMPROBANTE TC, ESTADO_DOCUMENTO E"
    sql = sql & " WHERE"
    sql = sql & " FC.RCL_NUMERO=RC.RCL_NUMERO"
    sql = sql & " AND FC.RCL_SUCURSAL=RC.RCL_SUCURSAL"
    sql = sql & " AND FC.RCL_FECHA=RC.RCL_FECHA"
    sql = sql & " AND FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND FC.EST_CODIGO=E.EST_CODIGO"
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND FC.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta)
    If chkTipo.Value = Checked Then sql = sql & " AND FC.TCO_CODIGO=" & XN(cboDocumento.ItemData(cboDocumento.ListIndex))
    
    sql = sql & " ORDER BY FC.FCL_SUCURSAL, FC.FCL_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") _
                            & Chr(9) & rec!FCL_FECHA & Chr(9) & Format(rec!FCL_TOTAL, "#0.00") & Chr(9) & rec!CLI_RAZSOC _
                            & Chr(9) & rec!EST_DESCRI & Chr(9) & rec!EST_CODIGO _
                            & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!TCO_CODIGO _
                            & Chr(9) & rec!CLI_CODIGO & Chr(9) & Format(rec!RCL_SUCURSAL, "0000") & "-" & Format(rec!RCL_NUMERO, "00000000") _
                            & Chr(9) & rec!RCL_FECHA & Chr(9) & rec!STK_CODIGO _
                            & Chr(9) & ""
                            
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Facturas...", vbExclamation, TIT_MSGBOX
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub
Private Sub BuscoGastosGrales()
    lblEstado.Caption = "Buscando Gastos Generales..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT RC.GGR_NROCOMP , RC.GGR_NROSUC, RC.GGR_FECHACOMP,RC.GGR_TOTAL,"
    sql = sql & " RC.TCO_CODIGO, TC.TCO_ABREVIA,RC.PROV_CODIGO,"
    sql = sql & " C.PROV_RAZSOC, E.EST_DESCRI, RC.EST_CODIGO,C.TPR_CODIGO"
    sql = sql & " FROM GASTOS_GENERALES RC, PROVEEDOR C, ESTADO_DOCUMENTO E, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " RC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND RC.PROV_CODIGO=C.PROV_CODIGO"
    sql = sql & " AND RC.EST_CODIGO=E.EST_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.PROV_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND RC.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND RC.GGR_FECHACOMP>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND RC.GGR_FECHACOMP<=" & XDQ(FechaHasta)
    If chkTipo.Value = Checked Then sql = sql & " AND RC.TCO_CODIGO=" & XN(cboDocumento.ItemData(cboDocumento.ListIndex))
    sql = sql & " ORDER BY RC.GGR_NROSUC, RC.GGR_NROCOMP"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!GGR_NROSUC, "0000") & "-" & Format(rec!GGR_NROCOMP, "00000000") _
                               & Chr(9) & rec!GGR_FECHACOMP & Chr(9) & Format(rec!GGR_TOTAL, "#0.00") & Chr(9) & rec!PROV_RAZSOC _
                               & Chr(9) & rec!EST_DESCRI & Chr(9) & rec!EST_CODIGO _
                               & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!TCO_CODIGO _
                               & Chr(9) & rec!PROV_CODIGO & Chr(9) & rec!TPR_CODIGO
                               
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Recibos...", vbExclamation, TIT_MSGBOX
        chkCliente.SetFocus
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub
Private Sub BuscoNC()
    lblEstado.Caption = "Buscando Notas de Creditos..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT RC.CPR_NUMERO, RC.CPR_NROSUC, RC.CPR_FECHA,RC.CPR_TOTAL,"
    sql = sql & " RC.TCO_CODIGO, TC.TCO_ABREVIA,RC.PROV_CODIGO,"
    sql = sql & " C.PROV_RAZSOC, E.EST_DESCRI, RC.EST_CODIGO,C.TPR_CODIGO"
    sql = sql & " FROM NOTA_CREDITO_PROVEEDOR RC, PROVEEDOR C, ESTADO_DOCUMENTO E, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " RC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND RC.PROV_CODIGO=C.PROV_CODIGO"
    sql = sql & " AND RC.EST_CODIGO=E.EST_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.PROV_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND RC.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND RC.CPR_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND RC.CPR_FECHA<=" & XDQ(FechaHasta)
    If chkTipo.Value = Checked Then sql = sql & " AND RC.TCO_CODIGO=" & XN(cboDocumento.ItemData(cboDocumento.ListIndex))
    sql = sql & " ORDER BY RC.CPR_NROSUC, RC.CPR_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!CPR_NROSUC, "0000") & "-" & Format(rec!CPR_NUMERO, "00000000") _
                               & Chr(9) & rec!CPR_FECHA & Chr(9) & Format(rec!CPR_TOTAL, "#0.00") & Chr(9) & rec!PROV_RAZSOC _
                               & Chr(9) & rec!EST_DESCRI & Chr(9) & rec!EST_CODIGO _
                               & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!TCO_CODIGO _
                               & Chr(9) & rec!PROV_CODIGO & Chr(9) & rec!TPR_CODIGO
                               
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Recibos...", vbExclamation, TIT_MSGBOX
        chkCliente.SetFocus
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub BuscoOrdenPago()
    lblEstado.Caption = "Buscando Recibos..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT RC.OPG_NUMERO, RC.OPG_NROSUC, RC.OPG_FECHA,RC.OPG_TOTAL,"
    sql = sql & " RC.TCO_CODIGO, TC.TCO_ABREVIA,RC.PROV_CODIGO,"
    sql = sql & " C.PROV_RAZSOC, E.EST_DESCRI, RC.EST_CODIGO,C.TPR_CODIGO"
    sql = sql & " FROM ORDEN_PAGO RC, PROVEEDOR C, ESTADO_DOCUMENTO E, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " RC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND RC.PROV_CODIGO=C.PROV_CODIGO"
    sql = sql & " AND RC.EST_CODIGO=E.EST_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.PROV_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND RC.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND RC.OPG_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND RC.OPG_FECHA<=" & XDQ(FechaHasta)
    If chkTipo.Value = Checked Then sql = sql & " AND RC.TCO_CODIGO=" & XN(cboDocumento.ItemData(cboDocumento.ListIndex))
    sql = sql & " ORDER BY RC.OPG_NROSUC, RC.OPG_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!OPG_NROSUC, "0000") & "-" & Format(rec!OPG_NUMERO, "00000000") _
                               & Chr(9) & rec!OPG_FECHA & Chr(9) & Format(rec!OPG_TOTAL, "#0.00") & Chr(9) & rec!PROV_RAZSOC _
                               & Chr(9) & rec!EST_DESCRI & Chr(9) & rec!EST_CODIGO _
                               & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!TCO_CODIGO _
                               & Chr(9) & rec!PROV_CODIGO & Chr(9) & rec!TPR_CODIGO
                               
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Recibos...", vbExclamation, TIT_MSGBOX
        chkCliente.SetFocus
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub BuscoRecibos()
    lblEstado.Caption = "Buscando Recibos..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT RC.REC_NUMERO, RC.REC_SUCURSAL, RC.REC_FECHA,RC.REC_TOTAL,"
    sql = sql & " RC.TCO_CODIGO, TC.TCO_ABREVIA,RC.CLI_CODIGO,"
    sql = sql & " C.CLI_RAZSOC, E.EST_DESCRI, RC.EST_CODIGO"
    sql = sql & " FROM RECIBO_CLIENTE RC, CLIENTE C, ESTADO_DOCUMENTO E, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " RC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND RC.EST_CODIGO=E.EST_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND RC.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND RC.REC_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND RC.REC_FECHA<=" & XDQ(FechaHasta)
    If chkTipo.Value = Checked Then sql = sql & " AND RC.TCO_CODIGO=" & XN(cboDocumento.ItemData(cboDocumento.ListIndex))
    sql = sql & " ORDER BY RC.REC_SUCURSAL, RC.REC_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!REC_SUCURSAL, "0000") & "-" & Format(rec!REC_NUMERO, "00000000") _
                               & Chr(9) & rec!REC_FECHA & Chr(9) & Format(rec!REC_TOTAL, "#0.00") & Chr(9) & rec!CLI_RAZSOC _
                               & Chr(9) & rec!EST_DESCRI & Chr(9) & rec!EST_CODIGO _
                               & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!TCO_CODIGO _
                               & Chr(9) & rec!CLI_CODIGO & Chr(9) & ""
                               
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Recibos...", vbExclamation, TIT_MSGBOX
        chkCliente.SetFocus
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdBuscarVendedor_Click()
    frmBuscar.TipoBusqueda = 4
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtVendedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtDesVen.Text = frmBuscar.grdBuscar.Text
        txtVendedor.SetFocus
    Else
        txtVendedor.SetFocus
    End If
End Sub

Private Sub cmdGrabar_Click()
    If MsgBox("¿Confirma Anular?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo SeClavo
    lblEstado.Caption = "Actualizando..."
    Screen.MousePointer = vbHourglass
    DBConn.BeginTrans
    
    Select Case TipodeAnulacion
        Case 1 'Presupuestos
            ActualizoPedido
        Case 2 'REMITOS
            ActualizoRemito
        Case 3 'FACTURAS
            ActualizoFactura
        Case 4 'RECIBOS
            ActualizoRecibo
        Case 5 'ORDEN COMPRA
            ActualizoOrdenCompra
        Case 6 'REMITO PROVEEDOR
            ActualizoRemitoProveedor
        Case 7 'FACTURAS PROVEEDOR
            ActualizoFacturaProveedor
        Case 8 'ORDEN PAGO
            ActualizoOrdenPago
        Case 9 'ORDEN PAGO
            ActualizoGastoGeneral
        Case 10 'NOTA DE CREDITO
            ActualizoNotaCredito
    End Select
    
    DBConn.CommitTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    CmdNuevo_Click
    Exit Sub

SeClavo:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub ActualizoPedido()
    For I = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(I, 4) <> GrdModulos.TextMatrix(I, 5) Then
            sql = "UPDATE NOTA_PEDIDO"
            sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(I, 5))
            sql = sql & " WHERE"
            sql = sql & " NPE_NUMERO=" & XN(GrdModulos.TextMatrix(I, 0))
            sql = sql & " AND NPE_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 1))
            DBConn.Execute sql
        End If
    Next
End Sub

Private Sub BorrarPedido()
    ' SEGUIR ACÁ
    If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
        sql = "DELETE FROM DETALLE_REMITO_CLIENTE"
        sql = sql & " WHERE DRC_SUCURSAL = " & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 1))
        sql = sql & " AND DRC_NUMERO = " & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 1))
    End If
End Sub
Private Sub ActualizoRemito()
    For I = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(I, 5) <> GrdModulos.TextMatrix(I, 6) Then
            sql = "UPDATE REMITO_CLIENTE"
            sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(I, 6))
            sql = sql & " WHERE"
            sql = sql & " RCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 0), 8))
            sql = sql & " AND RCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 0), 4))
            DBConn.Execute sql
            
            'PONGO AL Presupuesto EN PENDIENTE
            If GrdModulos.TextMatrix(I, 7) <> "" Then
                sql = "UPDATE NOTA_PEDIDO"
                sql = sql & " SET EST_CODIGO=1"
                sql = sql & " WHERE"
                sql = sql & " NPE_NUMERO=" & XN(GrdModulos.TextMatrix(I, 7))
                sql = sql & " AND NPE_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 8))
                DBConn.Execute sql
            End If
            
            'ACTUALIZO EL STOCK (STOCK PENDIENTE)
            'SI ES UN REMITO SIN FACTURA
            If GrdModulos.TextMatrix(I, 9) = "S" Then
                'VStockPendiente = ""
                Set Rec2 = New ADODB.Recordset
                sql = "SELECT DS.STK_CODIGO, DS.PTO_CODIGO, DS.DST_STKPEN, DR.DRC_CANTIDAD"
                sql = sql & " FROM DETALLE_STOCK DS, DETALLE_REMITO_CLIENTE DR"
                sql = sql & " WHERE"
                sql = sql & " DS.STK_CODIGO=" & XN(GrdModulos.TextMatrix(I, 10))
                sql = sql & " AND DR.RCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 0), 8))
                sql = sql & " AND DR.RCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 0), 4))
                sql = sql & " AND DS.PTO_CODIGO=DR.PTO_CODIGO"
                Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic

                If Rec2.EOF = False Then
                    Do While Rec2.EOF = False
                        'VStockPendiente = CStr(CInt(Rec2!DST_STKPEN) - CInt(Rec2!DRC_CANTIDAD))
                        sql = "UPDATE DETALLE_STOCK"
                        sql = sql & " SET"
                        sql = sql & " DST_STKFIS = DST_STKFIS + " & CInt(Rec2!DRC_CANTIDAD)
                        sql = sql & " WHERE STK_CODIGO=" & XN(GrdModulos.TextMatrix(I, 10))
                        sql = sql & " AND PTO_CODIGO LIKE '" & XN(Rec2!PTO_CODIGO) & "'"
                        DBConn.Execute sql
                        Rec2.MoveNext
                    Loop
                End If
                Rec2.Close
            End If
            
        End If
    Next
End Sub

Private Sub ActualizoFactura()
    For I = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(I, 6) <> GrdModulos.TextMatrix(I, 7) Then
            Set Rec1 = New ADODB.Recordset
            Set Rec2 = New ADODB.Recordset
            
            sql = "SELECT FCL_TCO_CODIGO FROM FACTURAS_NOTA_CREDITO_CLIENTE"
            sql = sql & " WHERE"
            sql = sql & " FCL_TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 8))
            sql = sql & " AND FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
            sql = sql & " AND FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
            
            
            Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
            If Rec2.EOF = True Then
            
                sql = "UPDATE FACTURA_CLIENTE"
                sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(I, 7))
                sql = sql & " WHERE"
                sql = sql & " TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 8))
                sql = sql & " AND FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
                sql = sql & " AND FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
                DBConn.Execute sql
                
                'EN REMITOS MULTIPLES TENGO QUE PONER PENDIENTE A LOS REMITOS ORIGINALES Y ANULO EL MULTIPLE
                If XN(Right(GrdModulos.TextMatrix(I, 10), 8)) > 90000000 Then
                    'PONGO PENDIENTE EL REMITO
                    sql = "UPDATE REMITO_CLIENTE"
                    sql = sql & " SET EST_CODIGO=2"
                    sql = sql & " WHERE"
                    sql = sql & " RCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 10), 8))
                    sql = sql & " AND RCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 10), 4))
                    sql = sql & " AND RCL_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 11))
                    DBConn.Execute sql
                    
                    sql = "SELECT * FROM REMITOS_FACTURA "
                    sql = sql & " WHERE REF_REMITOM = " & XN(Right(GrdModulos.TextMatrix(I, 10), 8))
                    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                    If Rec1.EOF = False Then
                     Do While Rec1.EOF = False
                        sql = "UPDATE REMITO_CLIENTE SET EST_CODIGO=1"
                        sql = sql & " WHERE"
                        sql = sql & " RCL_NUMERO=" & Rec1!RCL_NUMERO
                        sql = sql & " AND RCL_SUCURSAL=" & Rec1!RCL_SUCURSAL
                        DBConn.Execute sql
                        Rec1.MoveNext
                     Loop
                    
                    End If
                    Rec1.Close
                
                Else
                
                    'PONGO PENDIENTE EL REMITO
                    sql = "UPDATE REMITO_CLIENTE"
                    sql = sql & " SET EST_CODIGO=1"
                    sql = sql & " WHERE"
                    sql = sql & " RCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 10), 8))
                    sql = sql & " AND RCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 10), 4))
                    sql = sql & " AND RCL_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 11))
                    DBConn.Execute sql
                End If
                'ACTUALIZO CTA-CTE
                DBConn.Execute QuitoCtaCteCliente(GrdModulos.TextMatrix(I, 9), GrdModulos.TextMatrix(I, 8), _
                                                  Right(GrdModulos.TextMatrix(I, 1), 8), Left(GrdModulos.TextMatrix(I, 1), 4))
                                                  
                
                'ACTUALIZO EL STOCK (STOCK PENDIENTE)
'                VStockPendiente = ""
'
'                sql = "SELECT DS.STK_CODIGO, DS.PTO_CODIGO, DS.DST_STKPEN, DF.DFC_CANTIDAD"
'                sql = sql & " FROM DETALLE_STOCK DS, DETALLE_FACTURA_CLIENTE DF"
'                sql = sql & " WHERE"
'                sql = sql & " DS.STK_CODIGO=" & XN(GrdModulos.TextMatrix(I, 11))
'                sql = sql & " AND DF.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 7))
'                sql = sql & " AND DF.FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
'                sql = sql & " AND DF.FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
'                sql = sql & " AND DS.PTO_CODIGO=DF.PTO_CODIGO"
'                Rec3.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'                If Rec3.EOF = False Then
'                    Do While Rec3.EOF = False
'                        VStockPendiente = CStr(CInt(Rec3!DST_STKPEN) + CInt(Rec3!DFC_CANTIDAD))
'                        sql = "UPDATE DETALLE_STOCK"
'                        sql = sql & " SET"
'                        sql = sql & " DST_STKPEN=" & XN(VStockPendiente)
'                        sql = sql & " ,DST_STKFIS= DST_STKFIS + " & XN(Rec3!DFC_CANTIDAD)
'                        sql = sql & " WHERE STK_CODIGO=" & XN(GrdModulos.TextMatrix(I, 11))
'                        sql = sql & " AND PTO_CODIGO LIKE '" & Rec3!PTO_CODIGO & "'"
'                        DBConn.Execute sql
'                        Rec3.MoveNext
'                    Loop
'                End If
'                Rec3.Close
            Else
                MsgBox "La Factura número: " & GrdModulos.TextMatrix(I, 1) & ", no puede ser ANULADA" & Chr(13) & _
                                           " por estar relacionada con una Nota de Crédito", vbCritical, TIT_MSGBOX
                GrdModulos_DblClick
            End If
            If Rec2.State = 1 Then Rec2.Close
        End If
    Next
End Sub

Private Sub ActualizoRecibo()
    Dim SaldoFactura As String
    SaldoFactura = "0"
    
    For I = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(I, 6) <> GrdModulos.TextMatrix(I, 7) Then
            Set rec = New ADODB.Recordset
            
            sql = "UPDATE RECIBO_CLIENTE"
            sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(I, 7))
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 8))
            sql = sql & " AND REC_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
            sql = sql & " AND REC_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
            DBConn.Execute sql
            
            'ACTUALIZO EL SALDO DE LAS FACTURAS
            sql = "SELECT FR.FCL_TCO_CODIGO, FR.FCL_NUMERO, FR.FCL_SUCURSAL, FR.FCL_FECHA,"
            sql = sql & " FR.REC_IMPORTE,FC.FCL_SALDO"
            sql = sql & " FROM FACTURAS_RECIBO_CLIENTE FR, FACTURA_CLIENTE FC"
            sql = sql & " WHERE"
            sql = sql & " FR.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 8))
            sql = sql & " AND FR.REC_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
            sql = sql & " AND FR.REC_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
            sql = sql & " AND FR.FCL_TCO_CODIGO=FC.TCO_CODIGO"
            sql = sql & " AND FR.FCL_NUMERO=FC.FCL_NUMERO"
            sql = sql & " AND FR.FCL_SUCURSAL=FC.FCL_SUCURSAL"
            
            
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                Do While rec.EOF = False
                    SaldoFactura = CDbl(rec!REC_IMPORTE) + CDbl(rec!FCL_SALDO)
                    sql = "UPDATE FACTURA_CLIENTE"
                    sql = sql & " SET FCL_SALDO=" & XN(SaldoFactura)
                    sql = sql & " WHERE"
                    sql = sql & " TCO_CODIGO=" & XN(rec!FCL_TCO_CODIGO)
                    sql = sql & " AND FCL_NUMERO=" & XN(rec!FCL_NUMERO)
                    sql = sql & " AND FCL_SUCURSAL=" & XN(rec!FCL_SUCURSAL)
                    DBConn.Execute sql
                    SaldoFactura = "0"
                    rec.MoveNext
                Loop
            End If
            rec.Close
            
            'ACTUALIZO EL SALDO DE EL DINERO A CTA DEL CLIENTE
            sql = "SELECT RS.TCO_CODIGO, RS.REC_NUMERO, RS.REC_SUCURSAL,"
            sql = sql & " RS.REC_FECHA, RS.REC_SALDO, DR.DRE_COMIMP"
            sql = sql & " FROM DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE_SALDO RS"
            sql = sql & " WHERE"
            sql = sql & " RS.TCO_CODIGO=DR.DRE_TCO_CODIGO"
            sql = sql & " AND RS.REC_NUMERO=DR.DRE_COMNUMERO"
            sql = sql & " AND RS.REC_SUCURSAL=DR.DRE_COMSUCURSAL"
            sql = sql & " AND DR.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 8))
            sql = sql & " AND DR.REC_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
            sql = sql & " AND DR.REC_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
            If rec.EOF = False Then
                Do While rec.EOF = False
                    SaldoFactura = CDbl(rec!DRE_COMIMP) + CDbl(rec!REC_SALDO)
                    sql = "UPDATE RECIBO_CLIENTE_SALDO"
                    sql = sql & " SET REC_SALDO=" & XN(SaldoFactura)
                    sql = sql & " WHERE"
                    sql = sql & " TCO_CODIGO=" & XN(rec!TCO_CODIGO)
                    sql = sql & " AND REC_NUMERO=" & XN(rec!REC_NUMERO)
                    sql = sql & " AND REC_SUCURSAL=" & XN(rec!REC_SUCURSAL)
                    DBConn.Execute sql
                    
                    SaldoFactura = "0"
                    rec.MoveNext
                Loop
            End If
            rec.Close
            
            'ACTUALIZO LA CTA-CTE
            DBConn.Execute QuitoCtaCteCliente(GrdModulos.TextMatrix(I, 9), GrdModulos.TextMatrix(I, 8), _
                                              Right(GrdModulos.TextMatrix(I, 1), 8), Left(GrdModulos.TextMatrix(I, 1), 4))
        End If
    Next
End Sub
Private Sub ActualizoOrdenCompra()
    For I = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(I, 4) <> GrdModulos.TextMatrix(I, 5) Then
            sql = "UPDATE ORDEN_COMPRA"
            sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(I, 5))
            sql = sql & " WHERE"
            sql = sql & " OC_NUMERO=" & XN(GrdModulos.TextMatrix(I, 0))
            sql = sql & " AND OC_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 1))
            DBConn.Execute sql
        End If
    Next
End Sub

Private Sub ActualizoRemitoProveedor()
    For I = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(I, 4) <> GrdModulos.TextMatrix(I, 5) Then
            sql = "UPDATE REMITO_PROVEEDOR"
            sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(I, 5))
            sql = sql & " WHERE"
            sql = sql & " RPR_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 0), 8))
            sql = sql & " AND RPR_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 0), 4))
            DBConn.Execute sql
            
            'PONGO A LA ORDEN DE COMPRA EN DEFINITIVO
            If GrdModulos.TextMatrix(I, 7) <> "" Then
                sql = "UPDATE ORDEN_COMPRA"
                sql = sql & " SET EST_CODIGO=3"
                sql = sql & " WHERE"
                sql = sql & " OC_NUMERO=" & XN(GrdModulos.TextMatrix(I, 7))
                sql = sql & " AND OC_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 8))
                DBConn.Execute sql
            End If
            
            'ACTUALIZO EL STOCK (STOCK PENDIENTE)' VER CUANDO ARREGLE LA PARTE DEL STOCK
            'SI ES UN REMITO SIN FACTURA
'            If GrdModulos.TextMatrix(I, 9) = "S" Then
'                VStockPendiente = ""
'                Set Rec2 = New ADODB.Recordset
'                sql = "SELECT DS.STK_CODIGO, DS.PTO_CODIGO, DS.DST_STKPEN, DR.DRC_CANTIDAD"
'                sql = sql & " FROM DETALLE_STOCK DS, DETALLE_REMITO_CLIENTE DR"
'                sql = sql & " WHERE"
'                sql = sql & " DS.STK_CODIGO=" & XN(GrdModulos.TextMatrix(I, 10))
'                sql = sql & " AND DR.RPR_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 0), 8))
'                sql = sql & " AND DR.RPR_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 0), 4))
'                sql = sql & " AND DS.PTO_CODIGO=DR.PTO_CODIGO"
'                Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'                If Rec2.EOF = False Then
'                    Do While Rec2.EOF = False
'                        VStockPendiente = CStr(CInt(Rec2!DST_STKPEN) - CInt(Rec2!DRC_CANTIDAD))
'                        sql = "UPDATE DETALLE_STOCK"
'                        sql = sql & " SET"
'                        sql = sql & " DST_STKPEN=" & XN(VStockPendiente)
'                        sql = sql & " WHERE STK_CODIGO=" & XN(GrdModulos.TextMatrix(I, 10))
'                        sql = sql & " AND PTO_CODIGO=" & XN(Rec2!PTO_CODIGO)
'                        DBConn.Execute sql
'                        Rec2.MoveNext
'                    Loop
'                End If
'                Rec2.Close
'            End If
            
        End If
    Next
End Sub

Private Sub ActualizoFacturaProveedor()
    For I = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(I, 6) <> GrdModulos.TextMatrix(I, 7) Then
            Set Rec2 = New ADODB.Recordset
            
            sql = "SELECT FPR_TCO_CODIGO FROM FACTURAS_NOTA_CREDITO_PROVEEDOR"
            sql = sql & " WHERE"
            sql = sql & " FPR_TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 8))
            sql = sql & " AND FPR_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
            sql = sql & " AND FPR_NROSUC=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
            
            
            Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
            If Rec2.EOF = True Then
            
                sql = "UPDATE FACTURA_PROVEEDOR"
                sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(I, 7))
                sql = sql & " WHERE"
                sql = sql & " TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 8))
                sql = sql & " AND FPR_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
                sql = sql & " AND FPR_NROSUC=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
                DBConn.Execute sql
                
                'PONGO PENDIENTE EL REMITO
                sql = "UPDATE REMITO_PROVEEDOR"
                sql = sql & " SET EST_CODIGO=1"
                sql = sql & " WHERE"
                sql = sql & " RPR_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 10), 8))
                sql = sql & " AND RPR_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 10), 4))
                sql = sql & " AND RPR_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 11))
                DBConn.Execute sql
                
                'ACTUALIZO CTA-CTE
                DBConn.Execute QuitoCtaCteProveedores(GrdModulos.TextMatrix(I, 13), GrdModulos.TextMatrix(I, 9), _
                                                       GrdModulos.TextMatrix(I, 8), Left(GrdModulos.TextMatrix(I, 1), 4), _
                                                       Right(GrdModulos.TextMatrix(I, 1), 8))
                                    

                
                'ACTUALIZO EL STOCK (STOCK PENDIENTE)
                'VStockPendiente = ""
                sql = "SELECT DS.STK_CODIGO, DS.PTO_CODIGO, DS.DST_STKPEN, DF.DFC_CANTIDAD"
                sql = sql & " FROM DETALLE_STOCK DS, DETALLE_FACTURA_CLIENTE DF"
                sql = sql & " WHERE"
                sql = sql & " DS.STK_CODIGO=" & XN(GrdModulos.TextMatrix(I, 11))
                sql = sql & " AND DF.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 7))
                sql = sql & " AND DF.FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
                sql = sql & " AND DF.FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
                sql = sql & " AND DS.PTO_CODIGO=DF.PTO_CODIGO"
                Rec3.Open sql, DBConn, adOpenStatic, adLockOptimistic

                If Rec3.EOF = False Then
                    Do While Rec3.EOF = False
                        'VStockPendiente = CStr(CInt(Rec3!DST_STKPEN) - CInt(Rec3!DFC_CANTIDAD))
                        sql = "UPDATE DETALLE_STOCK"
                        sql = sql & " SET"
                        sql = sql & " DST_STKFIS = DST_STKFIS -" & CInt(Rec3!DFC_CANTIDAD)
                        sql = sql & " WHERE STK_CODIGO=" & XN(GrdModulos.TextMatrix(I, 11))
                        sql = sql & " AND PTO_CODIGO=" & XN(Rec3!PTO_CODIGO)
                        DBConn.Execute sql
                        Rec3.MoveNext
                    Loop
                End If
                Rec3.Close
            Else
                MsgBox "La Factura número: " & GrdModulos.TextMatrix(I, 1) & ", no puede ser ANULADA" & Chr(13) & _
                                           " por estar relacionada con una Nota de Crédito", vbCritical, TIT_MSGBOX
                GrdModulos_DblClick
            End If
            If Rec2.State = 1 Then Rec2.Close
        End If
    Next
End Sub
Private Sub ActualizoNotaCredito()

    Dim SaldoFactura As String
    SaldoFactura = "0"
    
    For I = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(I, 6) <> GrdModulos.TextMatrix(I, 7) Then
            Set rec = New ADODB.Recordset
            
            sql = "UPDATE NOTA_CREDITO_PROVEEDOR"
            sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(I, 7))
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 8))
            sql = sql & " AND CPR_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
            sql = sql & " AND CPR_NROSUC=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
            DBConn.Execute sql
            
            'QUITO GASTO DE CUENTACORRIENTE
            'DBConn.Execute QuitoCtaCteProveedores(GrdModulos.TextMatrix(I, 9), GrdModulos.TextMatrix(I, 8), _
                                                       GrdModulos.TextMatrix(I, 7), Left(GrdModulos.TextMatrix(I, 1), 4), _
                                                       Right(GrdModulos.TextMatrix(I, 1), 8))
        
        End If
    Next
End Sub
Private Sub ActualizoGastoGeneral()

    Dim SaldoFactura As String
    SaldoFactura = "0"
    
    For I = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(I, 5) <> GrdModulos.TextMatrix(I, 6) Then
            Set rec = New ADODB.Recordset
            
            sql = "UPDATE GASTOS_GENERALES"
            sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(I, 6))
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 7))
            sql = sql & " AND GGR_NROCOMP=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
            sql = sql & " AND GGR_NROSUC=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
            DBConn.Execute sql
            
            'QUITO GASTO DE CUENTACORRIENTE
            'DBConn.Execute QuitoCtaCteProveedores(GrdModulos.TextMatrix(I, 9), GrdModulos.TextMatrix(I, 8), _
                                                       GrdModulos.TextMatrix(I, 7), Left(GrdModulos.TextMatrix(I, 1), 4), _
                                                       Right(GrdModulos.TextMatrix(I, 1), 8))
        
        End If
    Next
End Sub
Private Sub ActualizoOrdenPago()
    Dim SaldoFactura As String
    SaldoFactura = "0"
    
    For I = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(I, 6) <> GrdModulos.TextMatrix(I, 7) Then
            Set rec = New ADODB.Recordset
            
            sql = "UPDATE ORDEN_PAGO"
            sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(I, 7))
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 8))
            sql = sql & " AND OPG_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
            sql = sql & " AND OPG_NROSUC=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
            DBConn.Execute sql
            
            'ACTUALIZO EL SALDO DE LAS FACTURAS
            sql = "SELECT FR.FPR_TCO_CODIGO, FR.FPR_NUMERO, FR.FPR_NROSUC, FR.FPR_FECHA,"
            sql = sql & " FR.OPG_IMPORTE,FC.FPR_SALDO"
            sql = sql & " FROM FACTURAS_ORDEN_PAGO FR, FACTURA_PROVEEDOR FC"
            sql = sql & " WHERE"
            sql = sql & " FR.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 8))
            sql = sql & " AND FR.OPG_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
            'sql = sql & " AND FR.OPG_NORSUC=" & XN(Left(GrdModulos.TextMatrix(I, 1), 4))
            sql = sql & " AND FR.OPG_FECHA=" & XDQ(GrdModulos.TextMatrix(I, 2))
            sql = sql & " AND FR.FPR_TCO_CODIGO=FC.TCO_CODIGO"
            sql = sql & " AND FR.FPR_NUMERO=FC.FPR_NUMERO"
            sql = sql & " AND FR.FPR_NROSUC=FC.FPR_NROSUC"
            
            
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                Do While rec.EOF = False
                    SaldoFactura = CDbl(rec!OPG_IMPORTE) + CDbl(rec!FPR_SALDO)
                    sql = "UPDATE FACTURA_PROVEEDOR"
                    sql = sql & " SET FPR_SALDO=" & XN(SaldoFactura)
                    sql = sql & " WHERE"
                    sql = sql & " TCO_CODIGO=" & XN(rec!FPR_TCO_CODIGO)
                    sql = sql & " AND FPR_NUMERO=" & XN(rec!FPR_NUMERO)
                    sql = sql & " AND FPR_NROSUC=" & XN(rec!FPR_NROSUC)
                    DBConn.Execute sql
                    SaldoFactura = "0"
                    rec.MoveNext
                Loop
            End If
            rec.Close
            
            'ACTUALIZO EL SALDO DE EL DINERO A CTA DEL CLIENTE
            sql = "SELECT RS.TCO_CODIGO, RS.OPG_NUMERO, RS.OPG_FECHA"
            sql = sql & " ,RS.OPG_SALDO, DR.DOP_COMIMP"
            sql = sql & " FROM DETALLE_ORDEN_PAGO DR, ORDEN_PAGO_SALDO RS"
            sql = sql & " WHERE"
            sql = sql & " RS.TCO_CODIGO=DR.DOP_TCO_CODIGO"
            sql = sql & " AND RS.OPG_NUMERO=DR.DOP_COMNUMERO"
            sql = sql & " AND RS.OPG_FECHA= DR.OPG_FECHA"
            '" & XDQ(GrdModulos.TextMatrix(I, 2))
            'sql = sql & " AND RS.OPG_NROSUC=DR.DOP_COMSUCURSAL"
            sql = sql & " AND DR.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(I, 7)) & ""
            sql = sql & " AND DR.OPG_NUMERO=" & XN(Right(GrdModulos.TextMatrix(I, 1), 8))
            
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
            If rec.EOF = False Then
                Do While rec.EOF = False
                    SaldoFactura = CDbl(rec!DRE_COMIMP) + CDbl(rec!OPG_SALDO)
                    sql = "UPDATE ORDEN_PAGO_SALDO"
                    sql = sql & " SET OPG_SALDO=" & XN(SaldoFactura)
                    sql = sql & " WHERE"
                    sql = sql & " TCO_CODIGO=" & XN(rec!TCO_CODIGO)
                    sql = sql & " AND OPG_NUMERO=" & XN(rec!OPG_NUMERO)
                    sql = sql & " AND OPG_FECHA=" & XDQ(rec!OPG_FECHA)
                    DBConn.Execute sql
                    
                    SaldoFactura = "0"
                    rec.MoveNext
                Loop
            End If
            rec.Close
            ' HAY QUE PROBAR TODO ESTO!!!!!
            'ACTUALIZO LA CTA-CTE
            'DBConn.Execute QuitoCtaCteProveedores(GrdModulos.TextMatrix(I, 8), GrdModulos.TextMatrix(I, 7), _
                                              Right(GrdModulos.TextMatrix(I, 1), 8), Left(GrdModulos.TextMatrix(I, 1), 4))
        
        
            DBConn.Execute QuitoCtaCteProveedores(GrdModulos.TextMatrix(I, 10), GrdModulos.TextMatrix(I, 9), GrdModulos.TextMatrix(I, 8), _
                                                       Left(GrdModulos.TextMatrix(I, 1), 4), Right(GrdModulos.TextMatrix(I, 1), 8))
        
        End If
    Next
End Sub

Private Sub CambiColoryEstado(Estado As Boolean)
    chkTipo.Enabled = Estado
    cboDocumento.Enabled = Estado
    chkVendedor.Enabled = Estado
    If Estado = False Then
        cboDocumento.BackColor = &H8000000F
    Else
        cboDocumento.BackColor = &H80000005
    End If
End Sub

Private Sub cmdRenumerar_Click()
    frmRenumerarFac.Show vbModal
End Sub

Private Sub CmdSalir_Click()
    Set frmAnulaDocumentos = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
     Set rec = New ADODB.Recordset
     Set Rec2 = New ADODB.Recordset
     Set Rec3 = New ADODB.Recordset
     
     Call Centrar_pantalla(Me)

    If TipodeAnulacion > 4 Then
        chkCliente.Caption = "Proveedor"
        chkVendedor.Caption = "Empleado"
        lblCliente.Caption = "Proveedor:"
        lblVendedor.Caption = "Empleado:"
        
    End If
    Select Case TipodeAnulacion
        Case 1 'Presupuestos
            frmAnulaDocumentos.Caption = "Anular Presupuestos de Clientes"
            frameBuscar.Caption = "Buscar Presupuestos por..."
            ConfiguroGrillaPedidos
            Call CambiColoryEstado(False)
            
        Case 2 'REMITOS
            frmAnulaDocumentos.Caption = "Anular Remitos de Clientes"
            frameBuscar.Caption = "Buscar Remitos por..."
            ConfiguroGrillaRemito
            Call CambiColoryEstado(False)
            
        Case 3 'FACTURAS
            cmdRenumerar.Enabled = True
            frmAnulaDocumentos.Caption = "Anular Facturas de Clientes"
            frameBuscar.Caption = "Buscar Facturas por..."
            'CARGO COMBO FACTURA
            LlenarComboFactura
            ConfiguroGrillaFactura
            Call CambiColoryEstado(True)
            
        Case 4 'RECIBOS
            cmdRenumerar.Enabled = True
            frmAnulaDocumentos.Caption = "Anular Recibos de Clientes"
            frameBuscar.Caption = "Buscar Recibos por..."
            'CARGO COMBO RECIBO
            LlenarComboRecibo
            ConfiguroGrillaRecibo
            Call CambiColoryEstado(True)
        
        Case 5 'ORDEN DE COMPRA
            frmAnulaDocumentos.Caption = "Anular Ordenes de Compra"
            frameBuscar.Caption = "Buscar Ordenes de Compra por..."
            ConfiguroGrillaPedidos
            Call CambiColoryEstado(False)
            
        Case 6 'REMITOS
            frmAnulaDocumentos.Caption = "Anular Remitos de Proveedores"
            frameBuscar.Caption = "Buscar Remitos por..."
            ConfiguroGrillaRemito
            Call CambiColoryEstado(False)
            
        Case 7 'FACTURAS proveedor
            
            frmAnulaDocumentos.Caption = "Anular Facturas de Proveedores"
            frameBuscar.Caption = "Buscar Facturas por..."
            'CARGO COMBO FACTURA
            LlenarComboFactura
            ConfiguroGrillaFactura
            Call CambiColoryEstado(True)
            
        Case 8 'ORDENES DE PAGO
            frmAnulaDocumentos.Caption = "Anular Ordenes de Pago"
            frameBuscar.Caption = "Buscar Ordenes de Pago por..."
            'CARGO COMBO RECIBO
            LlenarComboRecibo
            ConfiguroGrillaRecibo
            Call CambiColoryEstado(True)
        Case 9 'GASTOS GENERALES
            'sql = "UPDATE GASTOS_GENERALES SET EST_CODIGO = 3"
            'DBConn.Execute sql
            
            frmAnulaDocumentos.Caption = "Anular Gastos Generales"
            frameBuscar.Caption = "Buscar Gastos Generales por..."
            'CARGO COMBO RECIBO
            LlenarComboRecibo
            ConfiguroGrillaGastos
            Call CambiColoryEstado(True)
        Case 10 'NOTA CREDITO
            frmAnulaDocumentos.Caption = "Anular Notas de Credito"
            frameBuscar.Caption = "Buscar Notas de Credito por..."
            'CARGO COMBO RECIBO
            LlenarComboNC
            ConfiguroGrillaNC
            Call CambiColoryEstado(True)
            
    End Select
    txtCliente.Enabled = False
    txtVendedor.Enabled = False
    cboDocumento.Enabled = False
    cmdBuscarCli.Enabled = False
    cmdBuscarVendedor.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    lblEstado.Caption = ""
End Sub
Private Sub ConfiguroGrillaGastos()
    
    If TipodeAnulacion = 8 Then
        GrdModulos.FormatString = "^Tipo Gasto|^Número|^Fecha|Importe|Proveedor|^Estado|codigo estado|" _
                                & "codigo estado que cambio|TIPO RECIBO|COD CLIENTE|REPRESENTADA"
    Else
        GrdModulos.FormatString = "^Tipo Gasto|^Número|^Fecha|Importe|Proveedor|^Estado|codigo estado|" _
                                & "codigo estado que cambio|TIPO RECIBO|COD CLIENTE|REPRESENTADA"
    End If
    GrdModulos.ColWidth(0) = 1000 'TIPO_RECIBO
    GrdModulos.ColWidth(1) = 1300 'NRO RECIBO
    GrdModulos.ColWidth(2) = 1100 'FECHA_RECIBO
    GrdModulos.ColWidth(3) = 1100 'IMPORTE
    GrdModulos.ColWidth(4) = 3000 'Proveedor
    GrdModulos.ColWidth(5) = 2000 'ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(7) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(8) = 0    'TIPO RECIBO (TCO_CODIGO)
    GrdModulos.ColWidth(9) = 0    'CODIGO CLIENTE
    GrdModulos.ColWidth(10) = 0    'REPRESENTADA
    GrdModulos.Cols = 11
    GrdModulos.Rows = 2
    
End Sub
Private Sub ConfiguroGrillaNC()
    
    If TipodeAnulacion = 10 Then
        GrdModulos.FormatString = "^Tipo Rec|^Número|^Fecha|Importe|Proveedor|^Estado|codigo estado|" _
                                & "codigo estado que cambio|TIPO NC|COD CLIENTE|REPRESENTADA"
    Else
        GrdModulos.FormatString = "^Tipo Rec|^Número|^Fecha|Importe|Cliente|^Estado|codigo estado|" _
                                & "codigo estado que cambio|TIPO NC|COD CLIENTE|REPRESENTADA"
    End If
    GrdModulos.ColWidth(0) = 1000 'TIPO_NC
    GrdModulos.ColWidth(1) = 1300 'NRO NC
    GrdModulos.ColWidth(2) = 1100 'FECHA_NC
    GrdModulos.ColWidth(3) = 1100 'IMPORTE
    GrdModulos.ColWidth(4) = 3000 'CLIENTE
    GrdModulos.ColWidth(5) = 2000 'ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(7) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(8) = 0    'TIPO NC (TCO_CODIGO)
    GrdModulos.ColWidth(9) = 0    'CODIGO CLIENTE
    GrdModulos.ColWidth(10) = 0    'REPRESENTADA
    GrdModulos.Cols = 11
    GrdModulos.Rows = 2
    
End Sub
Private Sub ConfiguroGrillaRecibo()
    
    If TipodeAnulacion = 8 Then
        GrdModulos.FormatString = "^Tipo Rec|^Número|^Fecha|Importe|Proveedor|^Estado|codigo estado|" _
                                & "codigo estado que cambio|TIPO RECIBO|COD CLIENTE|REPRESENTADA"
    Else
        GrdModulos.FormatString = "^Tipo Rec|^Número|^Fecha|Importe|Cliente|^Estado|codigo estado|" _
                                & "codigo estado que cambio|TIPO RECIBO|COD CLIENTE|REPRESENTADA"
    End If
    GrdModulos.ColWidth(0) = 1000 'TIPO_RECIBO
    GrdModulos.ColWidth(1) = 1300 'NRO RECIBO
    GrdModulos.ColWidth(2) = 1100 'FECHA_RECIBO
    GrdModulos.ColWidth(3) = 1100 'IMPORTE
    GrdModulos.ColWidth(4) = 3000 'CLIENTE
    GrdModulos.ColWidth(5) = 2000 'ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(7) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(8) = 0    'TIPO RECIBO (TCO_CODIGO)
    GrdModulos.ColWidth(9) = 0    'CODIGO CLIENTE
    GrdModulos.ColWidth(10) = 0    'REPRESENTADA
    GrdModulos.Cols = 11
    GrdModulos.Rows = 2
    
End Sub

Private Sub ConfiguroGrillaFactura()
    If TipodeAnulacion = 7 Then
        GrdModulos.FormatString = "^Tipo Fac|^Número|^Fecha|Importe|Proveedor|^Estado|codigo estado|" _
                                & "codigo estado QUE CAMBIO|TIPO FACTURA|COD CLIENTE|" _
                                & "NRO REMITO|FECHA REMITO|STOCK CODIGO|REPRESENTADA"
    Else
        GrdModulos.FormatString = "^Tipo Fac|^Número|^Fecha|Importe|Cliente|^Estado|codigo estado|" _
                                & "codigo estado QUE CAMBIO|TIPO FACTURA|COD CLIENTE|" _
                                & "NRO REMITO|FECHA REMITO|STOCK CODIGO|REPRESENTADA"
    End If
    GrdModulos.ColWidth(0) = 1000 'TIPO_FACTURA
    GrdModulos.ColWidth(1) = 1300 'NRO FACTURA
    GrdModulos.ColWidth(2) = 1100 'FECHA_FACTURA
    GrdModulos.ColWidth(3) = 1100 'IMPORTE
    GrdModulos.ColWidth(4) = 4000 'CLIENTE
    GrdModulos.ColWidth(5) = 2000 'ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(7) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(8) = 0    'TIPO FACTURA (TCO_CODIGO)
    GrdModulos.ColWidth(9) = 0    'CODIGO CLIENTE
    GrdModulos.ColWidth(10) = 0    'NUMERO REMITO
    GrdModulos.ColWidth(11) = 0   'FECHA REMITO
    GrdModulos.ColWidth(12) = 0   'STOCK CODIGO
    GrdModulos.ColWidth(13) = 0   'REPRESENTADA
    GrdModulos.Cols = 14
    GrdModulos.Rows = 2
End Sub

Private Sub ConfiguroGrillaPedidos()
    If TipodeAnulacion = 5 Then
        GrdModulos.FormatString = "^Número|^Fecha|Importe|Proveedor|^Estado|codigo estado|" _
                                  & "codigo estado QUE CAMBIO|COD CLIENTE"
    Else
        GrdModulos.FormatString = "^Número|^Fecha|Importe|Cliente|^Estado|codigo estado|" _
                                  & "codigo estado QUE CAMBIO|COD CLIENTE"
    End If
    GrdModulos.ColWidth(0) = 1000 'NUMERO
    GrdModulos.ColWidth(1) = 1100 'FECHA
    GrdModulos.ColWidth(2) = 0 'IMPORTE
    GrdModulos.ColWidth(3) = 4150 'CLIENTE
    GrdModulos.ColWidth(4) = 2000 'ESTADO
    GrdModulos.ColWidth(5) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(7) = 0    'CODIGO CLIENTE
    GrdModulos.Cols = 8
    GrdModulos.Rows = 2
End Sub

Private Sub ConfiguroGrillaRemito()
    If TipodeAnulacion = 6 Then
        GrdModulos.FormatString = "^Número|^Fecha|Importe|Proveedor|^Estado|codigo estado|" _
                                & "codigo estado QUE CAMBIO|COD CLIENTE|NRO PEDIDO|" _
                                & "FECHA PEDIDO|REMITO SIN FACTURA|CODIGO STOCK"
    Else
        GrdModulos.FormatString = "^Número|^Fecha|Importe|Cliente|^Estado|codigo estado|" _
                                & "codigo estado QUE CAMBIO|COD CLIENTE|NRO PEDIDO|" _
                                & "FECHA PEDIDO|REMITO SIN FACTURA|CODIGO STOCK"
    End If
    GrdModulos.ColWidth(0) = 1300 'NUMERO
    GrdModulos.ColWidth(1) = 1200 'FECHA
    GrdModulos.ColWidth(2) = 1200 'IMPORTE
    GrdModulos.ColWidth(3) = 3850 'CLIENTE
    GrdModulos.ColWidth(4) = 2000 'ESTADO
    GrdModulos.ColWidth(5) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(7) = 0    'CODIGO CLIENTE
    GrdModulos.ColWidth(8) = 0    'NUMERO Presupuesto
    GrdModulos.ColWidth(9) = 0    'FECHA Presupuesto
    GrdModulos.ColWidth(10) = 0    'REMITO SIN FACTURA
    GrdModulos.ColWidth(11) = 0   'CODIGO STOCK
    GrdModulos.Cols = 12
    GrdModulos.Rows = 2
End Sub

Private Sub LlenarComboFactura()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'FAC%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboDocumento.AddItem rec!TCO_DESCRI
            cboDocumento.ItemData(cboDocumento.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboDocumento.ListIndex = -1
    End If
    rec.Close
End Sub
Private Sub LlenarComboNC()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'NC%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboDocumento.AddItem rec!TCO_DESCRI
            cboDocumento.ItemData(cboDocumento.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboDocumento.ListIndex = -1
    End If
    rec.Close
End Sub


Private Sub LlenarComboRecibo()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'RECIB%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboDocumento.AddItem rec!TCO_DESCRI
            cboDocumento.ItemData(cboDocumento.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboDocumento.ListIndex = -1
    End If
    rec.Close
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        Select Case TipodeAnulacion
            Case 1 'Presupuestos
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
                    MsgBox "No se puede cambiar el estado al Presupuesto" & Chr(13) & _
                           " el mismo ya fue Anulado", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 3 Then
                    MsgBox "No se puede cambiar el estado al Presupuesto" & Chr(13) & _
                           " ya que esta asignado a un Remito", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 1 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 1
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "PENDIENTE"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlue)
                End If
                
            Case 2 'REMITOS
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
                    MsgBox "No se puede cambiar el estado al Remito" & Chr(13) & _
                           " el mismo ya fue Anulado", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 3 And GrdModulos.TextMatrix(GrdModulos.RowSel, 10) = "N" Then
                    MsgBox "No se puede cambiar el estado al Remito" & Chr(13) & _
                           " ya que esta asignado a una Factura", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 1 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 1
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "PENDIENTE"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlue)
                
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3 And GrdModulos.TextMatrix(GrdModulos.RowSel, 9) = "S" Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                End If
                
            Case 3 'FACTURAS
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    MsgBox "No se puede cambiar el estado a la Factura" & Chr(13) & _
                           ",la misma ya fue Anulado", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
                
            Case 4 'RECIBOS
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    MsgBox "No se puede cambiar el estado al Recibo" & Chr(13) & _
                           ",el mismo ya fue Anulado", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
            Case 5 'orden compa
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
                    MsgBox "No se puede cambiar el estado de la Orden de Compra" & Chr(13) & _
                           " la misma ya fue Anulada", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 4 Then
                    MsgBox "No se puede cambiar el estado de la Orden de Compra" & Chr(13) & _
                           " la que esta asignada a un Remito", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 1 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 1
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "PENDIENTE"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlue)
                End If
             Case 6 'REMITOS PROVEEDOR
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
                    MsgBox "No se puede cambiar el estado al Remito" & Chr(13) & _
                           " el mismo ya fue Anulado", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 3 And GrdModulos.TextMatrix(GrdModulos.RowSel, 9) = "N" Then
                    MsgBox "No se puede cambiar el estado al Remito" & Chr(13) & _
                           " ya que esta asignado a una Factura", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 1 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 1
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "PENDIENTE"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlue)
                
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3 And GrdModulos.TextMatrix(GrdModulos.RowSel, 9) = "S" Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                End If
             Case 7 'FACTURAS PROVEEDOR
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    MsgBox "No se puede cambiar el estado a la Factura" & Chr(13) & _
                           ",la misma ya fue Anulado", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
             Case 8 'ORDEN PAGO
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    MsgBox "No se puede cambiar el estado a la Orden de Pago" & Chr(13) & _
                           ",la misma ya fue Anulada", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
             Case 9 'GASTOS GENERALES
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    MsgBox "No se puede cambiar el estado del Gasto General" & Chr(13) & _
                           ",el mismo ya fue Anulado", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
                Case 10 'NOTA CREDITO PROVEEDOR
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    MsgBox "No se puede cambiar el estado a la Nota de Credito" & Chr(13) & _
                           ",la misma ya fue Anulada", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
        End Select
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then GrdModulos_DblClick
End Sub

Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If TipodeAnulacion < 5 Then
        If txtCliente.Text <> "" Then
            Set rec = New ADODB.Recordset
            sql = "SELECT CLI_RAZSOC FROM CLIENTE"
            sql = sql & " WHERE CLI_CODIGO=" & XN(txtCliente)
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                txtDesCli.Text = rec!CLI_RAZSOC
            Else
                MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
                txtDesCli.Text = ""
                txtCliente.SetFocus
            End If
            rec.Close
        End If
    Else
        If txtCliente.Text <> "" Then
            Set rec = New ADODB.Recordset
            sql = "SELECT PROV_RAZSOC FROM PROVEEDOR"
            sql = sql & " WHERE PROV_CODIGO=" & XN(txtCliente)
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                txtDesCli.Text = rec!PROV_RAZSOC
            Else
                MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
                txtDesCli.Text = ""
                txtCliente.SetFocus
            End If
            rec.Close
        End If
    End If
    If chkTipo.Value = Unchecked And chkFecha.Value = Unchecked _
        And chkVendedor.Value = Unchecked And ActiveControl.Name <> "cmdBuscarCli" _
        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub

Private Sub txtVendedor_Change()
    If txtVendedor.Text = "" Then
        txtDesVen.Text = ""
    End If
End Sub

Private Sub txtVendedor_GotFocus()
    SelecTexto txtVendedor
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtVendedor_LostFocus()
    If txtVendedor.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT VEN_NOMBRE"
        sql = sql & " FROM VENDEDOR"
        sql = sql & " WHERE VEN_CODIGO=" & XN(txtVendedor)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            txtDesVen.Text = Trim(rec!VEN_NOMBRE)
        Else
            MsgBox "El Vendedor no existe", vbExclamation, TIT_MSGBOX
            txtDesVen.Text = ""
            txtVendedor.Text = ""
            txtVendedor.SetFocus
        End If
        rec.Close
    End If
    If chkTipo.Value = Unchecked And chkFecha.Value = Unchecked And ActiveControl.Name <> "cmdNuevo" _
        And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub

Private Sub cmdBuscarCli_Click()
    If TipodeAnulacion > 4 Then
        frmBuscar.TipoBusqueda = 5
        frmBuscar.TxtDescriB = ""
        frmBuscar.Show vbModal
        If frmBuscar.grdBuscar.Text <> "" Then
            frmBuscar.grdBuscar.Col = 1
            txtCliente.Text = frmBuscar.grdBuscar.Text
            txtCliente.SetFocus
            txtCliente_LostFocus
        Else
            txtCliente.SetFocus
        End If
    Else
        frmBuscar.TipoBusqueda = 1
        frmBuscar.TxtDescriB = ""
        frmBuscar.Show vbModal
        If frmBuscar.grdBuscar.Text <> "" Then
            frmBuscar.grdBuscar.Col = 0
            txtCliente.Text = frmBuscar.grdBuscar.Text
            txtCliente.SetFocus
            txtCliente_LostFocus
        Else
            txtCliente.SetFocus
        End If
    End If
End Sub

Private Sub CmdNuevo_Click()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    txtVendedor.Text = ""
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    cboDocumento.ListIndex = -1
    GrdModulos.Rows = 1
    GrdModulos.Rows = 2
    txtCliente.Enabled = False
    txtVendedor.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    cboDocumento.Enabled = False
    chkCliente.Value = Unchecked
    chkVendedor.Value = Unchecked
    chkFecha.Value = Unchecked
    chkTipo.Value = Unchecked
    chkCliente.SetFocus
End Sub

Private Sub chkCliente_Click()
    If chkCliente.Value = Checked Then
        txtCliente.Enabled = True
        cmdBuscarCli.Enabled = True
    Else
        txtCliente.Enabled = False
        cmdBuscarCli.Enabled = False
    End If
End Sub

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        FechaDesde.Enabled = True
        FechaHasta.Enabled = True
    Else
        FechaDesde.Enabled = False
        FechaHasta.Enabled = False
    End If
End Sub

Private Sub chkVendedor_Click()
    If chkVendedor.Value = Checked Then
        txtVendedor.Enabled = True
        cmdBuscarVendedor.Enabled = True
    Else
        txtVendedor.Enabled = False
        cmdBuscarVendedor.Enabled = False
    End If
End Sub

