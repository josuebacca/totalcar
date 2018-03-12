VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEntregaProductos 
   Caption         =   "Salida de Mercadería"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&An&ular"
      DisabledPicture =   "frmEntregaProductos.frx":0000
      Height          =   720
      Left            =   7215
      Picture         =   "frmEntregaProductos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmEntregaProductos.frx":0614
      Height          =   720
      Left            =   8085
      Picture         =   "frmEntregaProductos.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmEntregaProductos.frx":0C28
      Height          =   720
      Left            =   5460
      Picture         =   "frmEntregaProductos.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmEntregaProductos.frx":123C
      Height          =   720
      Left            =   6345
      Picture         =   "frmEntregaProductos.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6255
      Left            =   30
      TabIndex        =   20
      Top             =   60
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   529
      ForeColor       =   -2147483630
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
      TabPicture(0)   =   "frmEntregaProductos.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameTransporte"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grdgrilla"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frameGeneral"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameRemito"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmEntregaProductos.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdModulos"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      Begin VB.Frame FrameRemito 
         Caption         =   "Remito ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   4245
         TabIndex        =   33
         Top             =   390
         Width           =   4560
         Begin VB.TextBox txtNroRemito 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1515
            TabIndex        =   3
            Top             =   435
            Width           =   1065
         End
         Begin VB.TextBox txtCodigoStock 
            Height          =   300
            Left            =   3435
            TabIndex        =   34
            Top             =   660
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox txtRemSuc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   960
            MaxLength       =   4
            TabIndex        =   2
            Top             =   435
            Width           =   555
         End
         Begin MSFlexGridLib.MSFlexGrid grillaRemito 
            Height          =   675
            Left            =   75
            TabIndex        =   35
            Top             =   1155
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   1191
            _Version        =   393216
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   270
            BackColor       =   12648447
            BackColorBkg    =   -2147483633
            GridLines       =   0
            GridLinesFixed  =   1
            ScrollBars      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker FechaRemito 
            Height          =   315
            Left            =   960
            TabIndex        =   4
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61210625
            CurrentDate     =   41098
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   240
            TabIndex        =   37
            Top             =   465
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   345
            TabIndex        =   36
            Top             =   810
            Width           =   495
         End
      End
      Begin VB.Frame frameGeneral 
         Caption         =   "Datos Enterga..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   105
         TabIndex        =   21
         Top             =   390
         Width           =   4140
         Begin VB.TextBox txtdescristock 
            BackColor       =   &H8000000B&
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
            Height          =   315
            Left            =   945
            TabIndex        =   31
            Top             =   1455
            Width           =   3120
         End
         Begin VB.TextBox txtNentrega 
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
            Height          =   315
            Left            =   945
            MaxLength       =   10
            TabIndex        =   19
            Top             =   360
            Width           =   1485
         End
         Begin VB.ComboBox cboEmpleado 
            Height          =   315
            Left            =   945
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1125
            Width           =   3120
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   960
            TabIndex        =   0
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61210625
            CurrentDate     =   41098
         End
         Begin VB.Label lblEstadoSalida 
            AutoSize        =   -1  'True
            Caption         =   "ESTADO SALIDA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   930
            TabIndex        =   39
            Top             =   1755
            Width           =   1800
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   300
            TabIndex        =   38
            Top             =   1770
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Stock:"
            Height          =   195
            Left            =   420
            TabIndex        =   32
            Top             =   1485
            Width           =   465
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   285
            TabIndex        =   30
            Top             =   405
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Encargado:"
            Height          =   195
            Left            =   60
            TabIndex        =   23
            Top             =   1170
            Width           =   825
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha:"
            Height          =   195
            Index           =   1
            Left            =   405
            TabIndex        =   22
            Top             =   780
            Width           =   480
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Buscar por..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   -74880
         TabIndex        =   24
         Top             =   630
         Width           =   8685
         Begin VB.CheckBox chkEncargado 
            Caption         =   "Encargado"
            Height          =   195
            Left            =   225
            TabIndex        =   12
            Top             =   480
            Width           =   1155
         End
         Begin VB.ComboBox cboEmpleado1 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   390
            Width           =   4005
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   960
            Left            =   7530
            MaskColor       =   &H000000FF&
            Picture         =   "frmEntregaProductos.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Buscar Nota de Pedido"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   465
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   225
            TabIndex        =   13
            Top             =   750
            Width           =   810
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   2640
            TabIndex        =   15
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61210625
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   5190
            TabIndex        =   16
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   61210625
            CurrentDate     =   41098
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Encargado:"
            Height          =   195
            Left            =   1695
            TabIndex        =   41
            Top             =   435
            Width           =   825
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4155
            TabIndex        =   26
            Top             =   780
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1545
            TabIndex        =   25
            Top             =   780
            Width           =   1005
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3870
         Left            =   -74655
         TabIndex        =   27
         Top             =   2340
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6826
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdModulos 
         Height          =   3750
         Left            =   -74880
         TabIndex        =   18
         Top             =   2070
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   6615
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdgrilla 
         Height          =   2775
         Left            =   105
         TabIndex        =   7
         Top             =   3360
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   4895
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame FrameTransporte 
         Height          =   900
         Left            =   105
         TabIndex        =   42
         Top             =   2415
         Width           =   8700
         Begin VB.ComboBox cboTransporte 
            Height          =   315
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   165
            Width           =   3780
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   300
            Left            =   1215
            MaxLength       =   50
            TabIndex        =   6
            Top             =   510
            Width           =   7395
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Transporte:"
            Height          =   195
            Left            =   360
            TabIndex        =   44
            Top             =   195
            Width           =   810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   60
            TabIndex        =   43
            Top             =   555
            Width           =   1110
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   28
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Salida"
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
      Left            =   3105
      TabIndex        =   40
      Top             =   6585
      Width           =   2025
   End
   Begin VB.Label lblestado 
      AutoSize        =   -1  'True
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
      Left            =   120
      TabIndex        =   29
      Top             =   6540
      Width           =   750
   End
End
Attribute VB_Name = "frmEntregaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim Consulta As Boolean

Private Sub chkEncargado_Click()
    If chkEncargado.Value = Checked Then
        cboEmpleado1.ListIndex = 0
        cboEmpleado1.Enabled = True
    Else
        cboEmpleado1.ListIndex = -1
        cboEmpleado1.Enabled = False
    End If
End Sub

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        FechaDesde.SetFocus
        FechaDesde.Enabled = True
        FechaHasta.Enabled = True
    Else
        FechaDesde.Enabled = False
        FechaHasta.Enabled = False
        FechaDesde.Value = Null
        FechaHasta.Value = Null
    End If
End Sub

Private Sub CmdBorrar_Click()
    If txtNentrega.Text <> "" Then
        If MsgBox("Confirma Anular Salida de Mercadería", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            On Error GoTo HayError
            lblEstado.Caption = "Anulando..."
            Screen.MousePointer = vbNormal
            DBConn.BeginTrans
            'LE PONGO EL ESTADO ANULADO A LA SALIDA
            sql = "UPDATE ENTREGA_PRODUCTO"
            sql = sql & " SET EST_CODIGO=2" 'ESTADO ANULADO
            sql = sql & " WHERE EGA_CODIGO=" & XN(txtNentrega.Text)
            DBConn.Execute sql
            'Aca actualizo las ENTREGAS en DETALLE_STOCK
            For I = 1 To grdGrilla.Rows - 1
                sql = "UPDATE DETALLE_STOCK"
                sql = sql & " SET DST_STKFIS = DST_STKFIS + " & XN(grdGrilla.TextMatrix(I, 2))
                sql = sql & " ,DST_STKPEN = DST_STKPEN +  " & XN(grdGrilla.TextMatrix(I, 2))
                sql = sql & " WHERE STK_CODIGO= " & XN(txtCodigoStock.Text)
                sql = sql & " AND PTO_CODIGO =" & XN(grdGrilla.TextMatrix(I, 0))
                DBConn.Execute sql
            Next
            lblEstado.Caption = ""
            Screen.MousePointer = vbHourglass
            DBConn.CommitTrans
            CmdNuevo_Click
        End If
    End If
    Exit Sub
HayError:
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        DBConn.RollbackTrans
        MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    sql = " SELECT E.EGA_CODIGO, E.EGA_FECHA, E.EST_CODIGO, E.VEN_CODIGO,"
    sql = sql & " E.RCL_NUMERO, E.RCL_SUCURSAL, E.RCL_FECHA , V.VEN_NOMBRE,"
    sql = sql & " E.TRS_CODIGO, E.EGA_OBSERVACIONES"
    sql = sql & " FROM ENTREGA_PRODUCTO E, VENDEDOR V"
    sql = sql & " WHERE E.VEN_CODIGO = V.VEN_CODIGO"
    If chkEncargado.Value = Checked Then sql = sql & " AND E.VEN_CODIGO=" & XN(cboEmpleado1.ItemData(cboEmpleado1.ListIndex))
    If chkFecha.Value = Checked Then
        If Not IsNull(FechaDesde) Then sql = sql & " AND E.EGA_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND E.EGA_FECHA<=" & XDQ(FechaHasta)
    End If
    sql = sql & " ORDER BY E.EGA_CODIGO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
      
    If rec.EOF = False Then
        GrdModulos.Rows = 1
        Do While rec.EOF = False
            GrdModulos.AddItem Format(rec!EGA_CODIGO, "00000000") & Chr(9) & rec!EGA_FECHA _
                              & Chr(9) & Format(rec!RCL_SUCURSAL, "0000") & "-" & Format(rec!RCL_NUMERO, "00000000") _
                              & Chr(9) & rec!RCL_FECHA & Chr(9) & rec!VEN_NOMBRE & Chr(9) & rec!VEN_CODIGO _
                              & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!TRS_CODIGO & Chr(9) & rec!EGA_OBSERVACIONES
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
    End If
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    rec.Close
End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo HayError
    If validarentrega = False Then Exit Sub
         
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Guardando ..."
        
        DBConn.BeginTrans
        sql = "INSERT INTO ENTREGA_PRODUCTO(EGA_CODIGO, EGA_FECHA, VEN_CODIGO,"
        sql = sql & " RCL_NUMERO, RCL_SUCURSAL, RCL_FECHA, TRS_CODIGO,"
        sql = sql & " EGA_OBSERVACIONES, EST_CODIGO)    "
        sql = sql & " VALUES ("
        sql = sql & XN(txtNentrega) & ","
        sql = sql & XDQ(Fecha) & ","
        sql = sql & XS(cboEmpleado.ItemData(cboEmpleado.ListIndex)) & ","
        sql = sql & XN(txtNroRemito) & ","
        sql = sql & XN(txtRemSuc) & ","
        sql = sql & XDQ(FechaRemito.Value) & ","
        If cboTransporte.List(cboTransporte.ListIndex) = "<Ninguno>" Then
            sql = sql & "NULL,"
        Else
            sql = sql & XN(cboTransporte.ItemData(cboTransporte.ListIndex)) & ","
        End If
        sql = sql & XS(txtObservaciones.Text) & ","
        sql = sql & "3)" 'ESTADO DEFINITIVO
        DBConn.Execute sql
        
        'Aca actualizo las ENTREGAS en DETALLE_STOCK
        For I = 1 To grdGrilla.Rows - 1
            sql = "UPDATE DETALLE_STOCK"
            sql = sql & " SET DST_STKFIS = DST_STKFIS - " & XN(grdGrilla.TextMatrix(I, 2))
            sql = sql & " ,DST_STKPEN = DST_STKPEN -  " & XN(grdGrilla.TextMatrix(I, 2))
            sql = sql & " WHERE STK_CODIGO= " & XN(txtCodigoStock.Text)
            sql = sql & " AND PTO_CODIGO =" & XN(grdGrilla.TextMatrix(I, 0))
            DBConn.Execute sql
        Next
        
        'ACTUALIZO LA TABLA PARAMENTROS
        sql = "UPDATE PARAMETROS SET SALIDA_MERCADERIA=" & XN(txtNentrega.Text)
        DBConn.Execute sql
        
        Screen.MousePointer = vbNormal
        lblEstado.Caption = ""
        DBConn.CommitTrans
        CmdNuevo_Click
    Exit Sub
         
HayError:
         lblEstado.Caption = ""
         DBConn.RollbackTrans
         Screen.MousePointer = vbNormal
         MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Function validarentrega()
    If cboEmpleado.ListIndex = -1 Then
        MsgBox "No ha ingresado el Encargado de Depósito", vbExclamation, TIT_MSGBOX
        cboEmpleado.SetFocus
        validarentrega = False
        Exit Function
    End If
    If IsNull(Fecha.Value) Then
        MsgBox "No ha ingresado la Fecha de Entrada de Productos", vbExclamation, TIT_MSGBOX
        Fecha.SetFocus
        validarentrega = False
        Exit Function
    End If
    If grdGrilla.Rows = 1 Then
        MsgBox "Debe haber al menos un producto en la Grilla", vbExclamation, TIT_MSGBOX
        txtRemSuc.SetFocus
        validarentrega = False
        Exit Function
    End If
    If txtRemSuc.Text = "" Or txtNroRemito.Text Then
        MsgBox "No ha ingresado el Remito", vbExclamation, TIT_MSGBOX
        txtRemSuc.SetFocus
        validarentrega = False
        Exit Function
    End If
    validarentrega = True
    
End Function
Private Sub CmdNuevo_Click()
    Consulta = False 'no hace consulta
    frameGeneral.Enabled = True
    FrameRemito.Enabled = True
    FrameTransporte.Enabled = True
    txtNentrega.Text = ""
    LimpiarRemito
    txtdescristock.Text = ""
    txtObservaciones.Text = ""
    cboTransporte.ListIndex = 0
    cboEmpleado.ListIndex = 0
    Fecha.Value = Date
    grdGrilla.Rows = 1
    grdGrilla.HighLight = flexHighlightNever
    'BUSCA ESTADO
    Call BuscoEstado(1, lblEstadoSalida)
    'BUSCO NUMERO DE ENTREGA
    txtNentrega.Text = BuscoNroEntrega()
    Fecha.SetFocus
    cmdGrabar.Enabled = True
    CmdBorrar.Enabled = False
    tabDatos.Tab = 0
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmEntregaProductos = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    lblEstado.Caption = ""
    Call Centrar_pantalla(Me)
    preparogrilla
    'CARGO COMBO EMPLEADO
    cargocboEmpl
    'CARGO COMBO TRANSPORTE
    CargoComboTransporte
    Fecha = Date
    tabDatos.Tab = 0
    grdGrilla.HighLight = flexHighlightNever
    'BUSCA ESTADO
    Call BuscoEstado(1, lblEstadoSalida)
    'BUSCO NUMERO DE ENTREGA
    txtNentrega.Text = BuscoNroEntrega()
    Consulta = False 'no hace consulta
    CmdBorrar.Enabled = False
End Sub

Private Sub CargoComboTransporte()
    sql = "SELECT TRS_DESCRI,TRS_CODIGO FROM TRANSPORTE ORDER BY TRS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboTransporte.AddItem "<Ninguno>"
        Do While rec.EOF = False
            cboTransporte.AddItem rec!TRS_DESCRI
            cboTransporte.ItemData(cboTransporte.NewIndex) = rec!TRS_CODIGO
            rec.MoveNext
        Loop
        cboTransporte.ListIndex = 0
    End If
    rec.Close
End Sub

Private Function BuscoNroEntrega() As String
    sql = "SELECT (SALIDA_MERCADERIA) + 1 AS NUMERO FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscoNroEntrega = Format(rec!Numero, "00000000")
    Else
        BuscoNroEntrega = ""
    End If
    rec.Close
End Function

Private Sub preparogrilla()
    grdGrilla.FormatString = "Codigo|Producto|Cantidad"
    grdGrilla.ColWidth(0) = 1000 'codigo
    grdGrilla.ColWidth(1) = 5500 'producto
    grdGrilla.ColWidth(2) = 1000 'cantidad
    grdGrilla.ColWidth(3) = 0
    grdGrilla.Rows = 1
    '------------------------------------
    GrdModulos.FormatString = "^Numero|^Fecha|^Remito|Fec Remito|Encargado|" _
                             & "CODIGO VENDEDOR|ESTADO SALIDA|TRANSPORTE|OBSERVACIONES"
    GrdModulos.ColWidth(0) = 1200
    GrdModulos.ColWidth(1) = 1200
    GrdModulos.ColWidth(2) = 1300
    GrdModulos.ColWidth(3) = 1200
    GrdModulos.ColWidth(4) = 3000
    GrdModulos.ColWidth(5) = 0
    GrdModulos.ColWidth(6) = 0
    GrdModulos.ColWidth(7) = 0
    GrdModulos.ColWidth(8) = 0
    GrdModulos.Rows = 1
    '------------------------------------
    grillaRemito.ColWidth(0) = 800
    grillaRemito.ColWidth(1) = 3500
    grillaRemito.ColWidth(2) = 0
    grillaRemito.TextMatrix(0, 0) = "    Cliente:"
    grillaRemito.TextMatrix(1, 0) = " Sucursal:"
    '------------------------------------
End Sub

Private Sub cargocboEmpl()
    sql = "SELECT VEN_CODIGO, VEN_NOMBRE"
    sql = sql & " FROM VENDEDOR ORDER BY VEN_NOMBRE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboEmpleado.AddItem rec!VEN_NOMBRE
            cboEmpleado.ItemData(cboEmpleado.NewIndex) = rec!VEN_CODIGO
            cboEmpleado1.AddItem rec!VEN_NOMBRE
            cboEmpleado1.ItemData(cboEmpleado1.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        cboEmpleado.ListIndex = 0
        cboEmpleado1.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        CmdNuevo_Click
        Consulta = True
        txtNentrega.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        Fecha.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
        txtRemSuc.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 2), 4)
        txtNroRemito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 2), 8)
        txtNroRemito_LostFocus
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 5)), cboEmpleado)
        Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6)), lblEstadoSalida)
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 7)), cboTransporte)
        txtObservaciones.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
        cmdGrabar.Enabled = False
        frameGeneral.Enabled = False
        FrameRemito.Enabled = False
        FrameTransporte.Enabled = False
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = "2" Then
            CmdBorrar.Enabled = False
        Else
            CmdBorrar.Enabled = True
        End If
        tabDatos.Tab = 0
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
If tabDatos.Tab = 1 Then
    cboEmpleado1.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    cmdGrabar.Enabled = False
    CmdBorrar.Enabled = False
    LimpiarBusqueda
    If Me.Visible = True Then chkEncargado.SetFocus
  Else
    If Me.Visible = True Then grdGrilla.SetFocus
  End If
End Sub

Private Sub LimpiarBusqueda()
    cboEmpleado1.ListIndex = -1
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    GrdModulos.Rows = 1
    chkEncargado.Value = Unchecked
    chkFecha.Value = Unchecked
End Sub

Private Sub txtNentrega_GotFocus()
    SelecTexto txtNentrega
End Sub

Private Sub txtNentrega_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroRemito_Change()
    If txtNroRemito.Text = "" Then
        FechaRemito.Value = Null
    End If
End Sub

Private Sub txtNroRemito_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroRemito_LostFocus()
    Set rec = New ADODB.Recordset

    If txtNroRemito.Text <> "" Then
        txtNroRemito.Text = Format(txtNroRemito.Text, "00000000")
        'BUSCO DATOS DEL REMITO
        sql = "SELECT RC.RCL_NUMERO, RC.RCL_SUCURSAL, RC.RCL_FECHA, RC.EST_CODIGO,"
        sql = sql & "  RC.STK_CODIGO, E.EST_DESCRI, NP.CLI_CODIGO, NP.SUC_CODIGO"
        sql = sql & " FROM REMITO_CLIENTE RC, NOTA_PEDIDO NP, ESTADO_DOCUMENTO E"
        sql = sql & " WHERE RC.RCL_NUMERO=" & XN(txtNroRemito)
        sql = sql & " AND RC.RCL_SUCURSAL=" & XN(txtRemSuc.Text)
        sql = sql & " AND RC.NPE_NUMERO=NP.NPE_NUMERO"
        sql = sql & " AND RC.NPE_FECHA=NP.NPE_FECHA"
        sql = sql & " AND RC.EST_CODIGO=E.EST_CODIGO"

        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic

        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                MsgBox "Hay mas de un Remito con el Número: " & txtNroRemito.Text, vbInformation, TIT_MSGBOX
                rec.Close
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Buscando..."

            'CARGO CABECERA DEL REMITO
            FechaRemito.Value = rec!RCL_FECHA
            Set Rec1 = New ADODB.Recordset
            grillaRemito.TextMatrix(0, 1) = BuscoCliente(rec!CLI_CODIGO)
            grillaRemito.TextMatrix(1, 1) = BuscoSucursal(rec!SUC_CODIGO, rec!CLI_CODIGO)
            grillaRemito.TextMatrix(0, 2) = rec!CLI_CODIGO 'guardo el codigo del cliente
            txtCodigoStock.Text = rec!STK_CODIGO

            If rec!EST_CODIGO <> 3 Then
                Screen.MousePointer = vbNormal
                lblEstado.Caption = ""
                
                MsgBox "El Remito número: " & txtNroRemito.Text & Chr(13) & _
                       "No puede ser utilizado por su estado (" & rec!EST_DESCRI & ")", vbExclamation, TIT_MSGBOX
                
                cmdGrabar.Enabled = False
                rec.Close
                LimpiarRemito
                txtRemSuc.SetFocus
                Exit Sub
            Else
                cmdGrabar.Enabled = True
            End If
        Else
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
            MsgBox "El Remito no Existe", vbCritical, TIT_MSGBOX
            LimpiarRemito
            txtRemSuc.SetFocus
            rec.Close
            Exit Sub
        End If
        rec.Close
        
        If Consulta = False Then 'entra cuando no es consulta
            'UNA CONSULTANDO POR SI EXITE EL REMITO EN ENTREGA_PRODUCTOS
            sql = "SELECT EGA_CODIGO"
            sql = sql & " FROM ENTREGA_PRODUCTO"
            sql = sql & " WHERE RCL_NUMERO=" & XN(txtNroRemito)
            sql = sql & " AND RCL_SUCURSAL=" & XN(txtRemSuc.Text)
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
            If rec.EOF = False Then
                Screen.MousePointer = vbNormal
                lblEstado.Caption = ""
                MsgBox "El Remito ya fue usado en la Salida de Mercadería Nro: " & Format(rec!EGA_CODIGO, "00000000"), vbExclamation, TIT_MSGBOX
                LimpiarRemito
                rec.Close
                txtRemSuc.SetFocus
                Exit Sub
            End If
            rec.Close
        End If
        
        'BUSCO LA DESCRIPCION DEL STOCK
        sql = "SELECT REP_RAZSOC"
        sql = sql & " FROM STOCK S, REPRESENTADA R"
        sql = sql & " WHERE S.REP_CODIGO=R.REP_CODIGO"
        sql = sql & " AND S.STK_CODIGO=" & XN(txtCodigoStock.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtdescristock.Text = rec!REP_RAZSOC
        End If
        rec.Close
        
        '-----BUSCO LOS DATOS DEL DETALLE DEL REMITO----------------------------------
        sql = "SELECT DRC.DRC_CANTIDAD, DRC.PTO_CODIGO, DRC.DRC_NROITEM, P.PTO_DESCRI"
        sql = sql & " FROM DETALLE_REMITO_CLIENTE DRC, PRODUCTO P"
        sql = sql & " WHERE DRC.RCL_NUMERO=" & XN(txtNroRemito.Text)
        sql = sql & " AND DRC.RCL_SUCURSAL=" & XN(txtRemSuc.Text)
        sql = sql & " AND DRC.PTO_CODIGO=P.PTO_CODIGO"
        sql = sql & " ORDER BY DRC.DRC_NROITEM"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        grdGrilla.Rows = 1
        If rec.EOF = False Then
            grdGrilla.HighLight = flexHighlightAlways
            Do While rec.EOF = False
                grdGrilla.AddItem rec!PTO_CODIGO & Chr(9) & _
                                  rec!PTO_DESCRI & Chr(9) & _
                                  rec!DRC_CANTIDAD
                rec.MoveNext
            Loop
        End If
        rec.Close
        '--------------------------------------------------
        Screen.MousePointer = vbNormal
        lblEstado.Caption = ""
    End If
End Sub

Private Sub LimpiarRemito()
    txtRemSuc.Text = ""
    txtNroRemito.Text = ""
    FechaRemito.Value = Null
    txtCodigoStock.Text = ""
    grillaRemito.TextMatrix(0, 1) = ""
    grillaRemito.TextMatrix(1, 1) = ""
End Sub

Private Function BuscoCliente(Codigo As String) As String
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(Codigo)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            BuscoCliente = Rec1!CLI_RAZSOC
        Else
            BuscoCliente = "No se encontro el Cliente"
        End If
        Rec1.Close
End Function

Private Function BuscoSucursal(CodigoSuc As String, CodigoCli As String) As String
        sql = "SELECT * FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(CodigoSuc)
        sql = sql & " AND CLI_CODIGO=" & XN(CodigoCli)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec1.EOF = False Then
            BuscoSucursal = Rec1!SUC_DESCRI
        Else
            BuscoSucursal = "No se encontro la Sucursal"
        End If
        Rec1.Close
End Function

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtRemSuc_GotFocus()
    txtRemSuc.Text = Sucursal
    SelecTexto txtRemSuc
End Sub

Private Sub txtRemSuc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtRemSuc_LostFocus()
    If txtRemSuc.Text = "" Then
        txtRemSuc.Text = Sucursal
    Else
        txtRemSuc.Text = Format(txtRemSuc, "0000")
    End If
End Sub

