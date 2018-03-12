VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEntradaProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Mercadería"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEntradaProductos.frx":0000
   ScaleHeight     =   7635
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmEntradaProductos.frx":0D82
      Height          =   705
      Left            =   6645
      Picture         =   "frmEntradaProductos.frx":108C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6900
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmEntradaProductos.frx":1396
      Height          =   705
      Left            =   5760
      Picture         =   "frmEntradaProductos.frx":16A0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6900
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmEntradaProductos.frx":19AA
      Height          =   705
      Left            =   8415
      Picture         =   "frmEntradaProductos.frx":1CB4
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6900
      Width           =   870
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&An&ular"
      DisabledPicture =   "frmEntradaProductos.frx":1FBE
      Height          =   705
      Left            =   7530
      Picture         =   "frmEntradaProductos.frx":22C8
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6900
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6825
      Left            =   15
      TabIndex        =   28
      Top             =   30
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   12039
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
      TabPicture(0)   =   "frmEntradaProductos.frx":25D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmEntradaProductos.frx":25EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRDGrilla"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Datos Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   90
         TabIndex        =   36
         Top             =   420
         Width           =   9105
         Begin VB.TextBox txtObservaciones 
            Height          =   300
            Left            =   1245
            MaxLength       =   50
            TabIndex        =   8
            Top             =   1710
            Width           =   7740
         End
         Begin VB.ComboBox cboTransporte 
            Height          =   315
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1320
            Width           =   3120
         End
         Begin VB.TextBox txtNroSucursal 
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
            Left            =   5760
            MaxLength       =   4
            TabIndex        =   5
            Top             =   630
            Width           =   555
         End
         Begin VB.ComboBox cboRep 
            Height          =   315
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   285
            Width           =   3300
         End
         Begin VB.TextBox txtRemito 
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
            Left            =   6255
            MaxLength       =   8
            TabIndex        =   6
            Top             =   630
            Width           =   1050
         End
         Begin VB.ComboBox cboStock 
            Height          =   315
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   630
            Width           =   3120
         End
         Begin VB.TextBox txtNumero 
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
            Left            =   1245
            MaxLength       =   8
            TabIndex        =   17
            Top             =   240
            Width           =   870
         End
         Begin VB.ComboBox cboEmpleado 
            Height          =   315
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   975
            Width           =   3120
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   2910
            TabIndex        =   0
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   17170433
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker fecRemito 
            Height          =   315
            Left            =   5760
            TabIndex        =   7
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   17170433
            CurrentDate     =   41098
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   5055
            TabIndex        =   53
            Top             =   1395
            Width           =   540
         End
         Begin VB.Label lblEstadoRecepcion 
            AutoSize        =   -1  'True
            Caption         =   "ESTADO RECEPCION"
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
            Left            =   5685
            TabIndex        =   52
            Top             =   1380
            Width           =   2310
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   90
            TabIndex        =   46
            Top             =   1755
            Width           =   1110
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Transporte:"
            Height          =   195
            Left            =   375
            TabIndex        =   45
            Top             =   1350
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha:"
            Height          =   195
            Index           =   2
            Left            =   2325
            TabIndex        =   44
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Representada:"
            Height          =   195
            Left            =   4545
            TabIndex        =   43
            Top             =   330
            Width           =   1050
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Remito Nro:"
            Height          =   195
            Left            =   4755
            TabIndex        =   41
            Top             =   690
            Width           =   840
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Stock:"
            Height          =   195
            Left            =   720
            TabIndex        =   40
            Top             =   690
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   585
            TabIndex        =   39
            Top             =   300
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Remito:"
            Height          =   195
            Index           =   1
            Left            =   4575
            TabIndex        =   38
            Top             =   1020
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Encargado:"
            Height          =   195
            Left            =   360
            TabIndex        =   37
            Top             =   1005
            Width           =   825
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Agregar Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4140
         Left            =   90
         TabIndex        =   34
         Top             =   2565
         Width           =   9120
         Begin VB.TextBox txtdescri 
            Height          =   315
            Left            =   1065
            TabIndex        =   10
            Top             =   570
            Width           =   4515
         End
         Begin VB.TextBox txtCantidad 
            Height          =   315
            Left            =   5580
            MaxLength       =   10
            TabIndex        =   11
            Top             =   570
            Width           =   885
         End
         Begin VB.CommandButton cmdAsignar 
            Caption         =   "A&gregar"
            Height          =   420
            Left            =   7215
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Asignar Producto"
            Top             =   480
            Width           =   1020
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   330
            Left            =   6555
            MaskColor       =   &H000000FF&
            Picture         =   "frmEntradaProductos.frx":260A
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Buscar Producto"
            Top             =   570
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtcodigo 
            Height          =   315
            Left            =   105
            TabIndex        =   9
            Top             =   570
            Width           =   930
         End
         Begin VB.CommandButton cmdQuitar 
            Height          =   345
            Left            =   8670
            Picture         =   "frmEntradaProductos.frx":2914
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Quitar Producto"
            Top             =   1290
            Width           =   360
         End
         Begin MSFlexGridLib.MSFlexGrid GrdModulos 
            Height          =   2985
            Left            =   45
            TabIndex        =   18
            Top             =   1050
            Width           =   8580
            _ExtentX        =   15134
            _ExtentY        =   5265
            _Version        =   393216
            Cols            =   7
            FixedCols       =   0
            BackColorSel    =   8388736
            AllowBigSelection=   -1  'True
            FocusRect       =   0
            SelectionMode   =   1
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   51
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5655
            TabIndex        =   50
            ToolTipText     =   "Agregar Producto"
            Top             =   345
            Width           =   630
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1125
            TabIndex        =   49
            Top             =   330
            Width           =   840
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
         Height          =   1545
         Left            =   -74895
         TabIndex        =   29
         Top             =   555
         Width           =   9045
         Begin VB.CheckBox chkEncargado 
            Caption         =   "Encargado"
            Height          =   195
            Left            =   360
            TabIndex        =   20
            Top             =   795
            Width           =   1320
         End
         Begin VB.ComboBox cboEmpleado1 
            Height          =   315
            Left            =   3210
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   720
            Width           =   4005
         End
         Begin VB.ComboBox cboRep1 
            Height          =   315
            Left            =   3210
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   360
            Width           =   4005
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   360
            TabIndex        =   21
            Top             =   1080
            Width           =   810
         End
         Begin VB.CheckBox chkRepresentada 
            Caption         =   "Representada"
            Height          =   195
            Left            =   360
            TabIndex        =   19
            Top             =   510
            Width           =   1320
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1050
            Left            =   7830
            MaskColor       =   &H000000FF&
            Picture         =   "frmEntradaProductos.frx":3696
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Buscar Nota de Pedido"
            Top             =   345
            UseMaskColor    =   -1  'True
            Width           =   480
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   3210
            TabIndex        =   24
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   17170433
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   5760
            TabIndex        =   25
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   17170433
            CurrentDate     =   41098
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Encargado:"
            Height          =   195
            Left            =   2280
            TabIndex        =   55
            Top             =   750
            Width           =   825
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Representada:"
            Height          =   195
            Left            =   2025
            TabIndex        =   54
            Top             =   405
            Width           =   1050
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2085
            TabIndex        =   31
            Top             =   1110
            Width           =   1005
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4740
            TabIndex        =   30
            Top             =   1125
            Width           =   960
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3870
         Left            =   -74655
         TabIndex        =   32
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
      Begin MSFlexGridLib.MSFlexGrid GRDGrilla 
         Height          =   4365
         Left            =   -74925
         TabIndex        =   27
         Top             =   2160
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   7699
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   33
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Recepción"
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
      Left            =   3060
      TabIndex        =   47
      Top             =   7095
      Width           =   2475
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
      Left            =   210
      TabIndex        =   42
      Top             =   7035
      Width           =   750
   End
End
Attribute VB_Name = "frmEntradaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim resp As Integer
Dim I As Integer
Dim Linea As String
Dim RUBRO As String
Dim REPRE As String

Private Sub chkEncargado_Click()
    If chkEncargado.Value = Checked Then
        cboEmpleado1.Enabled = True
        cboEmpleado1.ListIndex = 0
    Else
        cboEmpleado1.Enabled = False
        cboEmpleado1.ListIndex = -1
    End If
End Sub

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        FechaDesde.Enabled = True
        FechaHasta.Enabled = True
    Else
        FechaDesde.Enabled = False
        FechaHasta.Enabled = False
        FechaDesde.Value = Null
        FechaHasta.Value = Null
    End If
End Sub

Private Sub chkRepresentada_Click()
    If chkRepresentada.Value = Checked Then
        cboRep1.Enabled = True
        cboRep1.ListIndex = 0
    Else
        cboRep1.Enabled = False
        cboRep1.ListIndex = -1
    End If
End Sub

Private Sub cmdAsignar_Click()
    If txtcodigo.Text <> "" Then
        GrdModulos.HighLight = flexHighlightAlways
        If txtCantidad <> "" Then
            If txtNumero.Text = "" Then
                For I = 1 To GrdModulos.Rows - 1
                    If GrdModulos.TextMatrix(I, 5) = CInt(txtcodigo.Text) Then
                        GrdModulos.TextMatrix(I, 4) = CDbl(GrdModulos.TextMatrix(I, 4)) + CDbl(txtCantidad.Text)
                        txtcodigo.Text = ""
                        txtcodigo.SetFocus
                        Exit Sub
                    End If
                Next
            Else
                For I = 1 To GrdModulos.Rows - 1
                    If GrdModulos.TextMatrix(I, 5) = CInt(txtcodigo.Text) Then
                        MsgBox "El producto ya fue ingresado", vbExclamation, TIT_MSGBOX
                        txtcodigo.SetFocus
                        Exit Sub
                    End If
                Next
            End If
            GrdModulos.AddItem txtdescri & Chr(9) & Linea & Chr(9) & _
                   RUBRO & Chr(9) & REPRE & Chr(9) & _
                   txtCantidad & Chr(9) & txtcodigo & Chr(9) & ""
             
            'txtIngNuevo_Click
            txtcodigo.Text = ""
            txtcodigo.SetFocus
        Else
            MsgBox "Debe Ingresar la cantidad", vbExclamation, TIT_MSGBOX
            txtCantidad.SetFocus
            Exit Sub
        End If
     Else
        MsgBox "Debe seleccionar un Producto"
    End If
End Sub

Private Sub Agregoproducto()
    GrdModulos.AddItem txtdescri & Chr(9) & "" & Chr(9) & _
                   "" & Chr(9) & "" & Chr(9) & _
                   txtCantidad & Chr(9) & txtcodigo
End Sub

Private Sub CmdBorrar_Click()
    If txtNumero.Text <> "" Then
        If GrdModulos.Rows <> 1 Then
            If MsgBox("Seguro desea Anular la Entrada de Productos Nº: " & XN(txtNumero.Text) & "? ", vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
                lblEstado.Caption = "Borrando..."
                Screen.MousePointer = vbHourglass
                On Error GoTo HayError1
                DBConn.BeginTrans
                'ANULO LA ENTRADA
                sql = "UPDATE ENTRADA_PRODUCTO"
                sql = sql & " SET EST_CODIGO=2"
                sql = sql & " WHERE EPR_CODIGO=" & XN(txtNumero)
                
                'ACTUALIZO EL DETALLE
                For I = 1 To GrdModulos.Rows - 1
                    sql = "UPDATE DETALLE_STOCK"
                    sql = sql & " SET DST_STKFIS = DST_STKFIS - " & XN(GrdModulos.TextMatrix(I, 4)) & ""
                    sql = sql & " WHERE STK_CODIGO = " & XN(cboStock.ItemData(cboStock.ListIndex)) & "  "
                    sql = sql & " AND PTO_CODIGO = " & XN(GrdModulos.TextMatrix(I, 5)) & ""
                    DBConn.Execute sql
                Next
                DBConn.CommitTrans
            End If
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            CmdNuevo_Click
        End If
    End If
  Exit Sub
HayError1:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Set rec = New ADODB.Recordset
    sql = "SELECT E.EPR_CODIGO, E.EPR_FECHA, V.VEN_NOMBRE"
    sql = sql & " FROM ENTRADA_PRODUCTO E,VENDEDOR V"
    sql = sql & " WHERE E.VEN_CODIGO = V.VEN_CODIGO"
    If chkRepresentada.Value = Checked Then sql = sql & " AND E.REP_CODIGO = " & XN(cboRep1.ItemData(cboRep1.ListIndex))
    If chkEncargado.Value = Checked Then sql = sql & " AND E.VEN_CODIGO = " & XN(cboEmpleado1.ItemData(cboEmpleado1.ListIndex))
    If chkFecha.Value = Checked Then
        If Not IsNull(FechaDesde) Then sql = sql & " AND E.EPR_FECHA>=" & XDQ(FechaDesde)
        If Not IsNull(FechaHasta) Then sql = sql & " AND E.EPR_FECHA<=" & XDQ(FechaHasta)
    End If
    sql = sql & " ORDER BY E.EPR_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
      
    If rec.EOF = False Then
        grdGrilla.Rows = 1
        Do While rec.EOF = False
            grdGrilla.AddItem Format(rec!EPR_CODIGO, "00000000") & Chr(9) & rec!EPR_FECHA & Chr(9) & _
                              rec!VEN_NOMBRE
            rec.MoveNext
        Loop
        grdGrilla.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
    End If
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    rec.Close
End Sub

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 2
    frmBuscar.Show vbModal
    
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtcodigo.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtdescri.Text = frmBuscar.grdBuscar.Text
        txtCantidad.SetFocus
    Else
        txtcodigo.SetFocus
    End If
    
    End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo HayError2
         
    If ValidarEntrada = False Then Exit Sub
           
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Guardando ..."
        DBConn.BeginTrans
        
        sql = "SELECT EPR_FECHA FROM ENTRADA_PRODUCTO"
        sql = sql & " WHERE EPR_CODIGO = " & XN(txtNumero.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = True Then
           'INSERTO EN LA TABLA ENTRADA_PRODUCTO
           sql = "INSERT INTO ENTRADA_PRODUCTO(EPR_CODIGO,EPR_FECHA,VEN_CODIGO,REP_CODIGO,"
           sql = sql & "EPR_NROSUCREM,EPR_NROREM,EPR_FECHAREM,STK_CODIGO,TRS_CODIGO,EPR_OBSERVACIONES)"
           sql = sql & " VALUES ("
           sql = sql & XN(txtNumero) & ","
           sql = sql & XDQ(Fecha) & ","
           sql = sql & XN(cboEmpleado.ItemData(cboEmpleado.ListIndex)) & ","
           If cboRep.List(cboRep.ListIndex) <> "<Ninguna>" Then
                sql = sql & XN(cboRep.ItemData(cboRep.ListIndex)) & ","
                sql = sql & XN(txtNroSucursal.Text) & ","
                sql = sql & XN(txtRemito) & ","
                sql = sql & XDQ(fecRemito.Value) & ","
           Else
                sql = sql & "NULL,NULL,NULL,NULL,"
           End If
           sql = sql & XN(cboStock.ItemData(cboStock.ListIndex)) & ","
           If cboTransporte.List(cboTransporte.ListIndex) <> "<Ninguno>" Then
                sql = sql & XN(cboTransporte.ItemData(cboTransporte.ListIndex)) & ","
           Else
                sql = sql & "NULL,"
           End If
           sql = sql & XS(txtObservaciones.Text) & ")"
           
           DBConn.Execute sql
           
           'INSERTO EN LA TABLA DETALLE_ENTRADA_PRODUCTO
           For I = 1 To GrdModulos.Rows - 1
               sql = "INSERT INTO DETALLE_ENTRADA_PRODUCTO(EPR_CODIGO,PTO_CODIGO,DEP_CANTIDAD)"
               sql = sql & " VALUES ("
               sql = sql & XN(txtNumero) & ","
               sql = sql & XN(GrdModulos.TextMatrix(I, 5)) & ","
               sql = sql & XN(GrdModulos.TextMatrix(I, 4)) & " )"
               DBConn.Execute sql
           Next
    
            'ACTUALIZO DETALLE_STOCK
            For I = 1 To GrdModulos.Rows - 1
                sql = "UPDATE DETALLE_STOCK"
                sql = sql & " SET DST_STKFIS = DST_STKFIS + " & XN(GrdModulos.TextMatrix(I, 4))
                sql = sql & " WHERE STK_CODIGO= " & XN(cboStock.ItemData(cboStock.ListIndex))
                sql = sql & " AND PTO_CODIGO =" & XN(GrdModulos.TextMatrix(I, 5))
                DBConn.Execute sql
            Next
            
            'ACTUALIZO LA TABLA PARAMENTROS
            sql = "UPDATE PARAMETROS SET RECEPCION_MERCADERIA=" & XN(txtNumero)
            DBConn.Execute sql
        Else
            MsgBox "La Recepción de Mercadería ya fue registrada", vbCritical, TIT_MSGBOX
        End If
        rec.Close
        Screen.MousePointer = vbNormal
        lblEstado.Caption = ""
        DBConn.CommitTrans
        CmdNuevo_Click
    Exit Sub
         
HayError2:
         lblEstado.Caption = ""
         DBConn.RollbackTrans
         Screen.MousePointer = vbNormal
         MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Function ValidarEntrada()
    If cboEmpleado.ListIndex = -1 Then
        MsgBox "No ha ingresado el Encargado de Depósito", vbExclamation, TIT_MSGBOX
        cboEmpleado.SetFocus
        ValidarEntrada = False
        Exit Function
    End If
    If IsNull(Fecha.Value) Then
        MsgBox "No ha ingresado la Fecha de Entrada de Productos", vbExclamation, TIT_MSGBOX
        Fecha.SetFocus
        ValidarEntrada = False
        Exit Function
    End If
    If cboRep.List(cboRep.ListIndex) <> "<Ninguna>" And txtRemito.Text = "" Then
        MsgBox "No ha ingresado el Nº de Remito ", vbExclamation, TIT_MSGBOX
        txtRemito.SetFocus
        ValidarEntrada = False
        Exit Function
    End If
    If cboRep.List(cboRep.ListIndex) = "<Ninguna>" And txtRemito.Text = "" And txtObservaciones.Text = "" Then
        MsgBox "Debe ingresar una Observación (para saber por que se realizó la entrada)", vbExclamation, TIT_MSGBOX
        txtObservaciones.SetFocus
        ValidarEntrada = False
        Exit Function
    End If
    If GrdModulos.Rows = 1 Then
        MsgBox "Debe haber ingresar al menos un producto en la Grilla ", vbExclamation, TIT_MSGBOX
        cmdAsignar.SetFocus
        ValidarEntrada = False
        Exit Function
    End If
    ValidarEntrada = True
End Function

Private Sub CmdNuevo_Click()
    txtNumero.Text = ""
    txtNroSucursal.Text = ""
    txtObservaciones.Text = ""
    txtRemito.Text = ""
    fecRemito.Value = Null
    cboEmpleado.ListIndex = 0
    cboRep.ListIndex = 0
    cboTransporte.ListIndex = 0
    Fecha.Value = Date
    txtcodigo.Text = ""
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    Call BuscoEstado(1, lblEstadoRecepcion)
    tabDatos.Tab = 0
    BuscoNumeroRecepcion
    CmdBorrar.Enabled = False
    cmdGrabar.Enabled = True
    cboStock.SetFocus
End Sub

Private Sub cmdQuitar_Click()
    If GrdModulos.Rows <> 1 Then
        resp = MsgBox("Seguro desea eliminar el Producto: " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 0)) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        lblEstado.Caption = "Borrando..."
        Screen.MousePointer = vbHourglass
        If GrdModulos.Rows = 2 Then
            GrdModulos.HighLight = flexHighlightNever
            GrdModulos.Rows = 1
            txtcodigo.SetFocus
        Else
            GrdModulos.RemoveItem (GrdModulos.RowSel)
            txtcodigo.SetFocus
        End If
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
    End If
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmEntradaProductos = Nothing
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
    Set Rec1 = New ADODB.Recordset
    lblEstado.Caption = ""
    
    Call Centrar_pantalla(Me)
    preparogrilla
    'CARGO COMBO EMPLEADO
    cargocboEmpl
    'CARGO COMBO STOCK
    CargocboStock
    'CARGO COMBO REPRESENTADA
    CargoComboRepresentada
    'CARGO COMBO TRANSPORTE
    CargoComboTransporte
    tabDatos.Tab = 0
    cmdAsignar.Enabled = False
    CmdBorrar.Enabled = False
    GrdModulos.HighLight = flexHighlightNever
    'BUSCO NUMERO DE RECEPCION DE MERCADERIA
    BuscoNumeroRecepcion
    Call BuscoEstado(1, lblEstadoRecepcion)
    Fecha.Value = Date
    Linea = ""
    RUBRO = ""
    REPRE = ""
End Sub

Private Sub BuscoNumeroRecepcion()
    sql = "SELECT (RECEPCION_MERCADERIA + 1) AS NUMERO_REP FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtNumero.Text = Format(rec!NUMERO_REP, "00000000")
    End If
    rec.Close
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

Private Sub CargoComboRepresentada()
    sql = "SELECT REP_RAZSOC,REP_CODIGO FROM REPRESENTADA ORDER BY REP_RAZSOC"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboRep.AddItem "<Ninguna>"
        Do While rec.EOF = False
            cboRep.AddItem rec!REP_RAZSOC
            cboRep.ItemData(cboRep.NewIndex) = rec!REP_CODIGO
            cboRep1.AddItem rec!REP_RAZSOC
            cboRep1.ItemData(cboRep1.NewIndex) = rec!REP_CODIGO
            rec.MoveNext
        Loop
        cboRep.ListIndex = 0
        cboRep1.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub preparogrilla()
    'GRILLA DONDE SE CRAGAN LOS PRODUCTOS
    GrdModulos.FormatString = "Producto|Linea|Rubro|Representada|Cantidad|CODPROD|marca"
    GrdModulos.ColWidth(0) = 2700 'Producto
    GrdModulos.ColWidth(1) = 1500 'Linea
    GrdModulos.ColWidth(2) = 1500 'Rubro
    GrdModulos.ColWidth(3) = 1900 'Representada
    GrdModulos.ColWidth(4) = 800  'Cantidad
    GrdModulos.ColWidth(5) = 0    'CODPROD
    GrdModulos.ColWidth(6) = 0    'marca para saber cunado actualizo el stock
    'X para cuando lo recupero de la tabla y tengo que modificarlo
    '"" para cuando no lo recupero de la base
    GrdModulos.Rows = 1
    'GRILLA PARA LA BUSQUEDA
    grdGrilla.FormatString = "^Numero|^Fecha|Encargado"
    grdGrilla.ColWidth(0) = 1200 'NUMERO
    grdGrilla.ColWidth(1) = 1300 'FECHA
    grdGrilla.ColWidth(2) = 5000 'EMPLEADO
    grdGrilla.Rows = 1
End Sub

Private Sub cargocboEmpl()
    sql = "SELECT * FROM VENDEDOR ORDER BY VEN_NOMBRE"
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

Private Sub GRDGrilla_DblClick()
    If grdGrilla.Rows > 1 Then
        CmdNuevo_Click
        txtNumero.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 0)
        Fecha.Value = grdGrilla.TextMatrix(grdGrilla.RowSel, 1)
        txtNumero_LostFocus
        tabDatos.Tab = 0
    End If
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then GRDGrilla_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 1 Then
      FechaDesde.Enabled = False
      FechaHasta.Enabled = False
      cboRep1.Enabled = False
      cboRep1.ListIndex = -1
      cboEmpleado1.Enabled = False
      cboEmpleado1.ListIndex = -1
      cmdGrabar.Enabled = False
      CmdBorrar.Enabled = False
      LimpiarBusqueda
      If Me.Visible = True Then chkRepresentada.SetFocus
    Else
      cmdGrabar.Enabled = True
      CmdBorrar.Enabled = True
    End If
End Sub

Private Sub LimpiarBusqueda()
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    grdGrilla.Rows = 1
    chkFecha.Value = Unchecked
    chkRepresentada.Value = Unchecked
    chkEncargado.Value = Unchecked
End Sub

Private Sub txtCantidad_GotFocus()
    SelecTexto txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_Change()
    If txtcodigo.Text = "" Then
        txtcodigo.Text = ""
        txtdescri.Text = ""
        txtCantidad.Text = ""
        Linea = ""
        RUBRO = ""
        REPRE = ""
        cmdAsignar.Enabled = False
    Else
        cmdAsignar.Enabled = True
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto txtcodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If txtcodigo.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = " SELECT P.PTO_DESCRI,L.LNA_DESCRI, "
        sql = sql & " R.RUB_DESCRI,RE.REP_RAZSOC,P.PTO_CODIGO"
        sql = sql & " FROM PRODUCTO P,LINEAS L,RUBROS R,REPRESENTADA RE,DETALLE_STOCK DS"
        sql = sql & " WHERE P.LNA_CODIGO = L.LNA_CODIGO"
        sql = sql & " AND P.RUB_CODIGO = R.RUB_CODIGO"
        sql = sql & " AND P.REP_CODIGO = RE.REP_CODIGO"
        sql = sql & " AND DS.PTO_CODIGO = P.PTO_CODIGO"
        sql = sql & " AND DS.STK_CODIGO = " & XN(cboStock.ItemData(cboStock.ListIndex))
        sql = sql & " AND P.PTO_CODIGO = " & XN(txtcodigo.Text)
        sql = sql & " ORDER BY P.PTO_CODIGO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtdescri.Text = rec!PTO_DESCRI
            Linea = rec!LNA_DESCRI
            RUBRO = rec!RUB_DESCRI
            REPRE = rec!REP_RAZSOC
        Else
            MsgBox "El Código no existe, o no pertenece al stock de " & cboStock.Text & "", vbExclamation, TIT_MSGBOX
            Linea = ""
            RUBRO = ""
            REPRE = ""
            txtcodigo.SetFocus
            
        End If
        rec.Close
    End If
End Sub

Private Sub CargocboStock()
    sql = "SELECT S.STK_CODIGO,R.REP_RAZSOC FROM STOCK S, REPRESENTADA R "
    sql = sql & " WHERE S.REP_CODIGO = R.REP_CODIGO"
    sql = sql & " ORDER BY S.STK_CODIGO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboStock.AddItem rec!REP_RAZSOC
            cboStock.ItemData(cboStock.NewIndex) = rec!STK_CODIGO
            rec.MoveNext
        Loop
        cboStock.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub txtDescri_Change()
    If txtdescri.Text = "" Then
        txtcodigo.Text = ""
    End If
End Sub

Private Sub txtdescri_GotFocus()
    SelecTexto txtdescri
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_LostFocus()
           
   If txtcodigo.Text = "" And txtdescri.Text <> "" Then
        Set rec = New ADODB.Recordset
        Screen.MousePointer = vbHourglass
        sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI,RE.REP_RAZSOC"
        sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L,REPRESENTADA RE"
        sql = sql & " WHERE P.RUB_CODIGO = R.RUB_CODIGO"
        sql = sql & " AND P.LNA_CODIGO = L.LNA_CODIGO AND L.LNA_CODIGO = R.LNA_CODIGO"
        sql = sql & " AND RE.REP_CODIGO=P.REP_CODIGO"
        sql = sql & " AND P.PTO_DESCRI LIKE '" & txtdescri.Text & "%'ORDER BY P.PTO_DESCRI"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                'grdGrilla.SetFocus
                frmBuscar.TipoBusqueda = 2
                frmBuscar.CodListaPrecio = 0
                frmBuscar.TxtDescriB.Text = txtdescri.Text
                frmBuscar.Show vbModal
                frmBuscar.grdBuscar.Col = 0
                txtcodigo.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
                frmBuscar.grdBuscar.Col = 1
                txtdescri.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
                frmBuscar.grdBuscar.Col = 2
                Linea = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2)
                frmBuscar.grdBuscar.Col = 3
                RUBRO = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
                frmBuscar.grdBuscar.Col = 4
                REPRE = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)
            Else
                txtcodigo.Text = Trim(rec!PTO_CODIGO)
                txtdescri.Text = Trim(rec!PTO_DESCRI)
                Linea = rec!LNA_DESCRI
                RUBRO = rec!RUB_DESCRI
                REPRE = rec!REP_RAZSOC
            End If
                
            
        Else
                MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
                txtdescri.Text = ""
        End If
        rec.Close
        Screen.MousePointer = vbNormal
    ElseIf txtcodigo.Text = "" And txtdescri.Text = "" Then
        MsgBox "Debe ingresar un Producto", vbExclamation, TIT_MSGBOX
        txtcodigo.SetFocus
    End If
    
End Sub

Private Sub txtNroSucursal_GotFocus()
    SelecTexto txtNroSucursal
End Sub

Private Sub txtNroSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroSucursal_LostFocus()
    If txtNroSucursal.Text <> "" Then
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    End If
End Sub

Private Sub txtNumero_GotFocus()
    SelecTexto txtNumero
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNumero_LostFocus()
    If txtNumero.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        sql = "SELECT * FROM ENTRADA_PRODUCTO"
        sql = sql & " WHERE EPR_CODIGO=" & XN(txtNumero)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Fecha.Value = Rec1!EPR_FECHA
            Call BuscaCodigoProxItemData(Rec1!VEN_CODIGO, cboEmpleado)
            Call BuscaCodigoProxItemData(Rec1!STK_CODIGO, cboStock)
            If Not IsNull(Rec1!REP_CODIGO) Then
               txtNroSucursal.Text = Format(Rec1!EPR_NROSUCREM, "0000")
               txtRemito.Text = Format(Rec1!EPR_NROREM, "00000000")
               fecRemito.Value = Rec1!EPR_FECHAREM
               Call BuscaCodigoProxItemData(Rec1!REP_CODIGO, cboRep)
            Else
               cboRep.ListIndex = 0
            End If
            If Not IsNull(Rec1!TRS_CODIGO) Then
               Call BuscaCodigoProxItemData(Rec1!TRS_CODIGO, cboTransporte)
            Else
               cboTransporte.ListIndex = 0
            End If
            CargoGrilla (txtNumero)
            Call BuscoEstado(CInt(Rec1!EST_CODIGO), lblEstadoRecepcion)
            txtObservaciones.Text = ChkNull(Rec1!EPR_OBSERVACIONES)
            If Rec1!EST_CODIGO = 2 Then
               CmdBorrar.Enabled = False
            Else
               CmdBorrar.Enabled = True
            End If
            cmdGrabar.Enabled = False
        Else
            MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
            CmdNuevo_Click
            cboStock.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub CargoGrilla(Campo As Integer)
    
    Screen.MousePointer = vbHourglass
    sql = "SELECT DISTINCT  P.PTO_DESCRI,L.LNA_DESCRI,R.RUB_DESCRI,"
    sql = sql & " RE.REP_RAZSOC,D.DEP_CANTIDAD, E.EPR_CODIGO, E.EPR_FECHA,P.PTO_CODIGO"
    sql = sql & " FROM ENTRADA_PRODUCTO E,PRODUCTO P,LINEAS L,RUBROS R, REPRESENTADA RE,DETALLE_ENTRADA_PRODUCTO D "
    sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO AND D.EPR_CODIGO = E.EPR_CODIGO "
    sql = sql & " AND L.LNA_CODIGO = P.LNA_CODIGO "
    sql = sql & " AND R.RUB_CODIGO = P.RUB_CODIGO  AND P.REP_CODIGO = RE.REP_CODIGO "
    sql = sql & " AND E.EPR_CODIGO = " & Campo & " ORDER BY E.EPR_CODIGO"
        
    lblEstado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        GrdModulos.Rows = 1
        GrdModulos.HighLight = flexHighlightAlways
        Do While Not rec.EOF
           GrdModulos.AddItem rec!PTO_DESCRI & Chr(9) & rec!LNA_DESCRI & Chr(9) & _
                              rec!RUB_DESCRI & Chr(9) & rec!REP_RAZSOC & Chr(9) & _
                              rec!DEP_CANTIDAD & Chr(9) & rec!PTO_CODIGO & Chr(9) & "X"
    
            rec.MoveNext
        Loop
        rec.MoveFirst
    Else
        lblEstado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        Me.txtNumero.SetFocus
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtRemito_GotFocus()
    SelecTexto txtRemito
End Sub

Private Sub txtRemito_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtRemito_LostFocus()
    If txtRemito.Text <> "" Then
        txtRemito.Text = Format(txtRemito.Text, "00000000")
    End If
End Sub
