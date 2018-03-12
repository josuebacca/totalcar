VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ABMProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM de Productos"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopiar 
      Caption         =   "&Copiar Desde"
      DisabledPicture =   "ABMProducto.frx":0000
      Height          =   750
      Left            =   2400
      Picture         =   "ABMProducto.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5295
      Width           =   870
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMProducto.frx":0694
      Height          =   750
      Left            =   5070
      Picture         =   "ABMProducto.frx":099E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5295
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   5175
      Left            =   120
      TabIndex        =   39
      Top             =   60
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
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
      TabPicture(0)   =   "ABMProducto.frx":0CA8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tabdatos2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDatos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMProducto.frx":0CC4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraDatos 
         Caption         =   " Datos del Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   240
         TabIndex        =   64
         Top             =   360
         Width           =   6390
         Begin VB.CheckBox chkBaja 
            Caption         =   "Dar de Baja"
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
            Left            =   5160
            TabIndex        =   71
            Top             =   5040
            Width           =   1620
         End
         Begin VB.TextBox txtcodrepu 
            Height          =   285
            Left            =   5160
            TabIndex        =   1
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtpCpra 
            Height          =   300
            Left            =   4470
            MaxLength       =   10
            TabIndex        =   7
            Top             =   3090
            Width           =   1020
         End
         Begin VB.CommandButton cmdNuevaPresentacion 
            Height          =   315
            Left            =   5715
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProducto.frx":0CE0
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Agregar Presentacion"
            Top             =   2565
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaLinea 
            Height          =   315
            Left            =   5715
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProducto.frx":106A
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Agregar Línea"
            Top             =   1755
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoRubro 
            Height          =   315
            Left            =   5715
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProducto.frx":13F4
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Agregar Rubro"
            Top             =   2160
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtStock 
            Height          =   300
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   8
            Top             =   3577
            Width           =   1020
         End
         Begin VB.TextBox txtPrecio 
            Height          =   300
            Left            =   1170
            MaxLength       =   10
            TabIndex        =   6
            Top             =   3090
            Width           =   1020
         End
         Begin VB.ComboBox cboPres 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2580
            Width           =   4350
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Left            =   1170
            MaxLength       =   16
            TabIndex        =   0
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtDescri 
            Height          =   900
            Left            =   1170
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   2
            Tag             =   "Descripción"
            Top             =   720
            Width           =   5070
         End
         Begin VB.ComboBox cboLinea 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1755
            Width           =   4350
         End
         Begin VB.ComboBox cboRubro 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2160
            Width           =   4350
         End
         Begin VB.Frame Frame2 
            Caption         =   "Lista de Precios"
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   240
            TabIndex        =   65
            Top             =   3960
            Width           =   6015
            Begin VB.ComboBox cboListaPrecio 
               Height          =   315
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   240
               Width           =   3705
            End
            Begin VB.CommandButton cmdLisPrecio 
               Height          =   315
               Left            =   5520
               MaskColor       =   &H000000FF&
               Picture         =   "ABMProducto.frx":177E
               Style           =   1  'Graphical
               TabIndex        =   66
               ToolTipText     =   "Agregar Presentacion"
               Top             =   240
               UseMaskColor    =   -1  'True
               Width           =   405
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "Lista de Precios:"
               Height          =   195
               Left            =   360
               TabIndex        =   67
               Top             =   270
               Width           =   1170
            End
         End
         Begin VB.TextBox txtSFisico 
            Height          =   300
            Left            =   4470
            MaxLength       =   50
            TabIndex        =   9
            Top             =   3577
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Precio Compra:"
            Height          =   195
            Left            =   3270
            TabIndex        =   81
            Top             =   3150
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Ref:"
            Height          =   195
            Left            =   4440
            TabIndex        =   80
            Top             =   360
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Marca:"
            Height          =   195
            Left            =   585
            TabIndex        =   79
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   78
            Top             =   360
            Width           =   660
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   195
            TabIndex        =   77
            Top             =   780
            Width           =   885
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Línea:"
            Height          =   195
            Left            =   615
            TabIndex        =   76
            Top             =   1800
            Width           =   465
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Rubro:"
            Height          =   195
            Left            =   600
            TabIndex        =   75
            Top             =   2220
            Width           =   480
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Precio Vta:"
            Height          =   195
            Left            =   300
            TabIndex        =   74
            Top             =   3143
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Stock Mín.:"
            Height          =   195
            Left            =   240
            TabIndex        =   73
            Top             =   3630
            Width           =   840
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Stock Fisico:"
            Height          =   195
            Left            =   3435
            TabIndex        =   72
            Top             =   3630
            Width           =   915
         End
      End
      Begin TabDlg.SSTab tabdatos2 
         Height          =   5295
         Left            =   3360
         TabIndex        =   44
         Top             =   360
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   9340
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Principales"
         TabPicture(0)   =   "ABMProducto.frx":1B08
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Se&cundarios"
         TabPicture(1)   =   "ABMProducto.frx":1B24
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frameTractor"
         Tab(1).ControlCount=   1
         Begin VB.Frame frameTractor 
            Caption         =   "Datos del Tractor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   4095
            Left            =   -74640
            TabIndex        =   45
            Top             =   480
            Width           =   6135
            Begin VB.TextBox mtxttipo 
               Height          =   285
               Left            =   1440
               TabIndex        =   16
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox mtxttipmod 
               Height          =   285
               Left            =   4200
               TabIndex        =   17
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox mtxttracci 
               Height          =   285
               Left            =   1440
               TabIndex        =   18
               Top             =   705
               Width           =   1335
            End
            Begin VB.CheckBox chkcabina 
               Caption         =   "Cabina"
               Height          =   255
               Left            =   4200
               TabIndex        =   19
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox mtxtmotmar 
               Height          =   285
               Left            =   1440
               TabIndex        =   20
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox mtxtmotmod 
               Height          =   285
               Left            =   4200
               TabIndex        =   21
               Top             =   1080
               Width           =   1335
            End
            Begin VB.CheckBox chkkitcon 
               Caption         =   "Kit Confort"
               Height          =   195
               Left            =   1440
               TabIndex        =   30
               Top             =   2925
               Width           =   1215
            End
            Begin VB.TextBox mtxtaspira 
               Height          =   285
               Left            =   1440
               TabIndex        =   22
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox mtxtchasis 
               Height          =   285
               Left            =   1440
               TabIndex        =   24
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox mtxtneumde 
               Height          =   285
               Left            =   1440
               TabIndex        =   26
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox mtxtmotnro 
               Height          =   285
               Left            =   4200
               TabIndex        =   23
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox mtxtserie 
               Height          =   285
               Left            =   4200
               TabIndex        =   25
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox mtxtnedeca 
               Height          =   285
               Left            =   4200
               TabIndex        =   27
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox mtxtnetrca 
               Height          =   285
               Left            =   4200
               TabIndex        =   29
               Top             =   2520
               Width           =   1335
            End
            Begin VB.TextBox mtxtsalhid 
               Height          =   285
               Left            =   4200
               TabIndex        =   31
               Top             =   2880
               Width           =   1335
            End
            Begin VB.TextBox mtxtcerfab 
               Height          =   285
               Left            =   4200
               TabIndex        =   33
               Top             =   3240
               Width           =   1335
            End
            Begin VB.TextBox mtxtposara 
               Height          =   285
               Left            =   1440
               TabIndex        =   32
               Top             =   3240
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox mtxtopcional1 
               Height          =   285
               Left            =   1440
               TabIndex        =   34
               Top             =   3600
               Width           =   1335
            End
            Begin VB.TextBox mtxtneumtr 
               Height          =   285
               Left            =   1440
               TabIndex        =   28
               Top             =   2520
               Width           =   1335
            End
            Begin VB.TextBox mtxtopcional2 
               Height          =   285
               Left            =   4200
               TabIndex        =   35
               Top             =   3600
               Width           =   1335
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   1035
               TabIndex        =   63
               Top             =   405
               Width           =   360
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Tracción:"
               Height          =   195
               Left            =   720
               TabIndex        =   62
               Top             =   750
               Width           =   675
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Modelo:"
               Height          =   195
               Left            =   3495
               TabIndex        =   61
               Top             =   405
               Width           =   570
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Motor Marca:"
               Height          =   195
               Left            =   450
               TabIndex        =   60
               Top             =   1125
               Width           =   945
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Modelo:"
               Height          =   195
               Left            =   3495
               TabIndex        =   59
               Top             =   1125
               Width           =   570
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Aspiración:"
               Height          =   195
               Left            =   615
               TabIndex        =   58
               Top             =   1485
               Width           =   780
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Motor Nro.:"
               Height          =   195
               Left            =   3270
               TabIndex        =   57
               Top             =   1485
               Width           =   795
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Chasis Nro.:"
               Height          =   195
               Left            =   540
               TabIndex        =   56
               Top             =   1845
               Width           =   855
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Serie:"
               Height          =   195
               Left            =   3660
               TabIndex        =   55
               Top             =   1845
               Width           =   405
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Neum. Delantero:"
               Height          =   195
               Left            =   150
               TabIndex        =   54
               Top             =   2205
               Width           =   1245
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Neum. Trasero:"
               Height          =   195
               Left            =   300
               TabIndex        =   53
               Top             =   2520
               Width           =   1095
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Cantidad:"
               Height          =   195
               Left            =   3390
               TabIndex        =   52
               Top             =   2205
               Width           =   675
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Cantidad:"
               Height          =   195
               Left            =   3390
               TabIndex        =   51
               Top             =   2520
               Width           =   675
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Salida Hid."
               Height          =   195
               Left            =   3300
               TabIndex        =   50
               Top             =   2925
               Width           =   765
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Posic. Aran:"
               Height          =   195
               Left            =   540
               TabIndex        =   49
               Top             =   3285
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Cert. Fabrica:"
               Height          =   195
               Left            =   3120
               TabIndex        =   48
               Top             =   3285
               Width           =   945
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Opcional1:"
               Height          =   195
               Left            =   630
               TabIndex        =   47
               Top             =   3645
               Width           =   765
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Opcional2:"
               Height          =   195
               Left            =   3300
               TabIndex        =   46
               Top             =   3645
               Width           =   765
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74865
         TabIndex        =   40
         Top             =   600
         Width           =   6420
         Begin VB.TextBox TxtDescriB 
            Height          =   330
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   36
            Top             =   240
            Width           =   4410
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   360
            Left            =   5880
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProducto.frx":1B40
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Producto:"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   41
            Top             =   270
            Width           =   690
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   3645
         Left            =   -74880
         TabIndex        =   38
         Top             =   1410
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   6429
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   42
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMProducto.frx":1E4A
      Height          =   750
      Left            =   4185
      Picture         =   "ABMProducto.frx":2154
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5295
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMProducto.frx":245E
      Height          =   750
      Left            =   5970
      Picture         =   "ABMProducto.frx":2768
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5295
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMProducto.frx":2A72
      Height          =   750
      Left            =   3285
      Picture         =   "ABMProducto.frx":2D7C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5295
      Width           =   870
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
      Left            =   120
      TabIndex        =   43
      Top             =   5295
      Width           =   750
   End
End
Attribute VB_Name = "ABMProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As ADODB.Recordset
Dim Rec1 As ADODB.Recordset

Dim sql As String
Dim resp As Integer
Public CODIGOLISTA As Integer


Private Sub cboLinea_Click()
    cborubro.Clear
    cboPres.Clear
End Sub

Private Sub cboLinea_LostFocus()
    cargocboRubro
End Sub

Private Sub cboRubro_Click()
    cboPres.Clear
End Sub

Private Sub cboRubro_LostFocus()
    cargocbotiporep
End Sub



Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(txtcodigo) <> "" Then
        resp = MsgBox("Seguro desea eliminar el Producto: " & Trim(txtdescri.Text) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Eliminando ..."
        
        DBConn.Execute "DELETE FROM PRODUCTO WHERE PTO_CODIGO LIKE '" & txtcodigo & "'"
        DBConn.Execute "DELETE FROM DETALLE_STOCK WHERE PTO_CODIGO LIKE '" & txtcodigo & "'"
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        CmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    MousePointer = vbHourglass
    TxtDescriB.Text = Replace(TxtDescriB, "'", "´")
    sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, L.LNA_DESCRI, R.RUB_DESCRI"
    sql = sql & " FROM PRODUCTO P, LINEAS L, RUBROS R"
    sql = sql & " WHERE"
    sql = sql & " P.RUB_CODIGO = R.RUB_CODIGO AND P.LNA_CODIGO=L.LNA_CODIGO"
    sql = sql & " AND R.LNA_CODIGO=L.LNA_CODIGO"
    sql = sql & " AND (PTO_DESCRI LIKE '" & TxtDescriB.Text & "%' "
    sql = sql & " OR PTO_CODIGO LIKE '" & TxtDescriB.Text & "%' )"
    sql = sql & " ORDER BY PTO_DESCRI"
        
    lblEstado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Do While Not rec.EOF
           GrdModulos.AddItem rec!PTO_CODIGO & Chr(9) & Trim(rec!PTO_DESCRI) & Chr(9) & _
                              rec!RUB_DESCRI & Chr(9) & rec!LNA_DESCRI
           rec.MoveNext
        Loop
        If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        TxtDescriB.SetFocus
    End If
    rec.Close
    MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub cmdCopiar_Click()
    If GrdModulos.Rows > 1 Then
        txtcodigo = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        cmdGrabar.Enabled = True
        CmdBorrar.Enabled = True
        TxtCodigo_LostFocus
        tabDatos.Tab = 0
        tabdatos2.Tab = 0
        txtcodigo = ""
    End If
End Sub

Private Sub cmdGrabar_Click()
    Dim PRECIVA As String
    If validarProuducto = False Then Exit Sub
    
    On Error GoTo HayError
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    DBConn.BeginTrans
    sql = "SELECT * FROM PRODUCTO WHERE PTO_CODIGO LIKE '" & txtcodigo.Text & "'"
    If rec.State = 1 Then
        rec.Close
    End If
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If cbolinea.ItemData(cbolinea.ListIndex) = 6 Then
            PRECIVA = txtPrecio * 1.105
    Else
        If cbolinea.ItemData(cbolinea.ListIndex) = 7 Then
            PRECIVA = txtPrecio * 1.21
        End If
    End If
    
    If rec.EOF = False Then
        
        sql = "UPDATE PRODUCTO "
        sql = sql & " SET PTO_DESCRI=" & XS(txtdescri, True)
        sql = sql & " , LNA_CODIGO=" & cbolinea.ItemData(cbolinea.ListIndex)
        sql = sql & " , RUB_CODIGO=" & cborubro.ItemData(cborubro.ListIndex)
        sql = sql & " , TPRE_CODIGO=" & cboPres.ItemData(cboPres.ListIndex)
        sql = sql & " , PTO_PRECIO=" & XN(txtPrecio)
        'sql = sql & " , PTO_PRECIOC=" & XN(txtpCpra)
        sql = sql & " , PTO_PRECIVA=" & XN(PRECIVA)
        sql = sql & " , PTO_STKMIN=" & XN(txtStock)
        If cboListaPrecio.ListIndex > 0 Then
            sql = sql & " , LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
        Else
            sql = sql & " , LIS_CODIGO=" & 0
        End If
        If chkBaja.Value = Checked Then
            sql = sql & " ,PTO_ESTADO=2"
        Else
            sql = sql & " ,PTO_ESTADO=1"
        End If
'        sql = sql & " , PTO_TIPO=" & XS(mtxttipo)
'        sql = sql & " , PTO_TIPMOD=" & XS(mtxttipmod)
'        sql = sql & " , PTO_TRACCI=" & XS(mtxttracci)
'        'If chkcabina.Value = 1 Then
'            sql = sql & " , PTO_CABINA=" & chkcabina.Value
'        'Else
'        '    sql = sql & " , PTO_CABINA=" & XN("0")
'        'End If
'        sql = sql & " , PTO_MOTMAR=" & XS(mtxtmotmar)
'        sql = sql & " , PTO_MOTMOD=" & XS(mtxtmotmod)
'        sql = sql & " , PTO_ASPIRA=" & XS(mtxtaspira)
'        sql = sql & " , PTO_MOTNRO=" & XS(mtxtmotnro)
'        sql = sql & " , PTO_CHASIS=" & XS(mtxtchasis)
'        sql = sql & " , PTO_SERIE=" & XS(mtxtserie)
'        sql = sql & " , PTO_NEUMDE=" & XS(mtxtneumde)
'        sql = sql & " , PTO_NEDECA=" & XN(mtxtnedeca)
'        sql = sql & " , PTO_NEUMTR=" & XS(mtxtneumtr)
'        sql = sql & " , PTO_NETRCA=" & XN(mtxtnetrca)
'        'If chkkitcon.Value = 1 Then
'            sql = sql & " , PTO_KITCON=" & chkkitcon.Value
'        'Else
'        '    sql = sql & " , PTO_KITCON=" & XN("0")
'        'End If
'
'        sql = sql & " , PTO_SALHID=" & XS(mtxtsalhid)
'        sql = sql & " , PTO_POSARA=" & XS(mtxtposara)
'        sql = sql & " , PTO_CERFAB=" & XS(mtxtcerfab)
'        sql = sql & " , PTO_OPCION1=" & XS(mtxtopcional1)
'        sql = sql & " , PTO_OPCION2=" & XS(mtxtopcional2)
'        sql = sql & " , PTO_CODCLA=" & XS(txtcodrepu)
        sql = sql & " , PTO_PRECIOC=" & XN(txtpCpra)
        
        
        sql = sql & " WHERE PTO_CODIGO=" & XS(txtcodigo)
        DBConn.Execute sql
        
         'UPDATE DETALLE STOCK
        sql = "UPDATE DETALLE_STOCK "
        sql = sql & " SET DST_STKFIS =" & XN(txtSFisico.Text)
        sql = sql & " WHERE PTO_CODIGO LIKE '" & txtcodigo.Text & "'"
        DBConn.Execute sql
        
    Else
        'TxtCodigo = "1"
        'sql = "SELECT MAX(PTO_CODIGO) as maximo FROM PRODUCTO WHERE PTO_CODIGO <> 99999999"
        'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        'If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = XN(rec.Fields!Maximo) + 1
        'rec.Close
                
        sql = "INSERT INTO PRODUCTO(PTO_CODIGO,PTO_DESCRI,LNA_CODIGO,"
        sql = sql & "RUB_CODIGO,TPRE_CODIGO,PTO_PRECIO,PTO_STKMIN,PTO_ESTADO,"
        'sql = sql & "PTO_TIPO,PTO_TIPMOD,PTO_TRACCI,PTO_CABINA,PTO_MOTMAR,PTO_MOTMOD,"
        'sql = sql & "PTO_ASPIRA,PTO_MOTNRO,PTO_CHASIS,PTO_SERIE,PTO_NEUMDE,PTO_NEDECA,"
        'sql = sql & "PTO_NEUMTR,PTO_NETRCA,PTO_KITCON,PTO_SALHID,PTO_POSARA,PTO_CERFAB,"
        'sql = sql & "PTO_OPCION1,PTO_OPCION2,
        sql = sql & " PTO_PRECIOC,LIS_CODIGO,PTO_CODCLA,PTO_PRECIVA)"
        sql = sql & " VALUES ("
        sql = sql & XS(txtcodigo, True) & ","
        sql = sql & XS(Trim(txtdescri), True) & ","
        sql = sql & cbolinea.ItemData(cbolinea.ListIndex) & ","
        sql = sql & cborubro.ItemData(cborubro.ListIndex) & ","
        sql = sql & cboPres.ItemData(cboPres.ListIndex) & ","
        sql = sql & XN(txtPrecio) & ","
        sql = sql & XN(txtStock) & ","
        If chkBaja.Value = True Then
            sql = sql & "2" & "," 'DADO DE BAJA
        Else
            sql = sql & "1" & "," 'NORMAL
        End If
'        sql = sql & XS(mtxttipo) & ","
'        sql = sql & XS(mtxttipmod) & ","
'        sql = sql & XS(mtxttracci) & ","
'        If chkcabina.Value = True Then
'            sql = sql & XN("1") & ","
'        Else
'            sql = sql & XN("0") & ","
'        End If
'        sql = sql & XS(mtxtmotmar) & ","
'        sql = sql & XS(mtxtmotmod) & ","
'        sql = sql & XS(mtxtaspira) & ","
'        sql = sql & XS(mtxtmotnro) & ","
'        sql = sql & XS(mtxtchasis) & ","
'        sql = sql & XS(mtxtserie) & ","
'        sql = sql & XS(mtxtneumde) & ","
'        sql = sql & XN(mtxtnedeca) & ","
'        sql = sql & XS(mtxtneumtr) & ","
'        sql = sql & XN(mtxtnetrca) & ","
'        If chkkitcon.Value = True Then
'            sql = sql & XN("1") & ","
'        Else
'            sql = sql & XN("0") & ","
'        End If
'
'        sql = sql & XS(mtxtsalhid) & ","
'        sql = sql & XS(mtxtposara) & ","
'        sql = sql & XS(mtxtcerfab) & ","
'        sql = sql & XS(mtxtopcional1) & ","
'        sql = sql & XS(mtxtopcional2) & ","
        sql = sql & XN(txtpCpra) & ","
        If cboListaPrecio.ListIndex > 0 Then
            sql = sql & cboListaPrecio.ItemData(cboListaPrecio.ListIndex) & ","
        Else
            sql = sql & 0 & ","
        End If
        sql = sql & XS(txtcodrepu) & ","
        sql = sql & XN(PRECIVA) & ")"
        DBConn.Execute sql
                
        'Aca Inserto el producto en la tabla DETALLE_STOCK
        
        sql = "SELECT * FROM DETALLE_STOCK WHERE PTO_CODIGO LIKE '" & txtcodigo.Text & "' "
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = True Then
            'Do While Rec1.EOF <> True
                sql = "INSERT INTO DETALLE_STOCK(STK_CODIGO,PTO_CODIGO,DST_STKFIS)"
                sql = sql & " VALUES ("
                sql = sql & 1 & ","
                sql = sql & XS(txtcodigo) & ","
                sql = sql & XN(txtSFisico) & ") "
                DBConn.Execute sql
            '    Rec1.MoveNext
            'Loop
        End If
        Rec1.Close
        
        
        'INSERTO EN LISTA DE PRECIOS
'        If CODIGOLISTA <> 0 Then
''            sql = "INSERT INTO DETALLE_LISTA_PRECIO(LIS_CODIGO,"
''            sql = sql & "PTO_CODIGO,LIS_PRECIO,LIS_PRECIOC)"
''            sql = sql & " VALUES ("
''            sql = sql & CODIGOLISTA & ","
''            sql = sql & XS(TxtCodigo) & ","
''            sql = sql & XN(txtPrecio.Text) & " ,"
''            sql = sql & XN(txtpCpra.Text) & " )"
''            DBConn.Execute sql
'            sql = " UPDATE PRODUCTO "
'            sql = sql & " SET LIS_CODIGO = " & CODIGOLISTA & " "
'            sql = sql & " WHERE PTO_CODIGO LIKE '" & TxtCodigo.Text & "'"
'            DBConn.Execute sql
'        End If
        
    
    End If
    Screen.MousePointer = vbNormal
    DBConn.CommitTrans
    rec.Close
    
    If Consulta = 3 Then
        Me.Hide
    Else
        CmdNuevo_Click
    End If
    
    Exit Sub
    
    
    
HayError:
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX

End Sub

Private Sub cmdLisPrecio_Click()
    Consulta = 2
    FrmListadePrecios.Show vbModal
    Set FrmListadePrecios = Nothing
    cboListaPrecio.Clear
    CargoCboListaPrecio
End Sub

Private Sub cmdNuevaLinea_Click()
    ABMLinea.Show vbModal
    Set ABMLinea = Nothing
    cbolinea.Clear
    cargocboLinea
End Sub

Private Sub cmdNuevaPresentacion_Click()
    ABMPresentacion.Show vbModal
    Set ABMPresentacion = Nothing
    cboPres.Clear
    cargocbotiporep
End Sub
Private Sub CmdNuevo_Click()
    txtcodigo.Text = ""
    txtdescri.Text = ""
    lblEstado.Caption = ""
    txtPrecio.Text = "0,00"
    txtpCpra.Text = "0,00"
    txtStock.Text = ""
    GrdModulos.Rows = 1
    chkBaja.Value = Unchecked
    cbolinea.ListIndex = 0
    cborubro.Clear
    'cboPres.ListIndex = 0
    txtcodigo.SetFocus
    
    mtxttipo.Text = ""
    mtxttipmod.Text = ""
    mtxttracci.Text = ""
    chkcabina.Value = False
    mtxtmotmar.Text = ""
    mtxtmotmod.Text = ""
    mtxtaspira.Text = ""
    mtxtmotnro.Text = ""
    mtxtchasis.Text = ""
    mtxtserie.Text = ""
    mtxtneumde.Text = ""
    mtxtnedeca.Text = ""
    mtxtneumtr.Text = ""
    mtxtnetrca.Text = ""
    chkkitcon.Value = False
    mtxtsalhid.Text = ""
    mtxtposara.Text = ""
    mtxtcerfab.Text = ""
    mtxtopcional1.Text = ""
    mtxtopcional2.Text = ""
    txtpCpra.Text = ""
    txtcodrepu.Text = ""
    txtSFisico.Text = ""
    cboListaPrecio.ListIndex = 0
End Sub

Private Sub cmdNuevoRubro_Click()
    ABMRubro.Show vbModal
    Set ABMRubro = Nothing
    cborubro.Clear
    cargocboRubro
End Sub

Private Sub CmdSalir_Click()
    If Consulta = 3 Then
        Consulta = 4
    End If
    Unload Me
    Set ABMProducto = Nothing
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
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
Set Rec2 = New ADODB.Recordset
Dim CANTREC    As Integer
    
'CORRER ESTE PROCESO  !!!!!!!!!!!!!!! CORRIDO EL 02/09/2006
'CORRI ESTE PROCESO CUANDO FALLO LA ACTUALIZACION DE PRECIO CIVA LISTA MAIZCO
'sql = "SELECT * FROM PRODUCTO WHERE LIS_CODIGO = 27"
'Dim IVA As Double
'IVA = 1.21
'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'If rec.EOF = False Then
'    Do While rec.EOF = False
'        sql = "UPDATE PRODUCTO SET PTO_PRECIVA = " & XN(rec!PTO_PRECIO * IVA) & " WHERE PTO_CODIGO LIKE '" & rec!PTO_CODIGO & "' AND LIS_CODIGO = " & rec!LIS_CODIGO
'        DBConn.Execute sql
'        rec.MoveNext
'    Loop
'End If
'rec.Close



''****** REMITO DE CLIENTES ******
'
'sql = "SELECT P.PTO_DESCRI,DR.PTO_CODIGO,DR.DRC_DETALLE "
'sql = sql & "FROM PRODUCTO P, DETALLE_REMITO_CLIENTE DR "
'sql = sql & " WHERE P.PTO_CODIGO = DR.PTO_CODIGO "
'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'Do While rec.EOF = False
'    sql = "UPDATE DETALLE_REMITO_CLIENTE "
'    sql = sql & "SET DRC_DETALLE = " & XS(rec!PTO_DESCRI) & " "
'    sql = sql & "WHERE PTO_CODIGO LIKE '" & rec!PTO_CODIGO & "'"
'    DBConn.Execute sql
'    rec.MoveNext
'Loop
'rec.Close
''************************************
'
''*****REMITO DE PROVEEDORES****
'sql = "SELECT P.PTO_DESCRI,DR.PTO_CODIGO,DR.DRPR_DETALLE "
'sql = sql & "FROM PRODUCTO P, DETALLE_REMITO_PROVEEDOR DR "
'sql = sql & " WHERE P.PTO_CODIGO = DR.PTO_CODIGO "
'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'Do While rec.EOF = False
'    sql = "UPDATE DETALLE_REMITO_PROVEEDOR "
'    sql = sql & "SET DRPR_DETALLE = " & XS(rec!PTO_DESCRI) & " "
'    sql = sql & "WHERE PTO_CODIGO LIKE '" & rec!PTO_CODIGO & "'"
'    DBConn.Execute sql
'    rec.MoveNext
'Loop
'rec.Close
''*************************************
'
''********* FACTURA DE CLIENTE ********
'sql = "SELECT P.PTO_DESCRI,DR.PTO_CODIGO,DR.DFC_DETALLE "
'sql = sql & "FROM PRODUCTO P, DETALLE_FACTURA_CLIENTE DR "
'sql = sql & " WHERE P.PTO_CODIGO = DR.PTO_CODIGO "
'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'Do While rec.EOF = False
'    sql = "UPDATE DETALLE_FACTURA_CLIENTE "
'    sql = sql & "SET DFC_DETALLE = " & XS(rec!PTO_DESCRI) & " "
'    sql = sql & "WHERE PTO_CODIGO LIKE '" & rec!PTO_CODIGO & "'"
'    DBConn.Execute sql
'    rec.MoveNext
'Loop
'rec.Close
''***************************************
''********* FACTURA DE PROVEEDOR ********
'sql = "SELECT P.PTO_DESCRI,DR.PTO_CODIGO,DR.DFP_DETALLE "
'sql = sql & "FROM PRODUCTO P, DETALLE_FACTURA_PROVEEDOR DR "
'sql = sql & " WHERE P.PTO_CODIGO = DR.PTO_CODIGO "
'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'Do While rec.EOF = False
'    sql = "UPDATE DETALLE_FACTURA_PROVEEDOR "
'    sql = sql & "SET DFP_DETALLE = " & XS(rec!PTO_DESCRI) & " "
'    sql = sql & "WHERE PTO_CODIGO LIKE '" & rec!PTO_CODIGO & "'"
'    DBConn.Execute sql
'    rec.MoveNext
'Loop
'rec.Close
''**************************************
''********* NOTA DE CREDITO DE CLIENTE ********
'sql = "SELECT P.PTO_DESCRI,DR.PTO_CODIGO,DR.DNC_DETALLE "
'sql = sql & "FROM PRODUCTO P, DETALLE_NOTA_CREDITO_CLIENTE DR "
'sql = sql & " WHERE P.PTO_CODIGO = DR.PTO_CODIGO "
'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'Do While rec.EOF = False
'    sql = "UPDATE DETALLE_NOTA_CREDITO_CLIENTE "
'    sql = sql & "SET DNC_DETALLE = " & XS(rec!PTO_DESCRI) & " "
'    sql = sql & "WHERE PTO_CODIGO LIKE '" & rec!PTO_CODIGO & "'"
'    DBConn.Execute sql
'    rec.MoveNext
'Loop
'rec.Close
''*****************************************
''********* NOTA DE CREDITO DE PROVEEDOR ********
'sql = "SELECT P.PTO_DESCRI,DR.PTO_CODIGO,DR.DNCP_DETALLE "
'sql = sql & "FROM PRODUCTO P, DETALLE_NOTA_CREDITO_PROVEEDOR DR "
'sql = sql & " WHERE P.PTO_CODIGO = DR.PTO_CODIGO "
'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'Do While rec.EOF = False
'    sql = "UPDATE DETALLE_NOTA_CREDITO_PROVEEDOR "
'    sql = sql & "SET DNCP_DETALLE = " & XS(rec!PTO_DESCRI) & " "
'    sql = sql & "WHERE PTO_CODIGO LIKE '" & rec!PTO_CODIGO & "'"
'    DBConn.Execute sql
'    rec.MoveNext
'Loop
'rec.Close
''*****************************************








'    sql = "SELECT * FROM PRODUCTO"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    CANTREC = rec.RecordCount
'    If rec.EOF = False Then
'        Do While Not rec.EOF
'            sql = "INSERT INTO DETALLE_STOCK(STK_CODIGO,PTO_CODIGO)" ',DST_STKFIS)"
'            sql = sql & " VALUES ("
'            sql = sql & 1 & ","
'            sql = sql & XS(rec!PTO_CODIGO) & ")"
'            'sql = sql & XN(txtStock) & ") "
'            DBConn.Execute sql
'            rec.MoveNext
'        Loop
'    End If
'    rec.Close

'PROCESO QUE USE PARA VER SI HAY MAQ CARGADAS EN UNA LISTA DE REPUESTOS
'sql = "SELECT * FROM PRODUCTO WHERE LIS_CODIGO = 27"
'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'If rec.EOF = False Then
'    Do While rec.EOF = False
'        If rec!LNA_CODIGO = 6 Then
'            MsgBox rec!PTO_CODIGO
'        End If
'        rec.MoveNext
'    Loop
'End If
'rec.Close

    
    Call Centrar_pantalla(Me)
    
    lblEstado.Caption = ""
    GrdModulos.FormatString = "Código|Descripción|Rubro|Linea"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 4000
    GrdModulos.ColWidth(2) = 2500
    GrdModulos.ColWidth(3) = 2500
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    tabdatos2.Tab = 0
    'TxtCodigo.SetFocus
    
    cargocboLinea
    'cargocboLista de Precio
    CargoCboListaPrecio
    'cargo combo Tipo de Presentación
    
    txtPrecio.Text = "0,00"
    txtpCpra.Text = "0,00"
    
    
End Sub

Private Sub CargoCboListaPrecio()
    sql = "SELECT LIS_CODIGO, LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO"
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        cboListaPrecio.AddItem ""
        Do While rec.EOF = False
            cboListaPrecio.AddItem rec!LIS_DESCRI
            cboListaPrecio.ItemData(cboListaPrecio.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboListaPrecio.ListIndex = 0
    End If
    rec.Close
End Sub
Private Sub cargocboLinea()
    sql = "SELECT * FROM LINEAS  ORDER BY LNA_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cbolinea.AddItem rec!LNA_DESCRI
            cbolinea.ItemData(cbolinea.NewIndex) = rec!LNA_CODIGO
            rec.MoveNext
        Loop
        cbolinea.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub cargocboRubro()
    If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Then Exit Sub
    cborubro.Clear
    sql = "SELECT * FROM RUBROS "
    sql = sql & " WHERE LNA_CODIGO= " & cbolinea.ItemData(cbolinea.ListIndex)
    sql = sql & " ORDER BY RUB_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cborubro.AddItem rec!RUB_DESCRI
            cborubro.ItemData(cborubro.NewIndex) = rec!RUB_CODIGO
            rec.MoveNext
        Loop
        cborubro.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub cargocbotiporep()
    cboPres.Clear
    sql = "SELECT * FROM TIPO_PRESENTACION "
    sql = sql & " WHERE LNA_CODIGO = " & cbolinea.ItemData(cbolinea.ListIndex) & " "
    If cborubro.ListIndex <> -1 Then
        sql = sql & " AND RUB_CODIGO = " & cborubro.ItemData(cborubro.ListIndex) & " "
    End If
    sql = sql & " ORDER BY TPRE_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboPres.AddItem rec!TPRE_DESCRI
            cboPres.ItemData(cboPres.NewIndex) = rec!TPRE_CODIGO
            rec.MoveNext
        Loop
        cboPres.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
           txtcodigo = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
           cmdGrabar.Enabled = True
           CmdBorrar.Enabled = True
           TxtCodigo_LostFocus
           tabDatos.Tab = 0
           tabdatos2.Tab = 0
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub mtxttipo_GotFocus()
    SelecTexto mtxttipo
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 0 And Me.Visible Then
        txtdescri.SetFocus
        cmdGrabar.Enabled = True
        CmdBorrar.Enabled = True
    End If
    If tabDatos.Tab = 1 Then
        TxtDescriB.Text = ""
        TxtDescriB.SetFocus
        cmdGrabar.Enabled = False
        CmdBorrar.Enabled = False
    End If
End Sub

Private Sub Text3_Change()

End Sub
Private Sub TxtCodigo_Change()
    If Trim(txtcodigo) = "" And CmdBorrar.Enabled Then
        CmdBorrar.Enabled = False
    ElseIf Trim(txtcodigo) <> "" Then
        CmdBorrar.Enabled = True
    End If
    
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto txtcodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
'    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    txtcodigo.Text = Replace(txtcodigo.Text, "'", "´")
    
    If txtcodigo.Text <> "" Then
        sql = "SELECT * FROM PRODUCTO"
        sql = sql & " WHERE PTO_CODIGO LIKE '" & txtcodigo.Text & "'"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtdescri.Text = Rec1!PTO_DESCRI
            Call BuscaCodigoProxItemData(Rec1!LNA_CODIGO, cbolinea)
            cboLinea_LostFocus
            Call BuscaCodigoProxItemData(Rec1!RUB_CODIGO, cborubro)
            cboRubro_LostFocus
            Call BuscaCodigoProxItemData(Rec1!TPRE_CODIGO, cboPres)
            txtPrecio.Text = IIf(IsNull(Rec1!PTO_PRECIO), "", Format(Rec1!PTO_PRECIO, "0.00"))
            txtStock.Text = IIf(IsNull(Rec1!PTO_STKMIN), "", Rec1!PTO_STKMIN)
            If Rec1!PTO_ESTADO = 2 Then chkBaja.Value = Checked
            txtdescri.SetFocus
            Call BuscaCodigoProxItemData(Rec1!LIS_CODIGO, cboListaPrecio)
            
'            mtxttipo.Text = IIf(IsNull(Rec1!PTO_TIPO), "", Rec1!PTO_TIPO)
'            mtxttipmod.Text = IIf(IsNull(Rec1!PTO_TIPMOD), "", Rec1!PTO_TIPMOD)
'            mtxttracci.Text = IIf(IsNull(Rec1!PTO_TRACCI), "", Rec1!PTO_TRACCI)
'            chkcabina.Value = IIf(IsNull(Rec1!PTO_CABINA), 0, Rec1!PTO_CABINA)
'            mtxtmotmar.Text = IIf(IsNull(Rec1!PTO_MOTMAR), "", Rec1!PTO_MOTMAR)
'            mtxtmotmod.Text = IIf(IsNull(Rec1!PTO_MOTMOD), "", Rec1!PTO_MOTMOD)
'            mtxtaspira.Text = IIf(IsNull(Rec1!PTO_ASPIRA), "", Rec1!PTO_ASPIRA)
'            mtxtmotnro.Text = IIf(IsNull(Rec1!PTO_MOTNRO), "", Rec1!PTO_MOTNRO)
'            mtxtchasis.Text = IIf(IsNull(Rec1!PTO_CHASIS), "", Rec1!PTO_CHASIS)
'            mtxtserie.Text = IIf(IsNull(Rec1!PTO_SERIE), "", Rec1!PTO_SERIE)
'            mtxtneumde.Text = IIf(IsNull(Rec1!PTO_NEUMDE), "", Rec1!PTO_NEUMDE)
'            mtxtnedeca.Text = IIf(IsNull(Rec1!PTO_NEDECA), "", Rec1!PTO_NEDECA)
'            mtxtneumtr.Text = IIf(IsNull(Rec1!PTO_NEUMTR), "", Rec1!PTO_NEUMTR)
'            mtxtnetrca.Text = IIf(IsNull(Rec1!PTO_NETRCA), "", Rec1!PTO_NETRCA)
'            chkkitcon.Value = IIf(IsNull(Rec1!PTO_KITCON), 0, Rec1!PTO_KITCON)
'            mtxtsalhid.Text = IIf(IsNull(Rec1!PTO_SALHID), "", Rec1!PTO_SALHID)
'            mtxtposara.Text = IIf(IsNull(Rec1!PTO_POSARA), "", Rec1!PTO_POSARA)
'            mtxtcerfab.Text = IIf(IsNull(Rec1!PTO_CERFAB), "", Rec1!PTO_CERFAB)
'            mtxtopcional1.Text = IIf(IsNull(Rec1!PTO_OPCION1), "", Rec1!PTO_OPCION1)
'            mtxtopcional2.Text = IIf(IsNull(Rec1!PTO_OPCION2), "", Rec1!PTO_OPCION2)
'            txtcodrepu.Text = IIf(IsNull(Rec1!PTO_CODCLA), "", Rec1!PTO_CODCLA)
            txtpCpra.Text = IIf(IsNull(Rec1!PTO_PRECIOC), "", Format(Rec1!PTO_PRECIOC, "0.00"))
            
            'Busco Stock en Detalle Stock
            sql = "SELECT DST_STKFIS FROM DETALLE_STOCK WHERE PTO_CODIGO LIKE '" & txtcodigo.Text & "'"
            Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec2.EOF = False Then
                txtSFisico.Text = IIf(IsNull(Rec2!DST_STKFIS), "", Rec2!DST_STKFIS)
            End If
            Rec2.Close
        Else
         'MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
         'TxtCodigo.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtcodrepu_Change()
'    SelecTexto txtcodrepu
End Sub

Private Sub txtcodrepu_KeyPress(KeyAscii As Integer)
 '   KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_Change()
    If Trim(txtdescri) = "" And cmdGrabar.Enabled Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If
End Sub

Private Sub txtdescri_GotFocus()
    SelecTexto txtdescri
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_LostFocus()
    txtdescri.Text = Replace(txtdescri, "'", "´")
End Sub

Private Sub txtpCpra_GotFocus()
    SelecTexto txtpCpra
End Sub

Private Sub txtpCpra_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtpCpra, KeyAscii)
End Sub

Private Sub txtpCpra_LostFocus()
    If txtpCpra.Text <> "" Then
        txtpCpra.Text = Valido_Importe(txtpCpra)
    Else
        txtpCpra.Text = "0,00"
    End If
End Sub

Private Sub txtPrecio_GotFocus()
    SelecTexto txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPrecio, KeyAscii)
End Sub
Function validarProuducto()
    If txtcodigo.Text = "" Then
        MsgBox "No ha ingresado el código del Producto", vbExclamation, TIT_MSGBOX
        txtcodigo.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If txtdescri.Text = "" Then
        MsgBox "No ha ingresado la Descripción", vbExclamation, TIT_MSGBOX
        txtdescri.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If cbolinea.ListIndex = -1 Then
        MsgBox "No ha seleccionado la Línea del Producto", vbExclamation, TIT_MSGBOX
        cbolinea.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If cborubro.Text = "" Then
        MsgBox "No ha seleccionado el Rubro del Producto", vbExclamation, TIT_MSGBOX
        cborubro.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If cboPres.ListIndex = -1 Then
        MsgBox "No ha seleccionado la Marca del Producto", vbExclamation, TIT_MSGBOX
        cboPres.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If txtPrecio.Text = "" Then
        MsgBox "No ha ingresado el Precio de Venta", vbExclamation, TIT_MSGBOX
        txtPrecio.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If txtpCpra.Text = "" Then
        MsgBox "No ha ingresado el Precio de Compra", vbExclamation, TIT_MSGBOX
        txtpCpra.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If txtStock.Text = "" Then
        MsgBox "No ha ingresado el Stock Mínimo", vbExclamation, TIT_MSGBOX
        txtStock.SetFocus
        validarProuducto = False
        Exit Function
    End If
'    If cboLinea.ItemData(cboLinea.ListIndex) = 7 Then 'si es repuesto
'        If txtcodrepu.Text = "" Then
'        MsgBox "No ha ingresado el Código de Clasificación del Repuesto", vbExclamation, TIT_MSGBOX
'        txtcodrepu.SetFocus
'        validarProuducto = False
'        Exit Function
'    End If
'    End If
    validarProuducto = True
End Function

Private Sub txtPrecio_LostFocus()
    If txtPrecio.Text <> "" Then
        txtPrecio.Text = Valido_Importe(txtPrecio)
    Else
        txtPrecio.Text = "0,00"
    End If
End Sub

Private Sub txtSFisico_GotFocus()
    SelecTexto txtSFisico
End Sub


Private Sub txtSFisico_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
Private Sub txtStock_GotFocus()
    SelecTexto txtStock
End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
