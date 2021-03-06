VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotaCreditoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Cr�dito Clientes..."
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   570
      Left            =   8175
      TabIndex        =   14
      Top             =   7695
      Width           =   990
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   570
      Left            =   10185
      TabIndex        =   16
      Top             =   7695
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   570
      Left            =   7170
      TabIndex        =   13
      Top             =   7695
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   570
      Left            =   9180
      TabIndex        =   15
      Top             =   7695
      Width           =   990
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7650
      Left            =   60
      TabIndex        =   26
      Top             =   15
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   13494
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   512
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
      TabPicture(0)   =   "frmNotaCreditoCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameNotaCredito"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameCliente"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tabLista"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmNotaCreditoCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "frameBuscar"
      Tab(1).ControlCount=   2
      Begin TabDlg.SSTab tabLista 
         Height          =   1215
         Left            =   120
         TabIndex        =   77
         Top             =   1905
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2143
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Accesorios"
         TabPicture(0)   =   "frmNotaCreditoCliente.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Repuestos"
         TabPicture(1)   =   "frmNotaCreditoCliente.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame5"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame5 
            Caption         =   "Lista de Precios"
            ForeColor       =   &H8000000D&
            Height          =   735
            Left            =   -74880
            TabIndex        =   82
            Top             =   360
            Width           =   3495
            Begin VB.ComboBox cboLPrecioRep 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   240
               Width           =   3225
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Lista de Precios"
            ForeColor       =   &H8000000D&
            Height          =   735
            Left            =   -74880
            TabIndex        =   80
            Top             =   360
            Width           =   3495
            Begin VB.ComboBox cboLPrecioRep1 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   81
               Top             =   240
               Width           =   3225
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Lista de Precios"
            ForeColor       =   &H8000000D&
            Height          =   735
            Left            =   120
            TabIndex        =   78
            Top             =   360
            Width           =   3495
            Begin VB.ComboBox cboListaPrecio 
               Height          =   315
               Left            =   600
               Style           =   2  'Dropdown List
               TabIndex        =   79
               Top             =   240
               Width           =   2505
            End
         End
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   4200
         TabIndex        =   58
         Top             =   330
         Width           =   6810
         Begin VB.TextBox txtIngBrutos 
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
            Left            =   5415
            TabIndex        =   69
            Top             =   1770
            Width           =   1215
         End
         Begin VB.TextBox txtCondicionIVA 
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
            Left            =   2490
            TabIndex        =   68
            Top             =   1770
            Width           =   2895
         End
         Begin VB.TextBox txtCUIT 
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
            Left            =   990
            TabIndex        =   67
            Top             =   1770
            Width           =   1455
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   330
            Left            =   1875
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaCreditoCliente.frx":0070
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Buscar Cliente"
            Top             =   260
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtDomici 
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
            Left            =   990
            MaxLength       =   50
            TabIndex        =   64
            Top             =   656
            Width           =   4365
         End
         Begin VB.TextBox txtCliLocalidad 
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
            Left            =   990
            TabIndex        =   62
            Top             =   1027
            Width           =   4365
         End
         Begin VB.TextBox txtProvincia 
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
            Left            =   990
            TabIndex        =   60
            Top             =   1398
            Width           =   4365
         End
         Begin VB.TextBox txtCliRazSoc 
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
            Left            =   2310
            MaxLength       =   50
            TabIndex        =   5
            Tag             =   "Descripci�n"
            Top             =   285
            Width           =   4320
         End
         Begin VB.TextBox txtCodCliente 
            Height          =   285
            Left            =   990
            MaxLength       =   40
            TabIndex        =   4
            Top             =   285
            Width           =   840
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos"
            Height          =   195
            Left            =   5595
            TabIndex        =   71
            Top             =   1560
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   285
            TabIndex        =   70
            Top             =   1815
            Width           =   600
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   210
            TabIndex        =   65
            Top             =   712
            Width           =   675
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   150
            TabIndex        =   63
            Top             =   1079
            Width           =   735
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   180
            TabIndex        =   61
            Top             =   1446
            Width           =   705
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "C�digo:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   345
            TabIndex        =   59
            Top             =   345
            Width           =   540
         End
      End
      Begin VB.Frame frameBuscar 
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
         Height          =   1680
         Left            =   -74610
         TabIndex        =   31
         Top             =   645
         Width           =   10410
         Begin VB.ComboBox cboNotaCredito1 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1095
            Width           =   2400
         End
         Begin VB.CheckBox chkTipoFactura 
            Caption         =   "Tipo"
            Height          =   195
            Left            =   540
            TabIndex        =   19
            Top             =   1125
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4050
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaCreditoCliente.frx":037A
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Buscar Cliente"
            Top             =   375
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1140
            Left            =   9690
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaCreditoCliente.frx":0684
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Buscar "
            Top             =   345
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
            Left            =   4485
            MaxLength       =   50
            TabIndex        =   32
            Tag             =   "Descripci�n"
            Top             =   375
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3000
            MaxLength       =   40
            TabIndex        =   20
            Top             =   375
            Width           =   975
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   540
            TabIndex        =   18
            Top             =   870
            Width           =   810
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   540
            TabIndex        =   17
            Top             =   600
            Width           =   855
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   3000
            TabIndex        =   21
            Top             =   720
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
            Left            =   5610
            TabIndex        =   22
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   52559873
            CurrentDate     =   41098
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   2535
            TabIndex        =   56
            Top             =   1020
            Width           =   360
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4575
            TabIndex        =   35
            Top             =   795
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1890
            TabIndex        =   34
            Top             =   780
            Width           =   1005
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   2370
            TabIndex        =   33
            Top             =   420
            Width           =   525
         End
      End
      Begin VB.Frame FrameNotaCredito 
         Caption         =   "Nota de Cr�dito..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   105
         TabIndex        =   28
         Top             =   330
         Width           =   4095
         Begin VB.TextBox txtNroNotaCredito 
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
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   2
            Top             =   600
            Width           =   1065
         End
         Begin VB.ComboBox cboNotaCredito 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   270
            Width           =   2400
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
            Height          =   315
            Left            =   1170
            MaxLength       =   4
            TabIndex        =   1
            Top             =   600
            Width           =   555
         End
         Begin MSComCtl2.DTPicker FechaNotaCredito 
            Height          =   315
            Left            =   1170
            TabIndex        =   3
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   52559873
            CurrentDate     =   41098
         End
         Begin VB.Label lblEstadoNotaCredito 
            AutoSize        =   -1  'True
            Caption         =   "EST. NOTA CREDITO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1170
            TabIndex        =   72
            Top             =   1320
            Width           =   1890
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   750
            TabIndex        =   43
            Top             =   285
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   615
            TabIndex        =   40
            Top             =   885
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "N�mero:"
            Height          =   195
            Left            =   510
            TabIndex        =   39
            Top             =   585
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   570
            TabIndex        =   38
            Top             =   1305
            Width           =   540
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4500
         Left            =   -74625
         TabIndex        =   25
         Top             =   2610
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7938
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame4 
         Height          =   540
         Left            =   4200
         TabIndex        =   41
         Top             =   2500
         Width           =   6815
         Begin VB.ComboBox cboConcepto 
            Height          =   315
            Left            =   970
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   120
            Width           =   5715
         End
         Begin VB.Label lblConcepto 
            AutoSize        =   -1  'True
            Caption         =   "Concepto:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   165
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4590
         Left            =   105
         TabIndex        =   29
         Top             =   3025
         Width           =   10935
         Begin VB.CommandButton cmdQuitarProducto 
            Height          =   330
            Left            =   10400
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaCreditoCliente.frx":2E26
            Style           =   1  'Graphical
            TabIndex        =   76
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Producto"
            Top             =   1050
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdAgregarProducto 
            Height          =   330
            Left            =   10400
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaCreditoCliente.frx":3BA8
            Style           =   1  'Graphical
            TabIndex        =   75
            TabStop         =   0   'False
            ToolTipText     =   "Agregar Producto"
            Top             =   705
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdBuscarProducto 
            Height          =   330
            Left            =   10400
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaCreditoCliente.frx":3EB2
            Style           =   1  'Graphical
            TabIndex        =   74
            TabStop         =   0   'False
            ToolTipText     =   "Buscar Producto"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CheckBox chkBonificaEnPesos 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en $"
            Height          =   285
            Left            =   390
            TabIndex        =   9
            Top             =   3840
            Width           =   1290
         End
         Begin VB.CheckBox chkBonificaEnPorsentaje 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en % "
            Height          =   285
            Left            =   390
            TabIndex        =   8
            Top             =   3540
            Width           =   1290
         End
         Begin VB.TextBox txtSubTotalBoni 
            Alignment       =   1  'Right Justify
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
            Left            =   4905
            TabIndex        =   54
            Top             =   3870
            Width           =   1155
         End
         Begin VB.TextBox txtImporteIva 
            Alignment       =   1  'Right Justify
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
            Left            =   6900
            TabIndex        =   51
            Top             =   3870
            Width           =   1155
         End
         Begin VB.TextBox txtImporteBoni 
            Alignment       =   1  'Right Justify
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
            Left            =   2850
            TabIndex        =   48
            Top             =   3870
            Width           =   1155
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
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
            Left            =   8970
            TabIndex        =   45
            Top             =   3870
            Width           =   1350
         End
         Begin VB.TextBox txtSubtotal 
            Alignment       =   1  'Right Justify
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
            Left            =   8970
            TabIndex        =   44
            Top             =   3540
            Width           =   1350
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   1455
            MaxLength       =   60
            TabIndex        =   12
            Top             =   4215
            Width           =   8865
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   480
            TabIndex        =   30
            Top             =   480
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3375
            Left            =   75
            TabIndex        =   7
            Top             =   120
            Width           =   10245
            _ExtentX        =   18071
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   3
            Cols            =   11
            FixedCols       =   0
            BackColorSel    =   12648447
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   1
            AllowUserResizing=   3
         End
         Begin VB.TextBox txtPorcentajeIva 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6900
            TabIndex        =   11
            Top             =   3540
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeBoni 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2850
            TabIndex        =   10
            Top             =   3540
            Width           =   1155
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   4110
            TabIndex        =   55
            Top             =   3930
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   6270
            TabIndex        =   53
            Top             =   3915
            Width           =   570
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "% I.V.A.:"
            Height          =   195
            Left            =   6240
            TabIndex        =   52
            Top             =   3570
            Width           =   600
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   2235
            TabIndex        =   50
            Top             =   3915
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Bonificaci�n:"
            Height          =   195
            Left            =   1890
            TabIndex        =   49
            Top             =   3570
            Width           =   915
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   8505
            TabIndex        =   47
            Top             =   3915
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   8175
            TabIndex        =   46
            Top             =   3570
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   210
            TabIndex        =   42
            Top             =   4260
            Width           =   1110
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   27
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Nota de Cr�dito"
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
      Left            =   3960
      TabIndex        =   73
      Top             =   7920
      Width           =   2985
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
      Left            =   225
      TabIndex        =   37
      Top             =   7755
      Width           =   750
   End
End
Attribute VB_Name = "frmNotaCreditoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim w As Integer
Dim TipoBusquedaDoc As Integer
Dim VBonificacion As Double
Dim VTotal As Double
Dim VEstadoNotaCredito As Integer

Private Sub cboNotaCredito_LostFocus()
    txtPorcentajeIva_LostFocus
End Sub

Private Sub chkBonificaEnPesos_Click()
    If chkBonificaEnPesos.Value = Checked Then
        chkBonificaEnPorsentaje.Value = Unchecked
        chkBonificaEnPorsentaje.Enabled = False
    Else
        chkBonificaEnPorsentaje.Enabled = True
    End If
    txtPorcentajeBoni.Text = ""
    txtImporteBoni.Text = ""
    txtSubTotalBoni.Text = ""
End Sub

Private Sub chkBonificaEnPorsentaje_Click()
    If chkBonificaEnPorsentaje.Value = Checked Then
        chkBonificaEnPesos.Value = Unchecked
        chkBonificaEnPesos.Enabled = False
    Else
        chkBonificaEnPesos.Enabled = True
    End If
    txtPorcentajeBoni.Text = ""
    txtImporteBoni.Text = ""
    txtSubTotalBoni.Text = ""
End Sub

Private Sub cmdAgregarProducto_Click()
    Consulta = 3
    ABMProducto.CODIGOLISTA = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
    ABMProducto.Show vbModal
    If Consulta <> 4 Then
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        If Trim(ABMProducto.txtcodigo) <> "" Then txtEdit.Text = ABMProducto.txtcodigo
        TxtEdit_KeyDown vbKeyReturn, 0
    End If
End Sub

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCodCliente.Text = frmBuscar.grdBuscar.Text
        txtCodCliente_LostFocus
        txtCliRazSoc.SetFocus
    Else
        txtCodCliente.SetFocus
    End If
End Sub

Private Sub cmdBuscarProducto_Click()
    Consulta = 3
    If tabLista.Tab = 0 Then
        FrmListadePrecios.tabLista.Tab = 0
        FrmListadePrecios.cboListaPrecio.ListIndex = cboListaPrecio.ListIndex
    Else
        FrmListadePrecios.tabLista.Tab = 1
        FrmListadePrecios.cboLPrecioRep.ListIndex = cboLPrecioRep.ListIndex
    End If
    FrmListadePrecios.Show vbModal
    If Consulta <> 4 Then
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        If Trim(FrmListadePrecios.GrdModulos.Text) <> "" Then txtEdit.Text = FrmListadePrecios.GrdModulos.Text
        TxtEdit_KeyDown vbKeyReturn, 0
    End If
    Consulta = 2
End Sub

Private Sub cmdQuitarProducto_Click()
    If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
        If MsgBox("Seguro que desea quitar el Detalle: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 1), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = ""
            txtSubtotal.Text = ""
            txtImporteIva.Text = ""
            txtTotal.Text = ""
        End If
     End If
End Sub

Private Sub grdGrilla_LostFocus()
    If grdGrilla.TextMatrix(1, 8) = "MAQUINARIA" Then 'pregunta si la linea es Maquinaria
        txtPorcentajeIva.Text = "10,50"
    Else
        txtPorcentajeIva.Text = "21,00"
    End If
    txtPorcentajeIva_LostFocus
End Sub

Private Sub txtCliRazSoc_GotFocus()
    SelecTexto txtCliRazSoc
End Sub

Private Sub txtCodCliente_Change()
    If txtCodCliente.Text = "" Then
        txtCliRazSoc.Text = ""
        txtProvincia.Text = ""
        txtCliLocalidad.Text = ""
        txtDomici.Text = ""
        txtCUIT.Text = ""
        txtIngBrutos.Text = ""
        txtCondicionIVA.Text = ""
    End If
End Sub

Private Sub txtCodCliente_GotFocus()
    SelecTexto txtCodCliente
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodCliente_LostFocus()
    If txtCodCliente.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        Rec1.Open BuscoCliente(txtCodCliente), DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtCliRazSoc.Text = Rec1!CLI_RAZSOC
            txtProvincia.Text = Rec1!PRO_DESCRI
            txtCliLocalidad.Text = Rec1!LOC_DESCRI
            txtDomici.Text = Rec1!CLI_DOMICI
            txtCUIT.Text = IIf(IsNull(Rec1!CLI_CUIT), "", Rec1!CLI_CUIT)
            txtCondicionIVA.Text = Rec1!IVA_DESCRI
            txtIngBrutos.Text = IIf(IsNull(Rec1!CLI_INGBRU), "NO INFORMADO", Rec1!CLI_INGBRU)
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtCliRazSoc_Change()
    If txtCliRazSoc.Text = "" Then
        txtCodCliente.Text = ""
        txtProvincia.Text = ""
        txtCliLocalidad.Text = ""
        txtDomici.Text = ""
        txtCUIT.Text = ""
        txtIngBrutos.Text = ""
        txtCondicionIVA.Text = ""
    End If
End Sub

Private Sub txtCliRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtCliRazSoc_LostFocus()
    If txtCodCliente.Text = "" And txtCliRazSoc.Text <> "" Then
        rec.Open BuscoCliente(txtCliRazSoc), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 1
                frmBuscar.TxtDescriB.Text = txtCliRazSoc.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 0
                    txtCodCliente.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 1
                    txtCliRazSoc.Text = frmBuscar.grdBuscar.Text
                    txtCodCliente_LostFocus
                    FechaDesde.SetFocus
                Else
                    txtCodCliente.SetFocus
                End If
            Else
                txtCodCliente.Text = rec!CLI_CODIGO
                txtCliRazSoc.Text = rec!CLI_RAZSOC
                txtCodCliente_LostFocus
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
        rec.Close
    ElseIf txtCodCliente.Text = "" And txtCliRazSoc.Text = "" Then
        MsgBox "Debe elegir un cliente", vbExclamation, TIT_MSGBOX
        txtCodCliente.SetFocus
    End If
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

Private Sub chkTipoFactura_Click()
    If chkTipoFactura.Value = Checked Then
        cboNotaCredito1.Enabled = True
        cboNotaCredito1.ListIndex = 0
    Else
        cboNotaCredito1.Enabled = False
        cboNotaCredito1.ListIndex = -1
    End If
End Sub

Private Sub chkTipoFactura_LostFocus()
    If chkTipoFactura.Value = Checked And chkCliente.Value = Unchecked _
        And chkFecha.Value = Unchecked Then cboNotaCredito1.SetFocus
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
     sql = "SELECT NC.*, C.CLI_RAZSOC, TC.TCO_ABREVIA, C.CLI_DOMICI"
     sql = sql & " FROM NOTA_CREDITO_CLIENTE NC,"
     sql = sql & " TIPO_COMPROBANTE TC, CLIENTE C"
     sql = sql & " WHERE"
     sql = sql & " NC.TCO_CODIGO=TC.TCO_CODIGO"
     sql = sql & " AND NC.CLI_CODIGO=C.CLI_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND NC.CLI_CODIGO=" & XN(txtCliente)
    If Not IsNull(FechaDesde) Then sql = sql & " AND NC.NCC_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND NC.NCC_FECHA<=" & XDQ(FechaHasta)
    If chkTipoFactura.Value = Checked Then sql = sql & " AND NC.TCO_CODIGO=" & XN(cboNotaCredito1.ItemData(cboNotaCredito1.ListIndex))
    sql = sql & " ORDER BY NC.NCC_FECHA,NC.NCC_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!NCC_SUCURSAL, "0000") & "-" & Format(rec!NCC_NUMERO, "00000000") _
                            & Chr(9) & rec!NCC_FECHA & Chr(9) & Format(rec!NCC_TOTAL, "#0.00") & Chr(9) & rec!CLI_RAZSOC _
                            & Chr(9) & rec!CLI_DOMICI & Chr(9) & rec!EST_CODIGO _
                            & Chr(9) & rec!NCC_BONIFICA & Chr(9) & rec!NCC_IVA _
                            & Chr(9) & rec!NCC_OBSERVACION & Chr(9) & rec!TCO_CODIGO _
                            & Chr(9) & rec!CNC_CODIGO & Chr(9) & rec!NCC_BONIPESOS _
                            & Chr(9) & rec!CLI_CODIGO & Chr(9) & ""
                            
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

Private Sub cmdBuscarCli_Click()
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
End Sub

Private Sub cmdGrabar_Click()
    Dim VStockFisico As String
    
    If ValidarNotaCredito = False Then Exit Sub
    If MsgBox("�Confirma Nota de Cr�dito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayErrorFactura
    
    DBConn.BeginTrans
    sql = "SELECT * FROM NOTA_CREDITO_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
    
    sql = sql & " AND NCC_NUMERO = " & XN(txtNroNotaCredito)
    sql = sql & " AND NCC_SUCURSAL=" & XN(txtNroSucursal)
    'sql = sql & " AND REP_CODIGO=" & XN(cboRep.ItemData(cboRep.ListIndex))
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = True Then
        sql = "INSERT INTO NOTA_CREDITO_CLIENTE"
        sql = sql & " (TCO_CODIGO, NCC_NUMERO, NCC_SUCURSAL, NCC_FECHA,"
        sql = sql & " CLI_CODIGO, NCC_BONIFICA, NCC_IVA, CNC_CODIGO,"
        sql = sql & " NCC_OBSERVACION,NCC_NUMEROTXT,NCC_SUBTOTAL,NCC_TOTAL,NCC_SALDO,NCC_BONIPESOS, EST_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & cboNotaCredito.ItemData(cboNotaCredito.ListIndex) & ","
        sql = sql & XN(txtNroNotaCredito) & ","
        sql = sql & XN(txtNroSucursal) & ","
        sql = sql & XDQ(FechaNotaCredito) & ","
        sql = sql & XN(txtCodCliente) & ","
        sql = sql & XN(txtPorcentajeBoni) & ","
        sql = sql & XN(txtPorcentajeIva) & ","
        sql = sql & cboConcepto.ItemData(cboConcepto.ListIndex) & ","
        sql = sql & XS(txtObservaciones) & ","
        sql = sql & XS(Format(txtNroNotaCredito.Text, "00000000")) & ","
        If txtSubTotalBoni.Text <> "" Then 'SUBTOTAL
            sql = sql & XN(txtSubTotalBoni) & ","
        Else
            sql = sql & XN(txtSubtotal) & ","
        End If
        sql = sql & XN(txtTotal) & "," 'TOTAL
        sql = sql & XN(txtTotal) & "," 'SALDO DE LA NOTA DE CREDITO
        If chkBonificaEnPesos.Value = Checked Then
            sql = sql & "'S'" & "," 'BONIFICA EN PESOS
        ElseIf chkBonificaEnPorsentaje.Value = Checked Then
            sql = sql & "'N'" & "," 'BONIFICA EN PORCENTAJE
        Else
            sql = sql & "NULL" & "," 'NO HAY BONIFICACION
        End If
        sql = sql & "3)" 'ESTADO DEFINITIVO
        DBConn.Execute sql
           
        'DETALLE NOTA CREDITO
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                sql = "INSERT INTO DETALLE_NOTA_CREDITO_CLIENTE"
                sql = sql & " (TCO_CODIGO,NCC_NUMERO,NCC_SUCURSAL,"
                sql = sql & "NCC_FECHA,DNC_NROITEM,PTO_CODIGO"
                sql = sql & ",DNC_CANTIDAD,DNC_PRECIO,DNC_BONIFICA,DNC_DETALLE)"
                sql = sql & " VALUES ("
                sql = sql & cboNotaCredito.ItemData(cboNotaCredito.ListIndex) & ","
                sql = sql & XN(txtNroNotaCredito) & ","
                sql = sql & XN(txtNroSucursal) & ","
                sql = sql & XDQ(FechaNotaCredito) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 9)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(I, 0)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 4)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(I, 1)) & ")"
                DBConn.Execute sql
            End If
        Next
        
        ' actualizo el stock cuando es una devolucion
        
        If cboConcepto.ItemData(cboConcepto.ListIndex) = 1 Then 'devolucion
            For I = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(I, 0) <> "" Then
                    sql = "UPDATE DETALLE_STOCK"
                    sql = sql & " SET"
                    sql = sql & " DST_STKFIS = DST_STKFIS + " & XN(grdGrilla.TextMatrix(I, 2))
                    sql = sql & " WHERE STK_CODIGO= 1"
                    sql = sql & " AND PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "'"
                    DBConn.Execute sql
                End If
            Next
        End If
        
        
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO A LA NOTA DE CREDITO QUE CORRESPONDA
        sql = "SELECT * FROM PARAMETROS"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
                Select Case cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
                    Case 4
                        sql = "UPDATE PARAMETROS SET NOTA_CREDITO_A=" & XN(txtNroNotaCredito)
                    Case 5
                        sql = "UPDATE PARAMETROS SET NOTA_CREDITO_B=" & XN(txtNroNotaCredito)
                End Select
                    DBConn.Execute sql
        End If
        Rec1.Close
        
        'ACTUALIZO LA CUENTA CORRIENTE DEL CLIENTE
        DBConn.Execute AgregoCtaCteCliente(txtCodCliente, CStr(cboNotaCredito.ItemData(cboNotaCredito.ListIndex)) _
                                            , txtNroNotaCredito, txtNroSucursal, _
                                            FechaNotaCredito, txtTotal, "H", CStr(Date))
        
        DBConn.CommitTrans
    Else
        MsgBox "La Nota de Cr�dito ya fue Registrada", vbCritical, TIT_MSGBOX
        DBConn.CommitTrans
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    cmdImprimir_Click
    
    CmdNuevo_Click
    Exit Sub
    
HayErrorFactura:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarNotaCredito() As Boolean
    If IsNull(FechaNotaCredito.Value) Then
        MsgBox "La Fecha de la Nota de Cr�dito es requerida", vbExclamation, TIT_MSGBOX
        FechaNotaCredito.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If txtNroSucursal.Text = "" Then
        MsgBox "Debe ingresar el N�mero de Sucursal", vbExclamation, TIT_MSGBOX
        txtNroSucursal.SetFocus
        ValidarNotaCredito = False
    End If
    If txtNroNotaCredito.Text = "" Then
        MsgBox "Debe ingresar el N�mero de Nota de Cr�dito", vbExclamation, TIT_MSGBOX
        txtNroNotaCredito.SetFocus
        ValidarNotaCredito = False
    End If
    If txtCodCliente.Text = "" Then
        MsgBox "Debe ingresar un Cliente", vbExclamation, TIT_MSGBOX
        txtCodCliente.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If cboConcepto.ListIndex = -1 Then
        MsgBox "Debe ingresar el concepto por el cual se emite la Nota de Cr�dito", vbExclamation, TIT_MSGBOX
        cboConcepto.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If chkBonificaEnPesos.Value = Checked Or chkBonificaEnPorsentaje.Value = Checked Then
        If txtPorcentajeBoni.Text = "" Then
            MsgBox "Debe ingresar la Bonificaci�n", vbExclamation, TIT_MSGBOX
            txtPorcentajeBoni.SetFocus
            ValidarNotaCredito = False
            Exit Function
        End If
    End If
    ValidarNotaCredito = True
End Function

Private Sub cmdImprimir_Click()
    If MsgBox("�Confirma Impresi�n Nota de Cr�dito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'PONE A LA IMPRESORA  COMO PREDETERMINADA
    Dim X As Printer
    Dim mDriver As String
    mDriver = IMPRESORA
    For Each X In Printers
        If X.DeviceName = mDriver Then
            ' La define como predeterminada del sistema.
            Set Printer = X
            Exit For
        End If
    Next
'-----------------------------------
    Set_Impresora
    ImprimirNotaCredito
End Sub

Public Sub ImprimirNotaCredito()
    Dim Renglon As Double
    Dim canttxt As Integer
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Imprimiendo..."
    
    For w = 1 To 2 'SE IMPRIME POR DUPLICADO
      '-----IMPRESION DEL ENCABEZADO------------------
        ImprimirEncabezado
        
      '---- IMPRESION DE LA NOTA CREDITO ------------------
        Renglon = 9.4
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                Imprimir 0.9, Renglon, False, Format(grdGrilla.TextMatrix(I, 0), "000000")  'codigo
                If Len(grdGrilla.TextMatrix(I, 1)) <= 36 Then
                    Imprimir 3.8, Renglon, False, grdGrilla.TextMatrix(I, 1) 'descripcion
                Else
                    'CortarCadena 3.8, Renglon, grdGrilla.TextMatrix(I, 1)
                    justifica_printer 3.8, 13, Renglon, grdGrilla.TextMatrix(I, 1)
                    canttxt = Len(grdGrilla.TextMatrix(I, 1))
                    canttxt = canttxt / 36 'es para sacar la cantidad de renglones
                    canttxt = Int(canttxt)
                End If
                Imprimir 13.15, Renglon, False, grdGrilla.TextMatrix(I, 2) 'cantidad
                Imprimir 15.3, Renglon, False, grdGrilla.TextMatrix(I, 3) 'precio
                Imprimir 15.7, Renglon, False, grdGrilla.TextMatrix(I, 4) 'bonoficacion
                Imprimir 17.6, Renglon, False, grdGrilla.TextMatrix(I, 6) 'importe
                Renglon = Renglon + (canttxt * 0.5) + 0.5

            End If
        Next I
            '-----OBSERVACIONES---------------------
            If txtObservaciones.Text <> "" Then
                Imprimir 1, Renglon + 9.5, True, "Observ.: "
                'CortarCadena 3.2, Renglon + 9.5, Trim(txtObservaciones)
                justifica_printer 3.2, 14, Renglon + 9.5, Trim(txtObservaciones.Text)
            End If
            'Imprimir 0, 16.5, True, "texto de bajo del detalle"
            '-------------IMPRIMO TOTALES--------------------
'            Imprimir 17, 20.5, True, txtSubtotal.Text
'            Imprimir 17, 22.5, True, txtSubtotal.Text
            If txtPorcentajeBoni.Text <> "" Then
                If chkBonificaEnPesos.Value = Checked Then
                    Imprimir 14.8, 22.7, True, "   $" & txtPorcentajeBoni.Text
                    Imprimir 17.2, 22.7, True, txtImporteBoni.Text
                Else
                    Imprimir 14.8, 22.7, True, "    " & txtPorcentajeBoni.Text
                    Imprimir 17.2, 22.7, True, txtImporteBoni.Text
                End If
                'Imprimir 7.1, 15.5, True, txtSubTotalBoni.Text
            End If
            Imprimir 17.2, 21.9, True, txtSubtotal.Text
            If txtPorcentajeBoni.Text <> "" Then
                Imprimir 17.2, 23.8, True, txtSubTotalBoni.Text
            Else
                Imprimir 17.2, 23.8, True, txtSubtotal.Text
            End If
            Imprimir 14.8, 24.7, True, "    " & txtPorcentajeIva.Text
            Imprimir 17.2, 24.7, True, txtImporteIva.Text
            Imprimir 16.7, 26.4, True, txtTotal.Text

        Printer.EndDoc
    Next w
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Public Sub ImprimirEncabezado()
 '-----------IMPRIME EL ENCABEZADO DE LA NOTA DE CREDITO-------------------
    Dim a�o As String
    
    Imprimir 15.8, 1, False, "NOTA DE CR�DITO"
    'a�o = String(4, Year(FechaNotaCredito))
    a�o = Year(FechaNotaCredito)
    Imprimir 14.5, 3, False, Format(Day(FechaNotaCredito), "00")
    Imprimir 16.1, 3, False, Format(Month(FechaNotaCredito), "00")
    Imprimir 17.8, 3, False, Mid(a�o, 3, 2)
        
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CUIT,C.CLI_INGBRU, L.LOC_DESCRI"
    sql = sql & ", P.PRO_DESCRI,CI.IVA_CODIGO"
    sql = sql & " FROM CLIENTE C, LOCALIDAD L,"
    sql = sql & " PROVINCIA P, CONDICION_IVA CI"
    sql = sql & " WHERE"
    sql = sql & " C.CLI_CODIGO=" & XN(txtCodCliente)
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        If Len(Trim(Rec1!CLI_RAZSOC)) < 36 Then
            Imprimir 2.3, 5.8, True, Trim(Rec1!CLI_RAZSOC)
        Else
            CortarCadena 2.3, 5.4, Trim(Rec1!CLI_RAZSOC)
        End If
        Imprimir 12.3, 5.4, False, Trim(IIf(IsNull(Rec1!CLI_DOMICI), "", Rec1!CLI_DOMICI))
        'FACTURA
        'Imprimir 13.8, 7.2, True, Format(txtNroFactura.Text, "00000000") & " del " & Format(FechaFactura.Text, "dd/mm/yyyy")
        Imprimir 12.3, 5.8, False, Trim(Rec1!LOC_DESCRI) & " - " & Trim(Rec1!PRO_DESCRI)
'        'Imprimir 1, 6.3, False, Trim(Rec1!IVA_DESCRI)
        If Rec1!IVA_CODIGO = 1 Then
            Imprimir 3.75, 6.85, False, "X"
        Else
            If Rec1!IVA_CODIGO = 4 Then 'Exento
                Imprimir 10.1, 6.85, False, "X"
            Else
                Imprimir 7, 6.85, False, "X"
            End If
        End If
        Imprimir 13, 6.85, False, IIf(IsNull(Rec1!CLI_CUIT), "", Format(Rec1!CLI_CUIT, "##-########-#"))
        Imprimir 13.8, 7.5, False, IIf(IsNull(Rec1!CLI_INGBRU), "", Format(Rec1!CLI_INGBRU, "###-#####-##"))
    End If
    Rec1.Close

'    If cboCondicion.ItemData(cboCondicion.ListIndex) = 1 Then
'        Imprimir 7.4, 8.5, False, "X"
'    Else
'        Imprimir 10.5, 8.5, False, "X"
'    End If
    
    
'    Imprimir 1.8, 7.2, False, cboConcepto.Text
'    Imprimir 0, 8, False, "C�digo"
'    Imprimir 2.5, 8, False, "Descripci�n"
'    Imprimir 10, 8, False, "Cantidad"
'    Imprimir 13, 8, False, "Precio"
'    Imprimir 15, 8, False, "Bonof."
'    Imprimir 17, 8, False, "Importe"
End Sub

Private Sub CmdNuevo_Click()
   For I = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(I, 0) = ""
        grdGrilla.TextMatrix(I, 1) = ""
        grdGrilla.TextMatrix(I, 2) = ""
        grdGrilla.TextMatrix(I, 3) = ""
        grdGrilla.TextMatrix(I, 4) = ""
        grdGrilla.TextMatrix(I, 5) = ""
        grdGrilla.TextMatrix(I, 6) = ""
        grdGrilla.TextMatrix(I, 7) = ""
        grdGrilla.TextMatrix(I, 8) = ""
        grdGrilla.TextMatrix(I, 9) = I
   Next
   txtCodCliente.Text = ""
   txtNroSucursal.Text = ""
   txtNroNotaCredito.Text = ""
   FechaNotaCredito.Value = Date
   lblEstadoNotaCredito.Caption = ""
   txtSubtotal.Text = ""
   txtTotal.Text = ""
   txtPorcentajeBoni.Text = ""
   txtPorcentajeIva.Text = ""
   txtImporteBoni.Text = ""
   txtSubTotalBoni.Text = ""
   txtImporteIva.Text = ""
   txtObservaciones.Text = ""
   lblEstado.Caption = ""
   cboConcepto.ListIndex = 0
   cmdGrabar.Enabled = True
   'BUSCO IVA
   BuscoIva
   'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoNotaCredito) 'ESTADO PENDIENTE
    VEstadoNotaCredito = 1
    '--------------
    chkBonificaEnPorsentaje.Value = Unchecked
    chkBonificaEnPesos.Value = Unchecked
    FrameNotaCredito.Enabled = True
    FrameCliente.Enabled = True
    tabDatos.Tab = 0
    cboNotaCredito.ListIndex = 0
    cboNotaCredito.SetFocus
    
End Sub
Private Sub CortarCadena(COLUMNA As Double, Renglon As Double, Cadena As String)
    Dim salto, Max, inf, I, leer, leerb As Integer
    Dim salto1, salto2, salto3, salto4, salto5, salto6, salto7 As Integer
    Dim salto1b, salto2b, salto3b, salto4b, salto5b, salto6b, salto7b As Integer
    Dim cadena1 As String
    Dim cadena2 As String
    Dim cadena3 As String
    Dim cadena4 As String
    Dim cadena5 As String
    Dim cadena6 As String
    Dim cadena7 As String
    Dim cadena8 As String
    
    cadena1 = ""
    cadena2 = ""
    cadena3 = ""
    cadena4 = ""
    cadena5 = ""
    cadena6 = ""
    cadena7 = ""
    
    
    salto = 1
    Max = 36 * salto
    inf = Max - 10
    'falta = 0
    'If Len(cadena) > 35 Then
        For I = 1 To Len(Cadena)
            If (Mid(Cadena, I, 1) = " ") And (I > inf) And (I < Max) Or (I > Max) Then
                
                    If salto = 1 Then
                    salto1 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        cadena1 = Left(Cadena, I)
                        cadena2 = Mid(Cadena, salto1, Max)
                        'Imprimir 3.2, renglon, False, Left(grdGrilla.TextMatrix(I, 1), 36) 'descripcion
                        'Imprimir 3.2, renglon + 0.5, False, Mid(grdGrilla.TextMatrix(I, 1), 37, 36) 'descripcion
                    
                    Else
                        cadena1 = Left(Cadena, I)
                    End If
                      'descripcion
                End If
                If salto = 2 Then
                    leer = I - salto1
                    salto2 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto1b = I
                        leerb = Len(Cadena) + 1 - salto1b
                        cadena2 = Mid(Cadena, salto1, leer)
                        cadena3 = Mid(Cadena, salto1b, leerb)  'descripcion
                    Else
                        cadena2 = Mid(Cadena, salto1, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 3 Then
                    Max = 36 + I
                    inf = Max - 10
                    leer = I - salto2
                    salto3 = I
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto2b = I
                        leerb = Len(Cadena) + 1 - salto2b
                        cadena3 = Mid(Cadena, salto2, leer)
                        cadena4 = Mid(Cadena, salto2b, leerb)
                    Else
                        cadena3 = Mid(Cadena, salto2, leer)  'descripcion
                    End If
                    
                End If
                If salto = 4 Then
                    leer = I - salto3
                    salto4 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto3b = I
                        leerb = Len(Cadena) + 1 - salto3b
                        cadena4 = Mid(Cadena, salto3, leer)
                        cadena5 = Mid(Cadena, salto3b, leerb)  'descripcion
                    Else
                         cadena4 = Mid(Cadena, salto3, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 5 Then
                    leer = I - salto4
                    salto5 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto4b = I
                        leerb = Len(Cadena) + 1 - salto4b
                        cadena5 = Mid(Cadena, salto4, leer)
                        cadena6 = Mid(Cadena, salto4b, leerb)  'descripcion
                    Else
                        cadena5 = Mid(Cadena, salto4, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 6 Then
                    leer = I - salto5
                    salto6 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto5b = I
                        leerb = Len(Cadena) + 1 - salto5b
                        cadena6 = Mid(Cadena, salto5, leer)
                        cadena7 = Mid(Cadena, salto5b, leerb)  'descripcion
                        
                    Else
                        cadena6 = Mid(Cadena, salto5, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 7 Then
                    leer = I - salto6
                    salto7 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto6b = I
                        leerb = Len(Cadena) + 1 - salto6b
                        cadena7 = Mid(Cadena, salto6, leer)
                        cadena8 = Mid(Cadena, salto6b, leerb)  'descripcion
                        
                    Else
                        cadena7 = Mid(Cadena, salto6, leer)  'descripcion
                    End If
                    
                End If
                
                salto = salto + 1
                'Max = valor * salto
                'inf = Max - 10
                
            End If
        Next
    
        Imprimir COLUMNA, Renglon, False, cadena1
        Imprimir COLUMNA, Renglon + 0.5, False, cadena2
        Imprimir COLUMNA, Renglon + 1, False, cadena3
        Imprimir COLUMNA, Renglon + 1.5, False, cadena4
        Imprimir COLUMNA, Renglon + 2, False, cadena5
        Imprimir COLUMNA, Renglon + 2.5, False, cadena6
        Imprimir COLUMNA, Renglon + 3, False, cadena7
        Imprimir COLUMNA, Renglon + 3.5, False, cadena8
    'Else
    '    cadena1 = cadena
    '    MsgBox cadena1
    'End If
    
End Sub
Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmNotaCreditoCliente = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl.Name <> "grdGrilla" And _
        Me.ActiveControl.Name <> "txtEdit" And _
        KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    grdGrilla.FormatString = "C�digo|Descripci�n|Cantidad|Precio|Bonif.|Pre.Bonif.|Importe|Rubro|Linea|Orden"
    grdGrilla.ColWidth(0) = 1200  'CODIGO
    grdGrilla.ColWidth(1) = 4300 'DESCRIPCION
    grdGrilla.ColWidth(2) = 1000 'CANTIDAD
    grdGrilla.ColWidth(3) = 1000 'PRECIO
    grdGrilla.ColWidth(4) = 1000 'BONOFICACION
    grdGrilla.ColWidth(5) = 1000 'PRE BONIFICACION
    grdGrilla.ColWidth(6) = 1000 'IMPORTE
    grdGrilla.ColWidth(7) = 2100 'RUBRO
    grdGrilla.ColWidth(8) = 2100 'LINEA
    grdGrilla.ColWidth(9) = 0    'ORDEN
    grdGrilla.Cols = 10
    grdGrilla.Rows = 1
    For I = 2 To 14
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                             & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & (I - 1)
    Next
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "^Tipo|^N�mero|^Fecha|Importe|Cliente|Domicilio|Cod_Estado|" _
                              & "PORCENTAJE BONIFICA|PORCENTAJE IVA|" _
                              & "OBSERVACIONES|COD TIPO COMPROBANTE NOTA CREDITO|" _
                              & "COD CONCEPTO|BONIFICA EN PESOS|" _
                              & "CODIGO CLIENTE|REPRESENTADA"
    GrdModulos.ColWidth(0) = 1000 'TIPO NOTA CREDITO
    GrdModulos.ColWidth(1) = 1300 'NUMERO
    GrdModulos.ColWidth(2) = 1100 'FECHA
    GrdModulos.ColWidth(3) = 1100 'IMPORTE
    GrdModulos.ColWidth(4) = 4000 'CLIENTE
    GrdModulos.ColWidth(5) = 2800 'Domicilio
    GrdModulos.ColWidth(6) = 0    'COD_ESTADO
    GrdModulos.ColWidth(7) = 0    'PORCENTAJE BONIFICA
    GrdModulos.ColWidth(8) = 0    'PORCENTAJE IVA
    GrdModulos.ColWidth(9) = 0    'OBSERVACIONES
    GrdModulos.ColWidth(10) = 0    'COD TIPO COMPROBANTE NOTA CREDITO
    GrdModulos.ColWidth(11) = 0   'COD CONCEPTO
    GrdModulos.ColWidth(12) = 0   'BONIFICA EN PESOS
    GrdModulos.ColWidth(13) = 0   'CODIGO CLIENTE
    GrdModulos.ColWidth(14) = 0   'REPRESENTADA
    GrdModulos.Rows = 1
    GrdModulos.Cols = 15
    frameBuscar.Caption = "Buscar Nota de Cr�dito por..."
    '------------------------------------
    lblEstado.Caption = ""
    'CARGO COMBO CON LOS TIPOS DE NOTA DE CREDITO
    LlenarComboNotaCredito
    'CARGO COMBO CON LOS CONCEPTOS DE NOTA DE CREDITO
    LlenarComboConcepto
    'CARGO LISTA DE PRECIOS
    CargoCboListaPrecio ' maquinas
    CargoCboLPrecioRep 'repuestos
    
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoNotaCredito) 'ESTADO PENDIENTE
    VEstadoNotaCredito = 1
    FechaNotaCredito.Value = Date
    TipoBusquedaDoc = 1 'ESTO ES PARA BUSCAR FACTURA(1), (2)PARA BUSCAR REMITOS
    tabDatos.Tab = 0
    'BUSCO IVA
    BuscoIva
End Sub

Private Sub CargoCboListaPrecio() '' Lista de Precios de Repuestos
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 2"   '2: Accesorios
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboListaPrecio.AddItem rec!LIS_DESCRI
            cboListaPrecio.ItemData(cboListaPrecio.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboListaPrecio.ListIndex = 0
    End If
    rec.Close
End Sub
Private Sub CargoCboLPrecioRep() '' Lista de Precios de Repuestos
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 1"   '1: Repuestos
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboLPrecioRep.AddItem rec!LIS_DESCRI
            cboLPrecioRep.ItemData(cboLPrecioRep.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboLPrecioRep.ListIndex = 0
    End If
    rec.Close
End Sub


Private Sub BuscoIva()
    sql = "SELECT IVA FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtPorcentajeIva.Text = IIf(IsNull(rec!IVA), "", Format(rec!IVA, "0.00"))
    End If
    rec.Close
End Sub

Private Sub LlenarComboConcepto()
    sql = "SELECT * FROM CONCEPTO_NOTA_CREDITO"
    sql = sql & " ORDER BY CNC_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboConcepto.AddItem rec!CNC_DESCRI
            cboConcepto.ItemData(cboConcepto.NewIndex) = rec!CNC_CODIGO
            rec.MoveNext
        Loop
        cboConcepto.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboNotaCredito()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'NOTA DE CRE%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboNotaCredito.AddItem rec!TCO_DESCRI
            cboNotaCredito.ItemData(cboNotaCredito.NewIndex) = rec!TCO_CODIGO
            cboNotaCredito1.AddItem rec!TCO_DESCRI
            cboNotaCredito1.ItemData(cboNotaCredito1.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboNotaCredito.ListIndex = 0
        cboNotaCredito1.ListIndex = -1
    End If
    rec.Close
End Sub

Private Function BuscoUltimaNotaCredito(TipoNC As Integer) As String
    'ACA BUSCA EL NUMERO DE REMITO SIGUIENTE AL ULTIMO CARGADO
    sql = "SELECT (NOTA_CREDITO_A) + 1 AS NC_A, (NOTA_CREDITO_B) + 1 AS NC_B"
    sql = sql & " FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Select Case TipoNC
            Case 4
                BuscoUltimaNotaCredito = IIf(IsNull(rec!NC_A), 1, rec!NC_A)
            Case 5
                BuscoUltimaNotaCredito = IIf(IsNull(rec!NC_B), 1, rec!NC_B)
            Case 6
                MsgBox "No hay Notas de Cr�dito del tipo C", vbExclamation, TIT_MSGBOX
                cboNotaCredito.SetFocus
        End Select
    End If
    rec.Close
End Function

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1
            LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
            grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = grdGrilla.RowSel
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            txtTotal.Text = txtSubtotal.Text
            grdGrilla.Col = 0
        Case 4
            VBonificacion = 0
            grdGrilla.Text = ""
            grdGrilla.Col = 5
            grdGrilla.Text = ""
            VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)))
            grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Valido_Importe(CStr(VBonificacion))
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            txtTotal.Text = txtSubtotal.Text
            grdGrilla.Col = 4
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
        Case 1
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = "" Then
                chkBonificaEnPorsentaje.SetFocus
            End If
        Case 2
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "" Then
                grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "1"
                TxtEdit_KeyDown 13, 2
            End If
        End Select
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
       (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Or (grdGrilla.Col = 4) Then
        If KeyAscii = vbKeyReturn Then
            If grdGrilla.Col = 4 Then
                If grdGrilla.row < grdGrilla.Rows - 1 Then
                    grdGrilla.row = grdGrilla.row + 1
                    grdGrilla.Col = 0
                Else
                    SendKeys "{TAB}"
                End If
            Else
                grdGrilla.Col = grdGrilla.Col + 1
            End If
        Else
            If (grdGrilla.Col <> 1) Then
                'If KeyAscii > 47 And KeyAscii < 58 Then
                    EDITAR grdGrilla, txtEdit, KeyAscii
                'End If
            Else
                EDITAR grdGrilla, txtEdit, KeyAscii
            End If
        End If
    End If
End Sub

Private Sub grdGrilla_LeaveCell()
    If txtEdit.Visible = False Then Exit Sub
    grdGrilla = txtEdit.Text
    txtEdit.Visible = False
End Sub

Private Sub grdGrilla_GotFocus()
    If grdGrilla.Rows > 1 Then
        If txtEdit.Visible = False Then
            grdGrilla.Col = 0
            Exit Sub
        End If
        grdGrilla = txtEdit.Text
        txtEdit.Visible = False
    End If
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        
        Set Rec1 = New ADODB.Recordset
        lblEstado.Caption = "Buscando..."
        Screen.MousePointer = vbHourglass
        'CABEZA NOTA CREDITO
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 10)), cboNotaCredito)
        
        txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        txtNroNotaCredito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        FechaNotaCredito.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
        Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6)), lblEstadoNotaCredito)
        VEstadoNotaCredito = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6))
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 9) <> "" Then
            txtObservaciones.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 9))
        End If
        txtCodCliente.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 13)
        txtCodCliente_LostFocus
        
        'CONDICION NOTA CREDITO (CONSEPTO)
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 11)), cboConcepto)
        
        '----BUSCO DETALLE DE LA NOTA DE CREDITO------------------
        sql = "SELECT DNC.*, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
        sql = sql & " FROM DETALLE_NOTA_CREDITO_CLIENTE DNC, PRODUCTO P, RUBROS R, LINEAS L"
        sql = sql & " WHERE DNC.NCC_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8))
        sql = sql & " AND DNC.NCC_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
        sql = sql & " AND DNC.NCC_FECHA=" & XDQ(GrdModulos.TextMatrix(GrdModulos.RowSel, 2))
        sql = sql & " AND DNC.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 10))
        sql = sql & " AND DNC.PTO_CODIGO=P.PTO_CODIGO"
        sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
        sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
        sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
        sql = sql & " ORDER BY DNC.DNC_NROITEM"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            I = 1
            Do While Rec1.EOF = False
                grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
                grdGrilla.TextMatrix(I, 1) = IIf(IsNull(Rec1!DNC_DETALLE), Rec1!PTO_DESCRI, Rec1!DNC_DETALLE)
                grdGrilla.TextMatrix(I, 2) = Rec1!DNC_CANTIDAD
                grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DNC_PRECIO)
                If IsNull(Rec1!DNC_BONIFICA) Then
                    grdGrilla.TextMatrix(I, 4) = ""
                Else
                    grdGrilla.TextMatrix(I, 4) = Valido_Importe(Rec1!DNC_BONIFICA)
                End If
                VBonificacion = 0
                If Not IsNull(Rec1!DNC_BONIFICA) Then
                    VBonificacion = (((CDbl(Rec1!DNC_CANTIDAD) * CDbl(Rec1!DNC_PRECIO)) * CDbl(Rec1!DNC_BONIFICA)) / 100)
                    VBonificacion = ((CDbl(Rec1!DNC_CANTIDAD) * CDbl(Rec1!DNC_PRECIO)) - VBonificacion)
                    grdGrilla.TextMatrix(I, 5) = Valido_Importe(CStr(VBonificacion))
                    grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
                Else
                    VBonificacion = (CDbl(Rec1!DNC_CANTIDAD) * CDbl(Rec1!DNC_PRECIO))
                    grdGrilla.TextMatrix(I, 5) = ""
                    grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
                End If
                grdGrilla.TextMatrix(I, 7) = Rec1!RUB_DESCRI
                grdGrilla.TextMatrix(I, 8) = Rec1!LNA_DESCRI
                grdGrilla.TextMatrix(I, 9) = Rec1!DNC_NROITEM
                I = I + 1
                Rec1.MoveNext
            Loop
            VBonificacion = 0
        End If
        Rec1.Close
        
        '--CARGO LOS TOTALES----
        txtSubtotal.Text = Valido_Importe(SumaBonificacion)
        txtTotal.Text = txtSubtotal.Text
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 12) = "S" Then
            chkBonificaEnPesos.Value = Checked
        ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 12) = "N" Then
            chkBonificaEnPorsentaje.Value = Checked
        Else
            chkBonificaEnPesos.Value = Unchecked
            chkBonificaEnPorsentaje.Value = Unchecked
        End If
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 7) <> "" Then
            txtPorcentajeBoni.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 7)
            txtPorcentajeBoni_LostFocus
        End If
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 8) <> "" Then
            txtPorcentajeIva = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
            txtPorcentajeIva_LostFocus
        End If
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        '--------------
        FrameNotaCredito.Enabled = False
        FrameCliente.Enabled = False
        '--------------
        tabDatos.Tab = 0
        cboConcepto.SetFocus
        '----------------------------------------------------------
    End If
End Sub

Private Function BuscarTipoDocAbre(Codigo As String) As String
    sql = "SELECT TCO_ABREVIA"
    sql = sql & " FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_CODIGO=" & XN(Codigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscarTipoDocAbre = rec!TCO_ABREVIA
    Else
        BuscarTipoDocAbre = ""
    End If
    rec.Close
End Function
Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    GrdModulos.Rows = 2
    txtCliente.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    cboNotaCredito1.Enabled = False
    cmdBuscarCli.Enabled = False
    cmdGrabar.Enabled = False
    LimpiarBusqueda
    If Me.Visible = True Then chkCliente.SetFocus
  Else
    If VEstadoNotaCredito = 1 Then
        cmdGrabar.Enabled = True
    Else
        cmdGrabar.Enabled = False
    End If
  End If
End Sub

Private Sub LimpiarBusqueda()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    GrdModulos.Rows = 1
    chkCliente.Value = Unchecked
    chkFecha.Value = Unchecked
    chkTipoFactura.Value = Unchecked
    
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
    If chkFecha.Value = Unchecked _
        And chkTipoFactura.Value = Unchecked _
        And ActiveControl.Name <> "cmdBuscarCli" _
        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub

Private Function BuscoCondicionIVA(IVACodigo As String) As String
    sql = "SELECT * FROM CONDICION_IVA"
    sql = sql & " WHERE IVA_CODIGO=" & XN(IVACodigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscoCondicionIVA = rec!IVA_DESCRI
    Else
        BuscoCondicionIVA = ""
    End If
    rec.Close
End Function

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    'If grdGrilla.Col = 0 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 0 Then
        'CarTexto KeyAscii
        txtEdit.MaxLength = 16
    End If
    If grdGrilla.Col = 1 Then
        txtEdit.MaxLength = 50
    End If
    If grdGrilla.Col = 2 Then
        KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
        txtEdit.MaxLength = 10
    End If
    If grdGrilla.Col = 3 Then
        KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
        txtEdit.MaxLength = 10
    End If
'    CarTexto KeyAscii

'    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    'If grdGrilla.Col = 0 Then KeyAscii = CarNumeroEntero(KeyAscii)
'    If grdGrilla.Col = 2 Then KeyAscii = CarNumeroEntero(KeyAscii)
'    If grdGrilla.Col = 3 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    If grdGrilla.Col = 4 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    CarTexto KeyAscii
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    VBonificacion = 0
    If KeyCode = vbKeyF1 Then
        frmBuscar.TipoBusqueda = 2
        frmBuscar.CodListaPrecio = 0
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        frmBuscar.Show vbModal
    End If

    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
            
            Case 0, 1 'PRODUCTO Y DESCRIPCION
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                txtEdit.Text = Replace(txtEdit.Text, "'", "�")
                If lblEstadoNotaCredito.Caption = "PENDIENTE" Then
                    sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI, P.PTO_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI"
                    sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, TIPO_PRESENTACION RE"
                    sql = sql & " WHERE"
                    If grdGrilla.Col = 0 Then
                        sql = sql & " P.PTO_CODIGO LIKE '" & txtEdit.Text & "'"
                    Else
                        sql = sql & " P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
                    End If
                    sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
                    sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                    sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
                Else
                    sql = "SELECT DNC.*,P.PTO_CODIGO,P.PTO_DESCRI, P.PTO_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI"
                    sql = sql & " FROM DETALLE_NOTA_CREDITO_CLIENTE DNC,PRODUCTO P, RUBROS R, LINEAS L, TIPO_PRESENTACION RE"
                    sql = sql & " WHERE DNC.PTO_CODIGO = P.PTO_CODIGO"
                    If grdGrilla.Col = 0 Then
                        sql = sql & " AND P.PTO_CODIGO LIKE '" & txtEdit.Text & "'"
                    Else
                        sql = sql & " AND  P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
                    End If
                    sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
                    sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                    sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
                End If
                rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                
                If rec.EOF = False Then
                    If rec.RecordCount > 1 Then
                        grdGrilla.SetFocus
                        frmBuscar.TipoBusqueda = 2
                        'LE DIGO EN QUE LISTA DE PRECIO BUSCAR LOS PRECIOS
                        frmBuscar.CodListaPrecio = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
                        frmBuscar.TxtDescriB.Text = txtEdit.Text
                        frmBuscar.Show vbModal
                        grdGrilla.Col = 0
                        grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
                        grdGrilla.Col = 1
                        grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
                        grdGrilla.Col = 3
                        grdGrilla.Text = Valido_Importe(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2))
                        grdGrilla.Col = 7
                        grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
                        grdGrilla.Col = 8
                        grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = grdGrilla.RowSel
                        grdGrilla.Col = 2
                    Else
                        grdGrilla.Col = 0
                        grdGrilla.Text = Trim(rec!PTO_CODIGO)
                        If lblEstadoNotaCredito.Caption = "PENDIENTE" Then
                            grdGrilla.Col = 1
                            grdGrilla.Text = Trim(rec!PTO_DESCRI)
                        Else
                            grdGrilla.Col = 1
                            grdGrilla.Text = Trim(rec!DNC_DETALLE)
                        End If
                        
                        grdGrilla.Col = 3
                        grdGrilla.Text = Valido_Importe(Trim(rec!PTO_PRECIO))
                        grdGrilla.Col = 7
                        grdGrilla.Text = Trim(rec!RUB_DESCRI)
                        grdGrilla.Col = 8
                        grdGrilla.Text = Trim(rec!LNA_DESCRI)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = grdGrilla.RowSel
                        grdGrilla.Col = 2
                    End If
                        If BuscoRepetetidos(grdGrilla.TextMatrix(grdGrilla.RowSel, 0), grdGrilla.RowSel) = False Then
                            grdGrilla.Col = 0
                            grdGrilla_KeyDown vbKeyDelete, 0
                        End If
                Else
                        MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
                        txtEdit.Text = ""
                        LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                        grdGrilla.Col = 0
                End If
                rec.Close
                Screen.MousePointer = vbNormal
   
            Case 2 'CANTIDAD
            
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                    If Trim(txtEdit) = "" Then txtEdit.Text = "1"
                    If txtEdit.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 0) Then
                        VBonificacion = (CInt(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)))
                    Else
                            VBonificacion = (CInt(txtEdit.Text) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)))
                    End If
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Valido_Importe(CStr(VBonificacion))
                    If grdGrilla.TextMatrix(grdGrilla.RowSel, 4) <> "" Then
                        VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 4))) / 100)
                        VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6)) - VBonificacion)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Valido_Importe(CStr(VBonificacion))
                    End If
                    txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                    txtTotal.Text = txtSubtotal.Text
                Else
                    txtEdit.Text = "1"
                End If
                
            Case 3 'PRECIO
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                    If Trim(txtEdit) = "" Then txtEdit.Text = "1"
                    If grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "" Then grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "1"
                    
                    VBonificacion = (CInt(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(txtEdit.Text))
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Valido_Importe(CStr(VBonificacion))
                    If grdGrilla.TextMatrix(grdGrilla.RowSel, 4) <> "" Then
                        VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 4))) / 100)
                        VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6)) - VBonificacion)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Valido_Importe(CStr(VBonificacion))
                    End If
                    txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                    txtTotal.Text = txtSubtotal.Text
                Else
                    txtEdit.Text = ""
                End If

            Case 4 'BONIFICACION
                If Trim(txtEdit) <> "" Then
                    If txtEdit.Text = ValidarPorcentaje(txtEdit) = False Then
                        Exit Sub
                    End If
                    VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6)) * CDbl(txtEdit.Text)) / 100)
                    VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6)) - VBonificacion)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Valido_Importe(CStr(VBonificacion))
                    txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                    txtTotal.Text = txtSubtotal.Text
                Else
                    MsgBox "Debe ingresar el Importe", vbExclamation, TIT_MSGBOX
                    grdGrilla.Col = 4
                End If
        End Select
        grdGrilla.SetFocus
    End If
    If KeyCode = vbKeyEscape Then
       txtEdit.Visible = False
       grdGrilla.SetFocus
    End If
End Sub

Private Function BuscoRepetetidos(Codigo As String, Linea As Integer) As Boolean
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            If Codigo = grdGrilla.TextMatrix(I, 0) And (I <> Linea) Then
                MsgBox "El producto ya fue elegido anteriormente", vbExclamation, TIT_MSGBOX
                BuscoRepetetidos = False
                Exit Function
            End If
        End If
    Next
    BuscoRepetetidos = True
End Function

Private Function BuscoVendedor(Codigo As String) As String
    sql = "SELECT VEN_NOMBRE"
    sql = sql & " FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO=" & XN(Codigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        BuscoVendedor = Trim(rec!VEN_NOMBRE)
    Else
        BuscoVendedor = "No se encontro el Vendedor"
    End If
    rec.Close
End Function

Private Function BuscoCliente(Cli As String) As String
    sql = "SELECT C.CLI_CODIGO, C.CLI_RAZSOC, C.CLI_DOMICI, P.PRO_DESCRI, L.LOC_DESCRI"
    sql = sql & ",C.CLI_CUIT,C.CLI_INGBRU,CI.IVA_DESCRI "
    sql = sql & " FROM CLIENTE C,  PROVINCIA P, LOCALIDAD L,CONDICION_IVA CI"
    sql = sql & " WHERE"
    If txtCodCliente.Text <> "" Then
        sql = sql & " C.CLI_CODIGO=" & XN(Cli)
    Else
        sql = sql & " C.CLI_RAZSOC LIKE '" & Cli & "%'"
    End If
    sql = sql & " AND C.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    BuscoCliente = sql
End Function

'Private Function BuscoCliente(Codigo As String) As String
'        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
'        sql = sql & " WHERE CLI_CODIGO=" & XN(Codigo)
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            BuscoCliente = rec!CLI_RAZSOC
'        Else
'            BuscoCliente = "No se encontro el Cliente"
'        End If
'        rec.Close
'End Function

Private Function BuscoSucursal(CodigoSuc As String, CodigoCli As String) As String
        sql = "SELECT * FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(CodigoSuc)
        sql = sql & " AND CLI_CODIGO=" & XN(CodigoCli)
        
        Set Rec1 = New ADODB.Recordset
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            BuscoSucursal = Rec1!SUC_DESCRI
        Else
            BuscoSucursal = "No se encontro la Sucursal"
        End If
        Rec1.Close
End Function

'Private Sub txtNroFactura_GotFocus()
'    SelecTexto txtNroFactura
'End Sub
'
'Private Sub txtNroFactura_KeyPress(KeyAscii As Integer)
'    KeyAscii = CarNumeroEntero(KeyAscii)
'End Sub
'
'Private Sub txtNroFactura_LostFocus()
'    If txtNroFactura.Text <> "" Then
'
'        Set Rec2 = New ADODB.Recordset
'        sql = "SELECT FC.FCL_NUMERO, FC.FCL_FECHA, FC.FCL_BONIFICA, FC.FCL_IVA, FC.FCL_BONIPESOS,"
'        sql = sql & "RC.EST_CODIGO, RC.STK_CODIGO, E.EST_DESCRI"
'        sql = sql & " ,NP.CLI_CODIGO, NP.SUC_CODIGO, NP.VEN_CODIGO"
'        sql = sql & " FROM FACTURA_CLIENTE FC, REMITO_CLIENTE RC, NOTA_PEDIDO NP, ESTADO_DOCUMENTO E"
'        sql = sql & " WHERE FC.FCL_NUMERO=" & XN(txtNroFactura)
'        If FechaFactura.Text <> "" Then
'            sql = sql & " AND FC.FCL_FECHA=" & XDQ(FechaFactura)
'        End If
'        'sql = sql & " AND FC.TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
'        sql = sql & " AND FC.RCL_NUMERO=RC.RCL_NUMERO"
'        sql = sql & " AND FC.RCL_FECHA=RC.RCL_FECHA"
'        sql = sql & " AND RC.NPE_NUMERO=NP.NPE_NUMERO"
'        sql = sql & " AND RC.NPE_FECHA=NP.NPE_FECHA"
'        sql = sql & " AND FC.EST_CODIGO=E.EST_CODIGO"
'
'        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'        If Rec2.EOF = False Then
'            If Rec2.RecordCount > 1 Then
'                MsgBox "Hay mas de una Factura con el N�mero: " & txtNroFactura.Text, vbInformation, TIT_MSGBOX
'                Rec2.Close
'                cmdBuscarFactura_Click
'                Exit Sub
'            End If
'            Screen.MousePointer = vbHourglass
'            lblEstado.Caption = "Buscando..."
'
'            'CARGO CABECERA DE LA FACTURA
'            FechaFactura.Text = Rec2!FCL_FECHA
'            grillaFactura.TextMatrix(0, 1) = BuscoCliente(Rec2!CLI_CODIGO)
'            grillaFactura.TextMatrix(1, 1) = BuscoSucursal(Rec2!SUC_CODIGO, Rec2!CLI_CODIGO)
'            grillaFactura.TextMatrix(2, 1) = BuscoVendedor(Rec2!VEN_CODIGO)
'            grillaFactura.TextMatrix(0, 2) = Rec2!CLI_CODIGO
'            If Rec2!EST_CODIGO = 2 Then
'                MsgBox "La Factura n�mero: " & txtNroFactura.Text & Chr(13) & Chr(13) & _
'                       "No puede ser asignado a la Nota de Cr�dito por su estado (" & Rec2!EST_DESCRI & ")", vbExclamation, TIT_MSGBOX
'                cmdGrabar.Enabled = False
'                Screen.MousePointer = vbNormal
'                lblEstado.Caption = ""
'                Rec2.Close
'                LimpiarFactura
'                Exit Sub
'            Else
'                cmdGrabar.Enabled = True
'            End If
'
'            '----BUSCO DETALLE DE LA FACTURA------------------
'            sql = "SELECT DFC.*, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
'            sql = sql & " FROM DETALLE_FACTURA_CLIENTE DFC, PRODUCTO P, RUBROS R, LINEAS L"
'            sql = sql & " WHERE DFC.FCL_NUMERO=" & XN(txtNroFactura)
'            sql = sql & " AND DFC.FCL_FECHA=" & XDQ(FechaFactura)
'            'sql = sql & " AND DFC.TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
'            sql = sql & " AND DFC.PTO_CODIGO=P.PTO_CODIGO"
'            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
'            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
'            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
'            sql = sql & " ORDER BY DFC.DFC_NROITEM"
'            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'            If Rec1.EOF = False Then
'                I = 1
'                Do While Rec1.EOF = False
'                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
'                    grdGrilla.TextMatrix(I, 1) = Rec1!PTO_DESCRI
'                    grdGrilla.TextMatrix(I, 2) = Rec1!DFC_CANTIDAD
'                    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DFC_PRECIO)
'                    If IsNull(Rec1!DFC_BONIFICA) Then
'                        grdGrilla.TextMatrix(I, 4) = ""
'                    Else
'                        grdGrilla.TextMatrix(I, 4) = Valido_Importe(Rec1!DFC_BONIFICA)
'                    End If
'                    VBonificacion = 0
'                    If Not IsNull(Rec1!DFC_BONIFICA) Then
'                        VBonificacion = (((CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO)) * CDbl(Rec1!DFC_BONIFICA)) / 100)
'                        VBonificacion = ((CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO)) - VBonificacion)
'                        grdGrilla.TextMatrix(I, 5) = Valido_Importe(CStr(VBonificacion))
'                        grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
'                    Else
'                        VBonificacion = (CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO))
'                        grdGrilla.TextMatrix(I, 5) = ""
'                        grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
'                    End If
'                    grdGrilla.TextMatrix(I, 7) = Rec1!RUB_DESCRI
'                    grdGrilla.TextMatrix(I, 8) = Rec1!LNA_DESCRI
'                    grdGrilla.TextMatrix(I, 9) = Rec1!DFC_NROITEM
'                    I = I + 1
'                    Rec1.MoveNext
'                Loop
'                VBonificacion = 0
'            End If
'            Rec1.Close
'            '--CARGO LOS TOTALES----
'            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
'            txtTotal.Text = txtSubtotal.Text
'            If Rec2!FCL_BONIPESOS = "S" Then
'                chkBonificaEnPesos.Value = Checked
'            ElseIf Rec2!FCL_BONIPESOS = "N" Then
'                chkBonificaEnPorsentaje.Value = Checked
'            Else
'                chkBonificaEnPesos.Value = Unchecked
'                chkBonificaEnPorsentaje.Value = Unchecked
'            End If
'            If Not IsNull(Rec2!FCL_BONIFICA) Then
'                txtPorcentajeBoni.Text = Rec2!FCL_BONIFICA
'                txtPorcentajeBoni_LostFocus
'            End If
'            If Not IsNull(Rec2!FCL_IVA) Then
'                txtPorcentajeIva = Rec2!FCL_IVA
'                txtPorcentajeIva_LostFocus
'            End If
'            Rec2.Close
'            lblEstado.Caption = ""
'            Screen.MousePointer = vbNormal
'            '--------------
'            FrameNotaCredito.Enabled = False
'            FrameFactura.Enabled = False
'            '--------------
'        Else
'            MsgBox "La Factura no existe", vbExclamation, TIT_MSGBOX
'            If Rec2.State = 1 Then Rec2.Close
'            LimpiarFactura
'        End If
'    End If
'End Sub

Private Function SumaTotal() As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 6) <> "" Then
            VTotal = VTotal + (CInt(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3)))
        End If
    Next
    SumaTotal = Valido_Importe(CStr(VTotal))
End Function

Private Function SumaBonificacion() As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 6) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(I, 6))
        End If
    Next
    SumaBonificacion = Valido_Importe(CStr(VTotal))
End Function

Private Sub txtNroNotaCredito_GotFocus()
    SelecTexto txtNroNotaCredito
End Sub

Private Sub txtNroNotaCredito_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaCredito_LostFocus()
    If txtNroNotaCredito.Text = "" Then
        'BUSCO EL NUMERO DE FACTURA QUE CORRESPONDE
        txtNroNotaCredito.Text = Format(BuscoUltimaNotaCredito(cboNotaCredito.ItemData(cboNotaCredito.ListIndex)), "00000000")
    Else
        txtNroNotaCredito.Text = Format(txtNroNotaCredito.Text, "00000000")
    End If
End Sub

Private Sub txtNroSucursal_GotFocus()
    SelecTexto txtNroSucursal
End Sub

Private Sub txtNroSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroSucursal_LostFocus()
    If txtNroSucursal.Text = "" Then
        txtNroSucursal.Text = Sucursal
    Else
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    End If
End Sub

Private Sub txtObservaciones_GotFocus()
    SelecTexto txtObservaciones
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtPorcentajeBoni_GotFocus()
    SelecTexto txtPorcentajeBoni
End Sub

Private Sub txtPorcentajeBoni_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorcentajeBoni, KeyAscii)
End Sub

Private Sub txtPorcentajeBoni_LostFocus()
    If txtPorcentajeBoni.Text <> "" And txtSubtotal.Text <> "" Then
        If chkBonificaEnPorsentaje.Value = Checked Then
            If ValidarPorcentaje(txtPorcentajeBoni) = False Then
                txtPorcentajeBoni.SetFocus
                Exit Sub
            End If
            txtImporteBoni.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeBoni.Text)) / 100
            txtImporteBoni.Text = Valido_Importe(txtImporteBoni.Text)
            txtTotal.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
            txtSubTotalBoni.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtSubTotalBoni.Text = Valido_Importe(txtSubTotalBoni.Text)
        ElseIf chkBonificaEnPesos.Value = Checked Then
            txtPorcentajeBoni.Text = Valido_Importe(txtPorcentajeBoni.Text)
            txtImporteBoni.Text = Valido_Importe(txtPorcentajeBoni.Text)
            txtTotal.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
            txtSubTotalBoni.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtSubTotalBoni.Text = Valido_Importe(txtSubTotalBoni.Text)
        Else
            txtPorcentajeBoni.Text = ""
            txtImporteBoni.Text = ""
            MsgBox "Debe elegir como bonifica", vbExclamation, TIT_MSGBOX
            chkBonificaEnPorsentaje.SetFocus
        End If
    End If
End Sub

Private Sub txtPorcentajeIva_GotFocus()
    SelecTexto txtPorcentajeIva
End Sub

Private Sub txtPorcentajeIva_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorcentajeIva, KeyAscii)
End Sub

Private Sub txtPorcentajeIva_LostFocus()
    If txtPorcentajeIva.Text <> "" And txtSubtotal.Text <> "" Then
        If ValidarPorcentaje(txtPorcentajeIva) = False Then
            txtPorcentajeIva.SetFocus
            Exit Sub
        End If
        If txtImporteBoni.Text <> "" Then
            txtImporteIva.Text = (CDbl(txtSubTotalBoni.Text) * CDbl(txtPorcentajeIva.Text)) / 100
            txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
            txtTotal.Text = CDbl(txtSubTotalBoni.Text) + CDbl(txtImporteIva.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
        Else
            If cboNotaCredito.ItemData(cboNotaCredito.ListIndex) = 4 Then
                txtImporteIva.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
                txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
                txtTotal.Text = CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text)
                txtTotal.Text = Valido_Importe(txtTotal.Text)
            Else
                txtImporteIva.Text = ""
                txtSubtotal.Text = Valido_Importe(SumaTotal)
                txtTotal.Text = Valido_Importe(SumaTotal)
            End If
        End If
    End If
End Sub

