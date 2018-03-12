VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form frmNotaDeditoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Dédito Clientes..."
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   8535
      TabIndex        =   12
      Top             =   7695
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10305
      TabIndex        =   14
      Top             =   7695
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   7635
      TabIndex        =   11
      Top             =   7695
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   9420
      TabIndex        =   13
      Top             =   7695
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7650
      Left            =   60
      TabIndex        =   30
      Top             =   15
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   13494
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
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
      TabPicture(0)   =   "frmNotaDeditoCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameNotaCredito"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameFactura"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmNotaDeditoCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame FrameFactura 
         Caption         =   "Factura..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   4050
         TabIndex        =   50
         Top             =   360
         Width           =   6990
         Begin VB.ComboBox cboFactura 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   255
            Width           =   2400
         End
         Begin VB.TextBox txtCodigoStock 
            Height          =   300
            Left            =   4995
            TabIndex        =   72
            Top             =   255
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.CommandButton cmdBuscarFactura 
            Height          =   315
            Left            =   3630
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoCliente.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Buscar Factura"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtNroFactura 
            Height          =   300
            Left            =   1110
            MaxLength       =   8
            TabIndex        =   2
            Top             =   615
            Width           =   1155
         End
         Begin FechaCtl.Fecha FechaFactura 
            Height          =   285
            Left            =   3015
            TabIndex        =   3
            Top             =   615
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin MSFlexGridLib.MSFlexGrid grillaFactura 
            Height          =   900
            Left            =   255
            TabIndex        =   51
            Top             =   930
            Width           =   6450
            _ExtentX        =   11377
            _ExtentY        =   1588
            _Version        =   393216
            Rows            =   3
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   260
            BackColor       =   12648447
            BackColorBkg    =   -2147483633
            GridLines       =   0
            GridLinesFixed  =   1
            ScrollBars      =   0
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
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   630
            TabIndex        =   74
            Top             =   285
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   2400
            TabIndex        =   54
            Top             =   630
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   390
            TabIndex        =   52
            Top             =   630
            Width           =   600
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
         Height          =   2010
         Left            =   -74610
         TabIndex        =   35
         Top             =   540
         Width           =   10410
         Begin VB.ComboBox cboNotaCredito1 
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1605
            Width           =   2400
         End
         Begin VB.CheckBox chkTipoFactura 
            Caption         =   "Tipo"
            Height          =   195
            Left            =   300
            TabIndex        =   20
            Top             =   1485
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoCliente.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Buscar Cliente"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarSuc 
            Height          =   315
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoCliente.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Buscar Sucursal"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3360
            TabIndex        =   23
            Top             =   945
            Width           =   990
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
            Left            =   4410
            TabIndex        =   42
            Top             =   960
            Width           =   5055
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   300
            TabIndex        =   18
            Top             =   959
            Width           =   1020
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1680
            Left            =   9690
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoCliente.frx":0956
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Buscar "
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   5865
            TabIndex        =   25
            Top             =   1290
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   3360
            TabIndex        =   24
            Top             =   1290
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
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
            Left            =   4845
            MaxLength       =   50
            TabIndex        =   37
            Tag             =   "Descripción"
            Top             =   255
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3360
            MaxLength       =   40
            TabIndex        =   21
            Top             =   255
            Width           =   975
         End
         Begin VB.TextBox txtSucursal 
            Height          =   300
            Left            =   3360
            MaxLength       =   40
            TabIndex        =   22
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtDesSuc 
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
            Left            =   4845
            MaxLength       =   50
            TabIndex        =   36
            Tag             =   "Descripción"
            Top             =   600
            Width           =   4620
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   300
            TabIndex        =   19
            Top             =   1221
            Width           =   810
         End
         Begin VB.CheckBox chkSucursal 
            Caption         =   "Sucursal"
            Height          =   195
            Left            =   300
            TabIndex        =   17
            Top             =   697
            Width           =   960
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   300
            TabIndex        =   16
            Top             =   435
            Width           =   855
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   2910
            TabIndex        =   71
            Top             =   1650
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   2535
            TabIndex        =   43
            Top             =   975
            Width           =   735
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4815
            TabIndex        =   41
            Top             =   1335
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2265
            TabIndex        =   40
            Top             =   1320
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
            Left            =   2745
            TabIndex        =   39
            Top             =   300
            Width           =   525
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2610
            TabIndex        =   38
            Top             =   645
            Width           =   660
         End
      End
      Begin VB.Frame FrameNotaCredito 
         Caption         =   "Nota de Credito..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   105
         TabIndex        =   32
         Top             =   360
         Width           =   3930
         Begin VB.ComboBox cboNotaCredito 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   330
            Width           =   2400
         End
         Begin VB.TextBox txtNroNotaCredito 
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
            Height          =   330
            Left            =   1170
            TabIndex        =   29
            Top             =   690
            Width           =   1155
         End
         Begin FechaCtl.Fecha FechaNotaCredito 
            Height          =   285
            Left            =   1170
            TabIndex        =   15
            Top             =   1065
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   750
            TabIndex        =   58
            Top             =   345
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   615
            TabIndex        =   55
            Top             =   1095
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   510
            TabIndex        =   49
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   570
            TabIndex        =   48
            Top             =   1530
            Width           =   540
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
            TabIndex        =   47
            Top             =   1545
            Width           =   1890
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4740
         Left            =   -74625
         TabIndex        =   28
         Top             =   2565
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8361
         _Version        =   393216
         Cols            =   15
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame4 
         Height          =   540
         Left            =   105
         TabIndex        =   56
         Top             =   2175
         Width           =   10935
         Begin VB.ComboBox cboConcepto 
            Height          =   315
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   165
            Width           =   4485
         End
         Begin VB.Label lblConcepto 
            AutoSize        =   -1  'True
            Caption         =   "Concepto:"
            Height          =   195
            Left            =   360
            TabIndex        =   73
            Top             =   210
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4950
         Left            =   105
         TabIndex        =   33
         Top             =   2640
         Width           =   10935
         Begin VB.CheckBox chkBonificaEnPesos 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en $"
            Height          =   285
            Left            =   390
            TabIndex        =   7
            Top             =   4200
            Width           =   1290
         End
         Begin VB.CheckBox chkBonificaEnPorsentaje 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en % "
            Height          =   285
            Left            =   390
            TabIndex        =   6
            Top             =   3900
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
            TabIndex        =   69
            Top             =   4230
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
            TabIndex        =   66
            Top             =   4230
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeIva 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6900
            TabIndex        =   9
            Top             =   3900
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
            TabIndex        =   63
            Top             =   4230
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeBoni 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2850
            TabIndex        =   8
            Top             =   3900
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
            TabIndex        =   60
            Top             =   4230
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
            TabIndex        =   59
            Top             =   3900
            Width           =   1350
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   1455
            MaxLength       =   60
            TabIndex        =   10
            Top             =   4575
            Width           =   8865
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   480
            TabIndex        =   34
            Top             =   480
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3735
            Left            =   75
            TabIndex        =   5
            Top             =   120
            Width           =   10725
            _ExtentX        =   18918
            _ExtentY        =   6588
            _Version        =   393216
            Rows            =   3
            Cols            =   11
            FixedCols       =   0
            BackColorSel    =   16777215
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   2
            HighLight       =   0
            ScrollBars      =   1
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
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   4110
            TabIndex        =   70
            Top             =   4290
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   6270
            TabIndex        =   68
            Top             =   4275
            Width           =   570
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "% I.V.A.:"
            Height          =   195
            Left            =   6240
            TabIndex        =   67
            Top             =   3930
            Width           =   600
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   2235
            TabIndex        =   65
            Top             =   4275
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Bonificación:"
            Height          =   195
            Left            =   1890
            TabIndex        =   64
            Top             =   3930
            Width           =   915
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   8505
            TabIndex        =   62
            Top             =   4275
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   8175
            TabIndex        =   61
            Top             =   3930
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   210
            TabIndex        =   57
            Top             =   4620
            Width           =   1110
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   31
         Top             =   570
         Width           =   1065
      End
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
      Height          =   345
      Left            =   225
      TabIndex        =   46
      Top             =   7755
      Width           =   750
   End
End
Attribute VB_Name = "frmNotaDeditoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim W As Integer
Dim TipoBusquedaDoc As Integer
Dim VBonificacion As Double
Dim VTotal As Double
Dim VEstadoNotaCredito As Integer

Private Sub cboNotaCredito_LostFocus()
    'BUSCO EL NUMERO DE NOTA DE CREDITO QUE CORRESPONDE
    txtNroNotaCredito.Text = BuscoUltimaNotaCredito(cboNotaCredito.ItemData(cboNotaCredito.ListIndex))
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

Private Sub chkSucursal_Click()
    If chkSucursal.Value = Checked Then
        txtSucursal.Enabled = True
        cmdBuscarSuc.Enabled = True
    Else
        txtSucursal.Enabled = False
        cmdBuscarSuc.Enabled = False
    End If
End Sub

Private Sub chkTipoFactura_Click()
    If chkTipoFactura.Value = Checked Then
        cboNotaCredito1.Enabled = True
    Else
        cboNotaCredito1.Enabled = False
    End If
End Sub

Private Sub chkTipoFactura_LostFocus()
    If chkTipoFactura.Value = Checked And chkCliente.Value = Unchecked _
        And chkSucursal.Value = Unchecked And chkVendedor.Value = Unchecked _
        And chkFecha.Value = Unchecked Then cboNotaCredito1.SetFocus
End Sub

Private Sub chkVendedor_Click()
    If chkVendedor.Value = Checked Then
        txtVendedor.Enabled = True
    Else
        txtVendedor.Enabled = False
    End If
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Select Case TipoBusquedaDoc
    
    Case 1 'BUSCA NOTA DE CREDITO
        
         sql = "SELECT NC.*, C.CLI_RAZSOC , S.SUC_DESCRI, V.VEN_NOMBRE, TC.TCO_DESCRI"
         sql = sql & " FROM NOTA_CREDITO_CLIENTE NC, FACTURA_CLIENTE FC, REMITO_CLIENTE RC ,NOTA_PEDIDO NP"
         sql = sql & ",TIPO_COMPROBANTE TC , CLIENTE C, SUCURSAL S, VENDEDOR V"
         sql = sql & " WHERE NC.FCL_TCO_CODIGO = FC.TCO_CODIGO"
         sql = sql & " AND  NC.FCL_FECHA = FC.FCL_FECHA"
         sql = sql & " AND  NC.FCL_NUMERO = FC.FCL_NUMERO"
         sql = sql & " AND FC.RCL_NUMERO=RC.RCL_NUMERO"
         sql = sql & " AND FC.RCL_FECHA=RC.RCL_FECHA"
         sql = sql & " AND RC.NPE_NUMERO=NP.NPE_NUMERO"
         sql = sql & " AND RC.NPE_FECHA=NP.NPE_FECHA"
         sql = sql & " AND NC.TCO_CODIGO=TC.TCO_CODIGO"
         sql = sql & " AND NP.CLI_CODIGO=C.CLI_CODIGO"
         sql = sql & " AND NP.SUC_CODIGO=S.SUC_CODIGO"
         sql = sql & " AND C.CLI_CODIGO=S.CLI_CODIGO"
         sql = sql & " AND NP.VEN_CODIGO=V.VEN_CODIGO"
     
        If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente)
        If txtSucursal.Text <> "" Then sql = sql & "AND NP.SUC_CODIGO=" & XN(txtSucursal)
        If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
        If FechaDesde <> "" Then sql = sql & " AND NC.NCC_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND NC.NCC_FECHA<=" & XDQ(FechaHasta)
        If chkTipoFactura.Value = Checked Then sql = sql & " AND NC.TCO_CODIGO=" & cboNotaCredito1.ItemData(cboNotaCredito1.ListIndex)
        sql = sql & " ORDER BY NC.NCC_NUMERO"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem rec!TCO_DESCRI & Chr(9) & rec!NCC_NUMERO & Chr(9) & rec!NCC_FECHA _
                                & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!SUC_DESCRI & Chr(9) & rec!VEN_NOMBRE _
                                & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!FCL_NUMERO & Chr(9) & rec!FCL_FECHA _
                                & Chr(9) & rec!NCC_BONIFICA & Chr(9) & rec!NCC_IVA & Chr(9) & rec!NCC_OBSERVACION _
                                & Chr(9) & rec!TCO_CODIGO & Chr(9) & rec!CNC_CODIGO & Chr(9) & rec!NCC_BONIPESOS _
                                & Chr(9) & rec!FCL_TCO_CODIGO
                rec.MoveNext
            Loop
            GrdModulos.SetFocus
        Else
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        End If
        
    Case 2 'BUSCA FACTURA
        
        sql = "SELECT FC.FCL_NUMERO, FC.FCL_FECHA,FC.TCO_CODIGO,"
        sql = sql & " C.CLI_RAZSOC, S.SUC_DESCRI, V.VEN_NOMBRE, TC.TCO_DESCRI"
        sql = sql & " FROM FACTURA_CLIENTE FC, REMITO_CLIENTE RC, CLIENTE C, SUCURSAL S,"
        sql = sql & " NOTA_PEDIDO NP, VENDEDOR V, TIPO_COMPROBANTE TC"
        sql = sql & " WHERE"
        sql = sql & " FC.RCL_NUMERO=RC.RCL_NUMERO"
        sql = sql & " AND FC.RCL_FECHA=RC.RCL_FECHA"
        sql = sql & " AND FC.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND RC.NPE_NUMERO=NP.NPE_NUMERO"
        sql = sql & " AND RC.NPE_FECHA=NP.NPE_FECHA"
        sql = sql & " AND NP.CLI_CODIGO=C.CLI_CODIGO"
        sql = sql & " AND NP.SUC_CODIGO=S.SUC_CODIGO"
        sql = sql & " AND C.CLI_CODIGO=S.CLI_CODIGO"
        sql = sql & " AND NP.VEN_CODIGO=V.VEN_CODIGO"
        If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente)
        If txtSucursal.Text <> "" Then sql = sql & "AND NP.SUC_CODIGO=" & XN(txtSucursal)
        If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
        If FechaDesde <> "" Then sql = sql & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta)
        If chkTipoFactura.Value = Checked Then sql = sql & " AND FC.TCO_CODIGO=" & cboNotaCredito1.ItemData(cboNotaCredito1.ListIndex)
        sql = sql & " ORDER BY FC.FCL_NUMERO"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem rec!TCO_DESCRI & Chr(9) & rec!FCL_NUMERO & Chr(9) & rec!FCL_FECHA _
                                & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!SUC_DESCRI _
                                & Chr(9) & rec!VEN_NOMBRE & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                                & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & rec!TCO_CODIGO
                rec.MoveNext
            Loop
            GrdModulos.SetFocus
        Else
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        End If
    End Select
    
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

Private Sub cmdBuscarFactura_Click()
    TipoBusquedaDoc = 2 'BUSCA FACTURAS
    tabDatos.Tab = 1
End Sub

Private Sub cmdBuscarSuc_Click()
    frmBuscar.TipoBusqueda = 3
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 3
        txtCliente.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 0
        txtSucursal.Text = frmBuscar.grdBuscar.Text
        txtSucursal.SetFocus
        txtSucursal_LostFocus
    Else
        txtSucursal.SetFocus
    End If
End Sub

Private Sub cmdGrabar_Click()
    Dim VStockFisico As String
    
    If ValidarNotaCredito = False Then Exit Sub
    
    On Error GoTo HayErrorFactura
    
    DBConn.BeginTrans
    sql = "SELECT * FROM NOTA_CREDITO_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
    sql = sql & " AND FCL_NUMERO = " & XN(txtNroNotaCredito)
    sql = sql & " AND FCL_FECHA=" & XDQ(FechaNotaCredito)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = False Then
        If MsgBox("Seguro que modificar la Nota de Crédito Nro.: " & Trim(txtNroNotaCredito), vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
            sql = "UPDATE NOTA_CREDITO_CLIENTE"
            sql = sql & " SET NCC_IVA=" & XN(txtPorcentajeIva)
            sql = sql & " ,CNC_CODIGO=" & cboConcepto.ItemData(cboConcepto.ListIndex)
            sql = sql & " ,NCC_OBSERVACION=" & XS(txtObservaciones)
            sql = sql & " ,NCC_BONIFICA=" & XN(txtPorcentajeBoni)
            If chkBonificaEnPesos.Value = Checked Then
                sql = sql & " ,NCC_BONIPESOS='S'" 'BONIFICA EN PESOS
            ElseIf chkBonificaEnPorsentaje.Value = Checked Then
                sql = sql & " ,NCC_BONIPESOS='N'" 'BONIFICA EN PORCENTAJE
            Else
                sql = sql & " ,NCC_BONIPESOS=NULL" 'NO BONIFICA
            End If
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
            sql = sql & " AND NCC_NUMERO=" & XN(txtNroFactura)
            sql = sql & " AND NCC_FECHA=" & XDQ(FechaFactura)
            DBConn.Execute sql
            
            sql = "DELETE FROM DETALLE_NOTA_CREDITO_CLIENTE"
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
            sql = sql & " AND NCC_NUMERO=" & XN(txtNroFactura)
            sql = sql & " AND NCC_FECHA=" & XDQ(FechaFactura)
            DBConn.Execute sql
            
            For I = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(I, 0) <> "" Then
                    sql = "INSERT INTO DETALLE_NOTA_CREDITO_CLIENTE"
                    sql = sql & " (TCO_CODIGO,NCC_NUMERO,NCC_FECHA,DNC_NROITEM,PTO_CODIGO"
                    sql = sql & ",DNC_CANTIDAD,DNC_PRECIO,DNC_BONIFICA)"
                    sql = sql & " VALUES ("
                    sql = sql & cboNotaCredito.ItemData(cboNotaCredito.ListIndex) & ","
                    sql = sql & XN(txtNroNotaCredito) & ","
                    sql = sql & XDQ(FechaNotaCredito) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 9)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 0)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 4)) & ")"
                    DBConn.Execute sql
                End If
            Next
            DBConn.CommitTrans
        End If
    Else
        sql = "INSERT INTO NOTA_CREDITO_CLIENTE"
        sql = sql & " (TCO_CODIGO, NCC_NUMERO, NCC_FECHA, FCL_NUMERO,"
        sql = sql & " FCL_FECHA, FCL_TCO_CODIGO, NCC_BONIFICA, NCC_IVA, CNC_CODIGO,"
        sql = sql & " NCC_OBSERVACION, NCC_BONIPESOS, EST_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & cboNotaCredito.ItemData(cboNotaCredito.ListIndex) & ","
        sql = sql & XN(txtNroNotaCredito) & ","
        sql = sql & XDQ(FechaNotaCredito) & ","
        sql = sql & XN(txtNroFactura) & ","
        sql = sql & XDQ(FechaFactura) & ","
        sql = sql & cboFactura.ItemData(cboFactura.ListIndex) & ","
        sql = sql & XN(txtPorcentajeBoni) & ","
        sql = sql & XN(txtPorcentajeIva) & ","
        sql = sql & cboConcepto.ItemData(cboConcepto.ListIndex) & ","
        sql = sql & XS(txtObservaciones) & ","
        If chkBonificaEnPesos.Value = Checked Then
            sql = sql & "'S'" & "," 'BONIFICA EN PESOS
        ElseIf chkBonificaEnPorsentaje.Value = Checked Then
            sql = sql & "'N'" & "," 'BONIFICA EN PORCENTAJE
        Else
            sql = sql & "NULL" & "," 'NO HAY BONIFICACION
        End If
        sql = sql & "3)" 'ESTADO DEFINITIVO
        DBConn.Execute sql
           
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                sql = "INSERT INTO DETALLE_NOTA_CREDITO_CLIENTE"
                sql = sql & " (TCO_CODIGO,NCC_NUMERO,NCC_FECHA,DNC_NROITEM,PTO_CODIGO"
                sql = sql & ",DNC_CANTIDAD,DNC_PRECIO,DNC_BONIFICA)"
                sql = sql & " VALUES ("
                sql = sql & cboNotaCredito.ItemData(cboNotaCredito.ListIndex) & ","
                sql = sql & XN(txtNroNotaCredito) & ","
                sql = sql & XDQ(FechaNotaCredito) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 9)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 0)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 4)) & ")"
                DBConn.Execute sql
            End If
        Next
        
        'ACTUALIZO EL STOCK (STOCK FISICO)
        If cboConcepto.ItemData(cboConcepto.ListIndex) = 1 Then
             For I = 1 To grdGrilla.Rows - 1
                 If grdGrilla.TextMatrix(I, 0) <> "" Then
                     VStockFisico = ""
                     Set Rec2 = New ADODB.Recordset
                     sql = "SELECT STK_CODIGO,PTO_CODIGO,DST_STKFIS"
                     sql = sql & " FROM DETALLE_STOCK"
                     sql = sql & " WHERE"
                     sql = sql & " STK_CODIGO=" & XN(txtCodigoStock)
                     sql = sql & " AND PTO_CODIGO=" & XN(grdGrilla.TextMatrix(I, 0))
                     Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
                     If Rec2.EOF = False Then
                         VStockFisico = CStr(CInt(grdGrilla.TextMatrix(I, 2)) + CInt(Rec2!DST_STKFIS))
                         sql = "UPDATE DETALLE_STOCK"
                         sql = sql & " SET"
                         sql = sql & " DST_STKFIS=" & XN(VStockFisico)
                         sql = sql & " WHERE STK_CODIGO=" & XN(txtCodigoStock)
                         sql = sql & " AND PTO_CODIGO=" & XN(grdGrilla.TextMatrix(I, 0))
                         DBConn.Execute sql
                     End If
                     Rec2.Close
                 End If
             Next
        End If
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO A LA NOTA DE CREDITO QUE CORRESPONDA
        Select Case cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
            Case 4
                sql = "UPDATE PARAMETROS SET NOTA_CREDITO_A=" & XN(txtNroNotaCredito)
            Case 5
                sql = "UPDATE PARAMETROS NOTA_CREDITO_B=" & XN(txtNroNotaCredito)
        End Select
                DBConn.Execute sql
        
        DBConn.CommitTrans
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    CmdNuevo_Click
    Exit Sub
    
HayErrorFactura:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description
End Sub

Private Function ValidarNotaCredito() As Boolean
    If FechaNotaCredito.Text = "" Then
        MsgBox "La Fecha de la Nota de Crédito es requerida", vbExclamation, TIT_MSGBOX
        FechaNotaCredito.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If txtNroFactura.Text = "" Then
        MsgBox "El número de la Factura es requerido", vbExclamation, TIT_MSGBOX
        txtNroFactura.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If FechaFactura.Text = "" Then
        MsgBox "La Fecha de la Factura es requerida", vbExclamation, TIT_MSGBOX
        FechaFactura.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If cboConcepto.ListIndex = -1 Then
        MsgBox "Debe ingresar el concepto por el cual se emite la Nota de Crédito", vbExclamation, TIT_MSGBOX
        cboConcepto.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If chkBonificaEnPesos.Value = Checked Or chkBonificaEnPorsentaje.Value = Checked Then
        If txtPorcentajeBoni.Text = "" Then
            MsgBox "Debe ingresar la Bonificación", vbExclamation, TIT_MSGBOX
            txtPorcentajeBoni.SetFocus
            ValidarNotaCredito = False
            Exit Function
        End If
    End If
    ValidarNotaCredito = True
End Function

Private Sub cmdImprimir_Click()
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
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Imprimiendo..."
    
    For W = 1 To 2 'SE IMPRIME POR DUPLICADO
      '-----IMPRESION DEL ENCABEZADO------------------
        ImprimirEncabezado
        
      '---- IMPRESION DE LA NOTA CREDITO ------------------
        Renglon = 8.5
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                Imprimir 0, Renglon, False, Format(grdGrilla.TextMatrix(I, 0), "000000")  'codigo
                If Len(grdGrilla.TextMatrix(I, 1)) < 25 Then
                    Imprimir 2.5, Renglon, False, grdGrilla.TextMatrix(I, 1) 'descripcion
                Else
                    Imprimir 2.5, Renglon, False, Left(grdGrilla.TextMatrix(I, 1), 24) & "..." 'descripcion
                End If
                Imprimir 10.6, Renglon, False, grdGrilla.TextMatrix(I, 2) 'cantidad
                Imprimir 13, Renglon, False, grdGrilla.TextMatrix(I, 3) 'precio
                Imprimir 15, Renglon, False, IIf(grdGrilla.TextMatrix(I, 4) = "", "0,00", grdGrilla.TextMatrix(I, 4)) 'bonoficacion
                Imprimir 17, Renglon, False, grdGrilla.TextMatrix(I, 6) 'importe
                Renglon = Renglon + 0.4
            End If
        Next I
            Imprimir 0, 16.5, True, "texto de bajo del detalle"
            '-------------IMPRIMO TOTALES--------------------
            Imprimir 17.2, 16.5, True, txtSubtotal.Text
            Imprimir 0.3, 18.6, True, txtSubtotal.Text
            If txtPorcentajeBoni.Text <> "" Then
                If chkBonificaEnPesos.Value = Checked Then
                    Imprimir 3.2, 18.4, True, "$" & txtPorcentajeBoni.Text
                    Imprimir 4.4, 18.8, True, txtImporteBoni.Text
                Else
                    Imprimir 3.2, 18.4, True, "%" & txtPorcentajeBoni.Text
                    Imprimir 4.4, 18.8, True, txtImporteBoni.Text
                End If
                Imprimir 7, 18.6, True, txtSubTotalBoni.Text
            End If
            Imprimir 10.2, 18.4, True, "%" & txtPorcentajeIva.Text
            Imprimir 11.2, 18.8, True, txtImporteIva.Text
            Imprimir 17, 18.6, True, txtTotal.Text
        Printer.EndDoc
    Next W
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Public Sub ImprimirEncabezado()
 '-----------IMPRIME EL ENCABEZADO DE LA FACTURA-------------------
    Imprimir 15.8, 2.5, False, Format(FechaNotaCredito, "dd/mm/yyyy")
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CUIT,C.CLI_INGBRU, L.LOC_DESCRI"
    sql = sql & ", P.PRO_DESCRI,CI.IVA_DESCRI"
    sql = sql & " FROM CLIENTE C, LOCALIDAD L, NOTA_PEDIDO NP, REMITO_CLIENTE RC,"
    sql = sql & " PROVINCIA P, CONDICION_IVA CI, FACTURA_CLIENTE FC"
    sql = sql & " WHERE"
    sql = sql & " FC.FCL_NUMERO=" & XN(txtNroFactura)
    sql = sql & " AND FC.FCL_FECHA=" & XDQ(FechaFactura)
    sql = sql & " AND FC.RCL_NUMERO=RC.RCL_NUMERO"
    sql = sql & " AND FC.RCL_FECHA=RC.RCL_FECHA"
    sql = sql & " AND RC.NPE_NUMERO=NP.NPE_NUMERO"
    sql = sql & " AND RC.NPE_FECHA=NP.NPE_FECHA"
    sql = sql & " AND NP.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Imprimir 1, 4.5, True, Trim(Rec1!CLI_RAZSOC)
        Imprimir 1, 4.9, False, Trim(Rec1!CLI_DOMICI)
        'FACTURA
        Imprimir 13.5, 6.9, True, Format(txtNroFactura.Text, "00000000") & " del " & Format(FechaFactura.Text, "dd/mm/yyyy")
        Imprimir 1, 5.3, False, "Loc: " & Trim(Rec1!LOC_DESCRI) & " -- Prov: " & Trim(Rec1!PRO_DESCRI)
        Imprimir 1, 6, False, Trim(Rec1!IVA_DESCRI)
        Imprimir 7, 6, False, IIf(IsNull(Rec1!CLI_CUIT), "NO INFORMADO", PongoGuionCuit(Trim(Rec1!CLI_CUIT)))
        Imprimir 13, 6, False, IIf(IsNull(Rec1!CLI_INGBRU), "NO INFORMADO", PongoGuionIngBrutos(Trim(Rec1!CLI_INGBRU)))
    End If
    Rec1.Close
    Imprimir 1.5, 6.9, False, cboConcepto.Text
    Imprimir 0, 7.7, False, "Código"
    Imprimir 2.5, 7.7, False, "Descripción"
    Imprimir 10, 7.7, False, "Cantidad"
    Imprimir 13, 7.7, False, "Precio"
    Imprimir 15, 7.7, False, "Bonof."
    Imprimir 17, 7.7, False, "Importe"
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
        grdGrilla.TextMatrix(I, 7) = ""
        grdGrilla.TextMatrix(I, 9) = I
   Next
   grillaFactura.TextMatrix(0, 1) = ""
   grillaFactura.TextMatrix(1, 1) = ""
   grillaFactura.TextMatrix(2, 1) = ""
   FechaFactura.Text = ""
   txtNroFactura.Text = ""
   FechaFactura.Text = ""
   txtNroNotaCredito.Text = ""
   FechaNotaCredito.Text = Date
   lblEstadoNotaCredito.Caption = ""
   txtSubtotal.Text = ""
   txtTotal.Text = ""
   txtCodigoStock.Text = ""
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
    lblEstadoNotaCredito.Caption = BuscoEstado(1) 'ESTADO PENDIENTE
    VEstadoNotaCredito = 1
    '--------------
    chkBonificaEnPorsentaje.Value = Unchecked
    chkBonificaEnPesos.Value = Unchecked
    FrameNotaCredito.Enabled = True
    FrameFactura.Enabled = True
    tabDatos.Tab = 0
    cboNotaCredito.ListIndex = 0
    cboNotaCredito.SetFocus
End Sub

Private Sub cmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmNotaDeditoCliente = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        TipoBusquedaDoc = 1 'BUSCA NOTA DE CREDITO
        GrdModulos.ColWidth(0) = 1300 'TIPO NOTA CREDITO O TIPO FACTURA
        chkTipoFactura.Visible = True
        lbltipoFac.Visible = True
        tabDatos.Tab = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl.Name <> "grdGrilla" And _
        Me.ActiveControl.Name <> "txtEdit" And _
        KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Function BuscoEstado(Codigo As Integer) As String
    sql = "SELECT EST_DESCRI"
    sql = sql & " FROM ESTADO_DOCUMENTO"
    sql = sql & " WHERE"
    sql = sql & " EST_CODIGO=" & Codigo
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscoEstado = rec!EST_DESCRI
    End If
    rec.Close
End Function

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    grdGrilla.FormatString = "Código|Descripción|Cantidad|Precio|Bonif.|Pre.Bonif.|Importe|Rubro|Linea|Orden"
    grdGrilla.ColWidth(0) = 800  'CODIGO
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
    GrdModulos.FormatString = "Tipo Fac|>Número|^Fecha|Cliente|Sucursal|Vendedor|Cod_Estado|" _
                              & "REMITO_NUMERO|REMITO_FECHA|PORCENTAJE BONIFICA|PORCENTAJE IVA|" _
                              & "OBSERVACIONES|COD TIPO COMPROBANTE NOTA CREDITO|" _
                              & "COD CONDICION VENTA|BONIFICA EN PESOS|COD TIPO COMPROBANTE FACTURA"
    GrdModulos.ColWidth(0) = 1100 'TIPO NOTA CREDITO O TIPO FACTURA
    GrdModulos.ColWidth(1) = 1100 'NUMERO
    GrdModulos.ColWidth(2) = 1000 'FECHA
    GrdModulos.ColWidth(3) = 4000 'CLIENTE
    GrdModulos.ColWidth(4) = 4000 'SUCURSAL
    GrdModulos.ColWidth(5) = 0    'VENDEDOR
    GrdModulos.ColWidth(6) = 0    'COD_ESTADO
    GrdModulos.ColWidth(7) = 0    'FACTURA_NUMERO
    GrdModulos.ColWidth(8) = 0    'FACTURA_FECHA
    GrdModulos.ColWidth(9) = 0    'PORCENTAJE BONIFICA
    GrdModulos.ColWidth(10) = 0   'PORCENTAJE IVA
    GrdModulos.ColWidth(11) = 0   'OBSERVACIONES
    GrdModulos.ColWidth(12) = 0   'COD TIPO COMPROBANTE NOTA CREDITO
    GrdModulos.ColWidth(13) = 0   'COD CONCEPTO
    GrdModulos.ColWidth(14) = 0   'BONIFICA EN PESOS
    GrdModulos.ColWidth(15) = 0   'COD TIPO COMPROBANTE FACTURA
    GrdModulos.Rows = 1
    '------------------------------------
    grillaFactura.ColWidth(0) = 950
    grillaFactura.ColWidth(1) = 5300
    grillaFactura.TextMatrix(0, 0) = "    Cliente:"
    grillaFactura.TextMatrix(1, 0) = " Sucursal:"
    grillaFactura.TextMatrix(2, 0) = "Vendedor:"
    '------------------------------------
    lblEstado.Caption = ""
    'CARGO COMBO CON LOS TIPOS DE NOTA DE CREDITO
    LlenarComboNotaCredito
    'CARGO COMBO CON LOS TIPOS DE FACTURA
    LlenarComboFactura
    'CARGO COMBO CON LOS CONCEPTOS DE NOTA DE CREDITO
    LlenarComboConcepto
    'CARGO ESTADO
    lblEstadoNotaCredito.Caption = BuscoEstado(1) 'ESTADO PENDIENTE
    VEstadoNotaCredito = 1
    FechaNotaCredito.Text = Date
    TipoBusquedaDoc = 1 'ESTO ES PARA BUSCAR FACTURA(1), (2)PARA BUSCAR REMITOS
    tabDatos.Tab = 0
    'BUSCO IVA
    BuscoIva
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
            rec.MoveNext
        Loop
        cboNotaCredito.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboFactura()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'FAC%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboFactura.AddItem rec!TCO_DESCRI
            cboFactura.ItemData(cboFactura.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboFactura.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboNCyFAC()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    If TipoBusquedaDoc = 1 Then
        sql = sql & " WHERE TCO_DESCRI LIKE 'NOTA DE CRE%'"
    ElseIf TipoBusquedaDoc = 2 Then
        sql = sql & " WHERE TCO_DESCRI LIKE 'FAC%'"
    End If
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    cboNotaCredito1.Clear
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboNotaCredito1.AddItem rec!TCO_DESCRI
            cboNotaCredito1.ItemData(cboNotaCredito1.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboNotaCredito1.ListIndex = 0
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
                MsgBox "No hay Notas de Crédito del tipo C", vbExclamation, TIT_MSGBOX
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
            EDITAR grdGrilla, txtEdit, KeyAscii
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
        Select Case TipoBusquedaDoc
    
        Case 1 'BUSCA NOTA CREDITO
        
            Set Rec1 = New ADODB.Recordset
            lblEstado.Caption = "Buscando..."
            Screen.MousePointer = vbHourglass
            'CABEZA NOTA CREDITO
            Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 12)), cboNotaCredito)
            txtNroNotaCredito.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
            FechaNotaCredito.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
            lblEstadoNotaCredito.Caption = BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6)))
            VEstadoNotaCredito = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 6))
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 11) <> "" Then
                txtObservaciones.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 11))
            End If
            'CABEZA FACTURA
            Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 15)), cboFactura)
            txtNroFactura.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 7)
            FechaFactura.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
            grillaFactura.TextMatrix(0, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 3)
            grillaFactura.TextMatrix(1, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 4)
            grillaFactura.TextMatrix(2, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
            'CONDICION NOTA CREDITO
            Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 13)), cboConcepto)
            '----BUSCO DETALLE DE LA NOTA DE CREDITO------------------
            sql = "SELECT DNC.*, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM DETALLE_NOTA_CREDITO_CLIENTE DNC, PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE DNC.NCC_NUMERO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 1))
            sql = sql & " AND DNC.NCC_FECHA=" & XDQ(GrdModulos.TextMatrix(GrdModulos.RowSel, 2))
            sql = sql & " AND DNC.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 12))
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
                    grdGrilla.TextMatrix(I, 1) = Rec1!PTO_DESCRI
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
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 14) = "S" Then
                chkBonificaEnPesos.Value = Checked
            ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 14) = "N" Then
                chkBonificaEnPorsentaje.Value = Checked
            Else
                chkBonificaEnPesos.Value = Unchecked
                chkBonificaEnPorsentaje.Value = Unchecked
            End If
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 9) <> "" Then
                txtPorcentajeBoni.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 9)
                txtPorcentajeBoni_LostFocus
            End If
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 10) <> "" Then
                txtPorcentajeIva = GrdModulos.TextMatrix(GrdModulos.RowSel, 10)
                txtPorcentajeIva_LostFocus
            End If
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            '--------------
            FrameNotaCredito.Enabled = False
            FrameFactura.Enabled = False
            '--------------
            tabDatos.Tab = 0
            cboConcepto.SetFocus
        '----------------------------------------------------------
        Case 2 'BUSCA FACTURA
        
            lblEstado.Caption = "Buscando..."
            Screen.MousePointer = vbHourglass
            'BUSCA TIPO FACTURA
            Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 12)), cboFactura)
            txtNroFactura.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
            FechaFactura.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
            
            grillaFactura.TextMatrix(0, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 3)
            grillaFactura.TextMatrix(1, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 4)
            grillaFactura.TextMatrix(2, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
        
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            tabDatos.Tab = 0
            txtNroFactura_LostFocus
            cboConcepto.SetFocus
        End Select
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub


Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    GrdModulos.Rows = 2
    txtCliente.Enabled = False
    txtSucursal.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    txtVendedor.Enabled = False
    cboNotaCredito1.Enabled = False
    cmdBuscarCli.Enabled = False
    cmdBuscarSuc.Enabled = False
    cmdGrabar.Enabled = False
    LimpiarBusqueda
    LlenarComboNCyFAC
    If Me.Visible = True Then chkCliente.SetFocus
    If TipoBusquedaDoc = 1 Then
        frameBuscar.Caption = "Buscar Nota de Crédito por..."
    Else
        frameBuscar.Caption = "Buscar Factura por..."
    End If
    
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
    txtSucursal.Text = ""
    txtDesSuc.Text = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    txtVendedor.Text = ""
    txtDesVen.Text = ""
    GrdModulos.Rows = 1
    chkCliente.Value = Unchecked
    chkSucursal.Value = Unchecked
    chkFecha.Value = Unchecked
    chkVendedor.Value = Unchecked
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
    If chkSucursal.Value = Unchecked And chkFecha.Value = Unchecked _
        And chkVendedor.Value = Unchecked And chkTipoFactura.Value = Unchecked _
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
    If grdGrilla.Col = 0 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 2 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 3 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
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
                sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, D.LIS_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI"
                sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, DETALLE_LISTA_PRECIO D"
                sql = sql & " WHERE"
                If grdGrilla.Col = 0 Then
                    sql = sql & " P.PTO_CODIGO=" & XN(txtEdit)
                Else
                    sql = sql & " P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
                End If
                    sql = sql & " AND D.LIS_CODIGO=" & XN(txtCodigoStock)
                    sql = sql & " AND P.PTO_CODIGO=D.PTO_CODIGO"
                    sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
                    sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
                    sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                    
                rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If rec.EOF = False Then
                    If rec.RecordCount > 1 Then
                        grdGrilla.SetFocus
                        frmBuscar.TipoBusqueda = 2
                        'LE DIGO EN QUE LISTA DE PRECIO BUSCAR LOS PRECIOS
                        frmBuscar.CodListaPrecio = CInt(txtCodigoStock)
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
                        grdGrilla.Col = 1
                        grdGrilla.Text = Trim(rec!PTO_DESCRI)
                        grdGrilla.Col = 3
                        grdGrilla.Text = Valido_Importe(Trim(rec!LIS_PRECIO))
                        grdGrilla.Col = 7
                        grdGrilla.Text = Trim(rec!RUB_DESCRI)
                        grdGrilla.Col = 8
                        grdGrilla.Text = Trim(rec!LNA_DESCRI)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = grdGrilla.RowSel
                        grdGrilla.Col = 2
                    End If
                        If BuscoRepetetidos(CLng(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), grdGrilla.RowSel) = False Then
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
                    VBonificacion = (CInt(txtEdit.Text) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)))
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Valido_Importe(CStr(VBonificacion))
                    If grdGrilla.TextMatrix(grdGrilla.RowSel, 4) <> "" Then
                        VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 4))) / 100)
                        VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6)) - VBonificacion)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Valido_Importe(CStr(VBonificacion))
                    End If
                    txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                    txtTotal.Text = txtSubtotal.Text
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

Private Function BuscoRepetetidos(Codigo As Long, Linea As Integer) As Boolean
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            If Codigo = CLng(grdGrilla.TextMatrix(I, 0)) And (I <> Linea) Then
                MsgBox "El producto ya fue elegido anteriormente", vbExclamation, TIT_MSGBOX
                BuscoRepetetidos = False
                Exit Function
            End If
        End If
    Next
    BuscoRepetetidos = True
End Function

Private Sub LimpiarFactura()
    txtNroFactura.Text = ""
    FechaFactura.Text = ""
    txtCodigoStock.Text = ""
    grillaFactura.TextMatrix(0, 1) = ""
    grillaFactura.TextMatrix(1, 1) = ""
    grillaFactura.TextMatrix(2, 1) = ""
    txtNroFactura.SetFocus
End Sub

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

Private Function BuscoCliente(Codigo As String) As String
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(Codigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            BuscoCliente = rec!CLI_RAZSOC
        Else
            BuscoCliente = "No se encontro el Cliente"
        End If
        rec.Close
End Function

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

Private Sub txtNroFactura_GotFocus()
    SelecTexto txtNroFactura
End Sub

Private Sub txtNroFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroFactura_LostFocus()
    If txtNroFactura.Text <> "" Then
    
        Set Rec2 = New ADODB.Recordset
        sql = "SELECT FC.FCL_NUMERO, FC.FCL_FECHA, FC.FCL_BONIFICA, FC.FCL_IVA, FC.FCL_BONIPESOS,"
        sql = sql & "RC.EST_CODIGO, RC.STK_CODIGO, E.EST_DESCRI"
        sql = sql & " ,NP.CLI_CODIGO, NP.SUC_CODIGO, NP.VEN_CODIGO"
        sql = sql & " FROM FACTURA_CLIENTE FC, REMITO_CLIENTE RC, NOTA_PEDIDO NP, ESTADO_DOCUMENTO E"
        sql = sql & " WHERE FC.FCL_NUMERO=" & XN(txtNroFactura)
        If FechaFactura.Text <> "" Then
            sql = sql & " AND FC.FCL_FECHA=" & XDQ(FechaFactura)
        End If
        sql = sql & " AND FC.TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
        sql = sql & " AND FC.RCL_NUMERO=RC.RCL_NUMERO"
        sql = sql & " AND FC.RCL_FECHA=RC.RCL_FECHA"
        sql = sql & " AND RC.NPE_NUMERO=NP.NPE_NUMERO"
        sql = sql & " AND RC.NPE_FECHA=NP.NPE_FECHA"
        sql = sql & " AND FC.EST_CODIGO=E.EST_CODIGO"

        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic

        If Rec2.EOF = False Then
            If Rec2.RecordCount > 1 Then
                MsgBox "Hay mas de una Factura con el Número: " & txtNroFactura.Text, vbInformation, TIT_MSGBOX
                Rec2.Close
                cmdBuscarFactura_Click
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Buscando..."

            'CARGO CABECERA DE LA FACTURA
            FechaFactura.Text = Rec2!FCL_FECHA
            grillaFactura.TextMatrix(0, 1) = BuscoCliente(Rec2!CLI_CODIGO)
            grillaFactura.TextMatrix(1, 1) = BuscoSucursal(Rec2!SUC_CODIGO, Rec2!CLI_CODIGO)
            grillaFactura.TextMatrix(2, 1) = BuscoVendedor(Rec2!VEN_CODIGO)
            txtCodigoStock.Text = Rec2!STK_CODIGO

            If Rec2!EST_CODIGO = 2 Then
                MsgBox "La Factura número: " & txtNroFactura.Text & Chr(13) & Chr(13) & _
                       "No puede ser asignado a la Nota de Crédito por su estado (" & Rec2!EST_DESCRI & ")", vbExclamation, TIT_MSGBOX
                cmdGrabar.Enabled = False
                Screen.MousePointer = vbNormal
                lblEstado.Caption = ""
                Rec2.Close
                LimpiarFactura
                Exit Sub
            Else
                cmdGrabar.Enabled = True
            End If
            
            '----BUSCO DETALLE DE LA FACTURA------------------
            sql = "SELECT DFC.*, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM DETALLE_FACTURA_CLIENTE DFC, PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE DFC.FCL_NUMERO=" & XN(txtNroFactura)
            sql = sql & " AND DFC.FCL_FECHA=" & XDQ(FechaFactura)
            sql = sql & " AND DFC.TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
            sql = sql & " AND DFC.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " ORDER BY DFC.DFC_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(I, 1) = Rec1!PTO_DESCRI
                    grdGrilla.TextMatrix(I, 2) = Rec1!DFC_CANTIDAD
                    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DFC_PRECIO)
                    If IsNull(Rec1!DFC_BONIFICA) Then
                        grdGrilla.TextMatrix(I, 4) = ""
                    Else
                        grdGrilla.TextMatrix(I, 4) = Valido_Importe(Rec1!DFC_BONIFICA)
                    End If
                    VBonificacion = 0
                    If Not IsNull(Rec1!DFC_BONIFICA) Then
                        VBonificacion = (((CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO)) * CDbl(Rec1!DFC_BONIFICA)) / 100)
                        VBonificacion = ((CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO)) - VBonificacion)
                        grdGrilla.TextMatrix(I, 5) = Valido_Importe(CStr(VBonificacion))
                        grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
                    Else
                        VBonificacion = (CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO))
                        grdGrilla.TextMatrix(I, 5) = ""
                        grdGrilla.TextMatrix(I, 6) = Valido_Importe(CStr(VBonificacion))
                    End If
                    grdGrilla.TextMatrix(I, 7) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 8) = Rec1!LNA_DESCRI
                    grdGrilla.TextMatrix(I, 9) = Rec1!DFC_NROITEM
                    I = I + 1
                    Rec1.MoveNext
                Loop
                VBonificacion = 0
            End If
            Rec1.Close
            '--CARGO LOS TOTALES----
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            txtTotal.Text = txtSubtotal.Text
            If Rec2!FCL_BONIPESOS = "S" Then
                chkBonificaEnPesos.Value = Checked
            ElseIf Rec2!FCL_BONIPESOS = "N" Then
                chkBonificaEnPorsentaje.Value = Checked
            Else
                chkBonificaEnPesos.Value = Unchecked
                chkBonificaEnPorsentaje.Value = Unchecked
            End If
            If Not IsNull(Rec2!FCL_BONIFICA) Then
                txtPorcentajeBoni.Text = Rec2!FCL_BONIFICA
                txtPorcentajeBoni_LostFocus
            End If
            If Not IsNull(Rec2!FCL_IVA) Then
                txtPorcentajeIva = Rec2!FCL_IVA
                txtPorcentajeIva_LostFocus
            End If
            Rec2.Close
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            '--------------
            FrameNotaCredito.Enabled = False
            FrameFactura.Enabled = False
            '--------------
        Else
            MsgBox "La Factura no existe", vbExclamation, TIT_MSGBOX
            If Rec2.State = 1 Then Rec2.Close
            LimpiarFactura
        End If
    End If
End Sub

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
    If txtPorcentajeBoni.Text <> "" Then
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
            txtImporteIva.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
            txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
            txtTotal.Text = CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
        End If
    End If
End Sub

Private Sub txtSucursal_Change()
    If txtSucursal.Text = "" Then
        txtDesSuc.Text = ""
    End If
End Sub

Private Sub txtSucursal_GotFocus()
    SelecTexto txtSucursal
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtSucursal_LostFocus()
    
    If txtSucursal.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, SUC_DESCRI FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(txtSucursal)
        If txtCliente.Text <> "" Then
         sql = sql & " AND CLI_CODIGO=" & XN(txtCliente)
        End If
        lblEstado.Caption = "Buscando..."
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtCliente.Text = Rec1!CLI_CODIGO
            txtCliente_LostFocus
            txtDesSuc.Text = Rec1!SUC_DESCRI
            lblEstado.Caption = ""
        Else
            lblEstado.Caption = ""
            MsgBox "La Sucursal no existe", vbExclamation, TIT_MSGBOX
            txtDesSuc.Text = ""
            txtSucursal.SetFocus
        End If
        Rec1.Close
    End If
    If chkFecha.Value = Unchecked And chkVendedor.Value = Unchecked _
    And chkTipoFactura.Value = Unchecked And ActiveControl.Name <> "cmdBuscarSuc" _
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
            txtVendedor.SetFocus
        End If
        rec.Close
    End If
    If chkFecha.Value = Unchecked And chkTipoFactura.Value = Unchecked _
    And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub
