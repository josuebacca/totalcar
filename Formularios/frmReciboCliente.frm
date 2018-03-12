VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReciboCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibo de Cliente"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   8280
      TabIndex        =   143
      Top             =   7035
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   10020
      TabIndex        =   50
      Top             =   7035
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   9150
      TabIndex        =   49
      Top             =   7035
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10890
      TabIndex        =   51
      Top             =   7035
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6945
      Left            =   30
      TabIndex        =   52
      Top             =   15
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   12250
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
      TabPicture(0)   =   "frmReciboCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label16"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameRemito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tabValores"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tabComprobantes"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FrameRecibo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtObservaciones"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmReciboCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtObservaciones 
         BackColor       =   &H00C0FFFF&
         Height          =   465
         Left            =   1110
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   6360
         Width           =   10410
      End
      Begin VB.Frame FrameRecibo 
         Caption         =   "Recibo..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   4455
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
            Left            =   870
            MaxLength       =   4
            TabIndex        =   1
            Top             =   825
            Width           =   555
         End
         Begin VB.TextBox txtNroRecibo 
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
            Left            =   1455
            MaxLength       =   8
            TabIndex        =   2
            Top             =   825
            Width           =   1065
         End
         Begin VB.ComboBox cboRecibo 
            Height          =   315
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   2400
         End
         Begin MSComCtl2.DTPicker FechaRecibo 
            Height          =   315
            Left            =   870
            TabIndex        =   3
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   58392577
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaRendicion 
            Height          =   315
            Left            =   3000
            TabIndex        =   4
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58392577
            CurrentDate     =   41098
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   315
            TabIndex        =   141
            Top             =   1335
            Width           =   495
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Rendición:"
            Height          =   195
            Left            =   2190
            TabIndex        =   140
            Top             =   1335
            Width           =   765
         End
         Begin VB.Label lblEstadoRecibo 
            AutoSize        =   -1  'True
            Caption         =   "EST. RECIBO"
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
            Left            =   870
            TabIndex        =   78
            Top             =   1665
            Width           =   1185
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   270
            TabIndex        =   77
            Top             =   1650
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   210
            TabIndex        =   76
            Top             =   870
            Width           =   600
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   450
            TabIndex        =   75
            Top             =   390
            Width           =   360
         End
      End
      Begin TabDlg.SSTab tabComprobantes 
         Height          =   3975
         Left            =   120
         TabIndex        =   95
         Top             =   2370
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   7011
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Aplicar a"
         TabPicture(0)   =   "frmReciboCliente.frx":0038
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame1"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "C&omprobantes Pendientes"
         TabPicture(1)   =   "frmReciboCliente.frx":0054
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame5"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame5 
            Caption         =   "Comprobantes Pendientes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3435
            Left            =   90
            TabIndex        =   121
            Top             =   480
            Width           =   5535
            Begin VB.CommandButton cmdAceptarComprobantes 
               Caption         =   "A&ceptar"
               Height          =   360
               Left            =   4485
               TabIndex        =   11
               Top             =   2985
               Width           =   900
            End
            Begin VB.TextBox txtImporteApagar 
               Alignment       =   1  'Right Justify
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
               Left            =   1395
               TabIndex        =   9
               Top             =   3000
               Width           =   1185
            End
            Begin VB.TextBox txtSaldo 
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
               Left            =   1395
               TabIndex        =   8
               Top             =   2625
               Width           =   1185
            End
            Begin VB.CommandButton cmdAgregarFacturas 
               Caption         =   "A&gregar"
               Height          =   360
               Left            =   3555
               TabIndex        =   10
               Top             =   2985
               Width           =   900
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaAplicar 
               Height          =   2205
               Left            =   105
               TabIndex        =   7
               Top             =   300
               Width           =   5310
               _ExtentX        =   9366
               _ExtentY        =   3889
               _Version        =   393216
               Cols            =   6
               FixedCols       =   0
               BackColorSel    =   8388736
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Importe a pagar:"
               Height          =   195
               Left            =   195
               TabIndex        =   123
               Top             =   3045
               Width           =   1155
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Saldo:"
               Height          =   195
               Left            =   900
               TabIndex        =   122
               Top             =   2670
               Width           =   450
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Aplicar a..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3405
            Left            =   -74940
            TabIndex        =   117
            Top             =   480
            Width           =   5565
            Begin VB.CommandButton cmdAceptarFacturas 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   4515
               TabIndex        =   14
               Top             =   525
               Width           =   945
            End
            Begin VB.CommandButton cmdAgregarFactura 
               Caption         =   "Agregar Com"
               Height          =   360
               Left            =   2400
               TabIndex        =   12
               Top             =   525
               Width           =   1140
            End
            Begin VB.CommandButton cmdQuitarComprobantes 
               Caption         =   "Quitar"
               Height          =   360
               Left            =   3555
               TabIndex        =   13
               Top             =   525
               Width           =   945
            End
            Begin VB.TextBox txtTotalAplicar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
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
               Left            =   1050
               TabIndex        =   118
               Top             =   2880
               Width           =   1170
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaAplicar1 
               Height          =   1860
               Left            =   105
               TabIndex        =   119
               Top             =   915
               Width           =   5415
               _ExtentX        =   9551
               _ExtentY        =   3281
               _Version        =   393216
               Cols            =   5
               FixedCols       =   0
               RowHeightMin    =   250
               BackColorSel    =   8388736
               FocusRect       =   0
               HighLight       =   0
               SelectionMode   =   1
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   570
               TabIndex        =   124
               Top             =   2925
               Width           =   405
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Total Valores Recibidos:"
               Height          =   195
               Left            =   360
               TabIndex        =   120
               Top             =   3420
               Width           =   1725
            End
         End
      End
      Begin TabDlg.SSTab tabValores 
         Height          =   3975
         Left            =   5865
         TabIndex        =   21
         Top             =   2370
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   7011
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         Tab             =   1
         TabsPerRow      =   5
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&Valores"
         TabPicture(0)   =   "frmReciboCliente.frx":0070
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Cheques"
         TabPicture(1)   =   "frmReciboCliente.frx":008C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&Moneda"
         TabPicture(2)   =   "frmReciboCliente.frx":00A8
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame4"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "&Comprobantes"
         TabPicture(3)   =   "frmReciboCliente.frx":00C4
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame7"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Fa&cturas"
         TabPicture(4)   =   "frmReciboCliente.frx":00E0
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame6"
         Tab(4).ControlCount=   1
         Begin VB.Frame Frame6 
            Caption         =   "Facturas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3435
            Left            =   -74940
            TabIndex        =   135
            Top             =   480
            Width           =   5535
            Begin VB.CommandButton cmdAgregarACta 
               Caption         =   "A&gregar"
               Height          =   360
               Left            =   3585
               TabIndex        =   47
               Top             =   2970
               Width           =   900
            End
            Begin VB.TextBox txtSaldoACta 
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
               Left            =   1410
               TabIndex        =   45
               Top             =   2610
               Width           =   1185
            End
            Begin VB.TextBox txtImporteACta 
               Alignment       =   1  'Right Justify
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
               Left            =   1410
               TabIndex        =   46
               Top             =   2985
               Width           =   1185
            End
            Begin VB.CommandButton cmaAceptarACta 
               Caption         =   "A&ceptar"
               Height          =   360
               Left            =   4500
               TabIndex        =   48
               Top             =   2970
               Width           =   900
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaAFavor 
               Height          =   2205
               Left            =   90
               TabIndex        =   44
               Top             =   285
               Width           =   5310
               _ExtentX        =   9366
               _ExtentY        =   3889
               _Version        =   393216
               Cols            =   6
               FixedCols       =   0
               BackColorSel    =   8388736
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Saldo:"
               Height          =   195
               Left            =   915
               TabIndex        =   137
               Top             =   2655
               Width           =   450
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Left            =   795
               TabIndex        =   136
               Top             =   3030
               Width           =   570
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Comprobantes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3435
            Left            =   -74940
            TabIndex        =   126
            Top             =   480
            Width           =   5535
            Begin MSComCtl2.DTPicker FechaComprobantes 
               Height          =   315
               Left            =   1335
               TabIndex        =   36
               Top             =   637
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   58392577
               CurrentDate     =   41098
            End
            Begin VB.TextBox txtImporteComprobante 
               Height          =   315
               Left            =   1335
               MaxLength       =   15
               TabIndex        =   39
               Top             =   1320
               Width           =   1140
            End
            Begin VB.TextBox txtNroCompSuc 
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
               Left            =   1335
               MaxLength       =   4
               TabIndex        =   37
               Top             =   975
               Width           =   555
            End
            Begin VB.CommandButton cmdCancelarComprobante 
               Caption         =   "Cancelar"
               Height          =   360
               Left            =   1305
               TabIndex        =   43
               Top             =   2970
               Width           =   960
            End
            Begin VB.CommandButton cmdAceptarComprobante 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   330
               TabIndex        =   41
               Top             =   2970
               Width           =   960
            End
            Begin VB.TextBox txtNroComprobantes 
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
               Left            =   1905
               MaxLength       =   8
               TabIndex        =   38
               Top             =   975
               Width           =   1065
            End
            Begin VB.ComboBox cboComprobantes 
               Height          =   315
               Left            =   1335
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   300
               Width           =   2970
            End
            Begin VB.CommandButton cmdAgregarComprobante 
               Caption         =   "Agregar"
               Height          =   345
               Left            =   4485
               TabIndex        =   40
               Top             =   1290
               Width           =   720
            End
            Begin VB.TextBox txtTotalComprobante 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
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
               Left            =   3975
               TabIndex        =   127
               Top             =   2940
               Width           =   1035
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaComp 
               Height          =   1245
               Left            =   315
               TabIndex        =   128
               Top             =   1665
               Width           =   4950
               _ExtentX        =   8731
               _ExtentY        =   2196
               _Version        =   393216
               Cols            =   5
               FixedCols       =   0
               BackColorSel    =   8388736
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Left            =   720
               TabIndex        =   133
               Top             =   1380
               Width           =   570
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Comprobante:"
               Height          =   195
               Left            =   300
               TabIndex        =   132
               Top             =   1050
               Width           =   990
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   930
               TabIndex        =   131
               Top             =   330
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fecha:"
               Height          =   195
               Index           =   3
               Left            =   795
               TabIndex        =   130
               Top             =   690
               Width           =   495
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   3525
               TabIndex        =   129
               Top             =   3000
               Width           =   405
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3435
            Left            =   -74940
            TabIndex        =   111
            Top             =   480
            Width           =   5532
            Begin VB.CommandButton cmdCancelarMoneda 
               Caption         =   "Cancelar"
               Height          =   360
               Left            =   2115
               TabIndex        =   34
               Top             =   2955
               Width           =   960
            End
            Begin VB.CommandButton cmdAceptarMoneda 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   1140
               TabIndex        =   33
               Top             =   2955
               Width           =   960
            End
            Begin VB.TextBox txtTotalEfectivo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
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
               Left            =   3030
               TabIndex        =   115
               Top             =   2505
               Width           =   1035
            End
            Begin VB.ComboBox cboMoneda 
               Height          =   315
               Left            =   1125
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   495
               Width           =   1950
            End
            Begin VB.TextBox txtEftImporte 
               Height          =   315
               Left            =   1125
               TabIndex        =   31
               Top             =   930
               Width           =   1005
            End
            Begin VB.CommandButton cmdAgregarEfectivo 
               Caption         =   "Agregar"
               Height          =   345
               Left            =   2370
               TabIndex        =   32
               Top             =   915
               Width           =   720
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaEfectivo 
               Height          =   1095
               Left            =   1095
               TabIndex        =   112
               Top             =   1350
               Width           =   3285
               _ExtentX        =   5794
               _ExtentY        =   1931
               _Version        =   393216
               Cols            =   3
               FixedCols       =   0
               BackColorSel    =   8388736
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   2580
               TabIndex        =   116
               Top             =   2565
               Width           =   405
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Moneda:"
               Height          =   195
               Left            =   480
               TabIndex        =   114
               Top             =   525
               Width           =   630
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Index           =   2
               Left            =   540
               TabIndex        =   113
               Top             =   975
               Width           =   570
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Cheques"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3435
            Left            =   60
            TabIndex        =   96
            Top             =   480
            Width           =   5535
            Begin VB.CommandButton cmdBuscaCheque 
               Height          =   315
               Left            =   3120
               MaskColor       =   &H000000FF&
               Picture         =   "frmReciboCliente.frx":00FC
               Style           =   1  'Graphical
               TabIndex        =   139
               ToolTipText     =   "Buscar Cheques"
               Top             =   330
               UseMaskColor    =   -1  'True
               Width           =   405
            End
            Begin VB.CommandButton cmdCancelarCheques 
               Caption         =   "Cancelar"
               Height          =   360
               Left            =   1560
               TabIndex        =   29
               Top             =   3015
               Width           =   960
            End
            Begin VB.CommandButton cmdAceptarCheques 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   585
               TabIndex        =   28
               Top             =   3015
               Width           =   960
            End
            Begin VB.Frame frameBanco 
               Caption         =   "Banco"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   90
               TabIndex        =   101
               Top             =   660
               Width           =   4635
               Begin VB.TextBox TxtSUCURSAL 
                  Height          =   285
                  Left            =   2280
                  MaxLength       =   3
                  TabIndex        =   25
                  Top             =   255
                  Width           =   450
               End
               Begin VB.TextBox TxtBANCO 
                  Height          =   285
                  Left            =   525
                  MaxLength       =   3
                  TabIndex        =   23
                  Top             =   240
                  Width           =   450
               End
               Begin VB.TextBox TxtLOCALIDAD 
                  Height          =   285
                  Left            =   1410
                  MaxLength       =   3
                  TabIndex        =   24
                  Top             =   240
                  Width           =   450
               End
               Begin VB.TextBox TxtCODIGO 
                  Height          =   285
                  Left            =   3360
                  MaxLength       =   6
                  TabIndex        =   26
                  Top             =   255
                  Width           =   765
               End
               Begin VB.TextBox TxtBanDescri 
                  BackColor       =   &H00C0C0C0&
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
                  Left            =   60
                  TabIndex        =   104
                  Top             =   615
                  Width           =   4500
               End
               Begin VB.TextBox TxtCodInt 
                  BackColor       =   &H80000018&
                  Height          =   300
                  Left            =   2670
                  TabIndex        =   103
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   420
               End
               Begin VB.CommandButton CmdBanco 
                  DisabledPicture =   "frmReciboCliente.frx":0406
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
                  Left            =   4170
                  Picture         =   "frmReciboCliente.frx":0710
                  Style           =   1  'Graphical
                  TabIndex        =   102
                  Top             =   225
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Loc:"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   11
                  Left            =   1035
                  TabIndex        =   108
                  Top             =   270
                  Width           =   315
               End
               Begin VB.Label lbl 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bco:"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   10
                  Left            =   150
                  TabIndex        =   107
                  Top             =   270
                  Width           =   330
               End
               Begin VB.Label lbl 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Suc:"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   5
                  Left            =   1935
                  TabIndex        =   106
                  Top             =   270
                  Width           =   330
               End
               Begin VB.Label lbl 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código:"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   0
                  Left            =   2790
                  TabIndex        =   105
                  Top             =   285
                  Width           =   540
               End
            End
            Begin VB.TextBox TxtCheNumero 
               Height          =   315
               Left            =   1110
               MaxLength       =   10
               TabIndex        =   22
               Top             =   330
               Width           =   1380
            End
            Begin VB.CommandButton cmdNuevoCheque 
               Height          =   315
               Left            =   2610
               MaskColor       =   &H000000FF&
               Picture         =   "frmReciboCliente.frx":085A
               Style           =   1  'Graphical
               TabIndex        =   100
               ToolTipText     =   "Cargar Cheques"
               Top             =   330
               UseMaskColor    =   -1  'True
               Width           =   405
            End
            Begin VB.CommandButton cmdAgregarCheque 
               Caption         =   "Agregar"
               Height          =   345
               Left            =   4755
               TabIndex        =   27
               Top             =   1425
               Width           =   720
            End
            Begin VB.TextBox txtTotalCheques 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
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
               Left            =   4305
               TabIndex        =   99
               Top             =   3045
               Width           =   1035
            End
            Begin VB.TextBox TxtCheImport 
               Height          =   330
               Left            =   3780
               TabIndex        =   97
               Top             =   315
               Width           =   900
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaCheques 
               Height          =   1170
               Left            =   75
               TabIndex        =   98
               Top             =   1815
               Width           =   5385
               _ExtentX        =   9499
               _ExtentY        =   2064
               _Version        =   393216
               Cols            =   9
               FixedCols       =   0
               BackColorSel    =   8388736
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
            End
            Begin MSComCtl2.DTPicker TxtCheFecVto 
               Height          =   315
               Left            =   3960
               TabIndex        =   142
               Top             =   480
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   58392577
               CurrentDate     =   41098
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nro Cheque:"
               Height          =   195
               Index           =   7
               Left            =   90
               TabIndex        =   110
               Top             =   375
               Width           =   900
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   3840
               TabIndex        =   109
               Top             =   3105
               Width           =   405
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Valores Recibidos..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3405
            Left            =   -74940
            TabIndex        =   91
            Top             =   480
            Width           =   5565
            Begin VB.CommandButton cmdAgregarVALCTA 
               Caption         =   "Agregar Val"
               Height          =   360
               Left            =   3375
               TabIndex        =   18
               Top             =   525
               Width           =   1065
            End
            Begin VB.CommandButton cmdAceptarValores 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   4755
               TabIndex        =   20
               Top             =   180
               Width           =   705
            End
            Begin VB.CommandButton cmdAgregarEFT 
               Caption         =   "Agregar Eft"
               Height          =   360
               Left            =   1215
               TabIndex        =   16
               Top             =   525
               Width           =   1065
            End
            Begin VB.CommandButton cmdAgregarCOMP 
               Caption         =   "Agregar Com"
               Height          =   360
               Left            =   2295
               TabIndex        =   17
               Top             =   525
               Width           =   1065
            End
            Begin VB.CommandButton cmdQuitarVal 
               Caption         =   "&Quitar"
               Height          =   360
               Left            =   4755
               TabIndex        =   19
               Top             =   540
               Width           =   705
            End
            Begin VB.CommandButton cmdAgregarCHE 
               Caption         =   "Agregar Che"
               Height          =   360
               Left            =   135
               TabIndex        =   15
               Top             =   525
               Width           =   1065
            End
            Begin VB.TextBox txtTotalValores 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
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
               Left            =   840
               TabIndex        =   92
               Top             =   2895
               Width           =   1170
            End
            Begin MSFlexGridLib.MSFlexGrid grillaValores 
               Height          =   1875
               Left            =   105
               TabIndex        =   93
               Top             =   915
               Width           =   5415
               _ExtentX        =   9551
               _ExtentY        =   3307
               _Version        =   393216
               Cols            =   6
               FixedCols       =   0
               RowHeightMin    =   250
               BackColorSel    =   8388736
               FocusRect       =   0
               HighLight       =   0
               SelectionMode   =   1
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   360
               TabIndex        =   125
               Top             =   2940
               Width           =   405
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Total Valores Recibidos:"
               Height          =   195
               Left            =   360
               TabIndex        =   94
               Top             =   3420
               Width           =   1725
            End
         End
      End
      Begin VB.Frame FrameRemito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   4605
         TabIndex        =   64
         Top             =   360
         Width           =   6900
         Begin VB.TextBox txtCliLocalidad 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3825
            TabIndex        =   89
            Top             =   720
            Width           =   2925
         End
         Begin VB.TextBox txtProvincia 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1155
            TabIndex        =   88
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox txtCliRazSoc 
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
            Left            =   2385
            MaxLength       =   50
            TabIndex        =   87
            Tag             =   "Descripción"
            Top             =   390
            Width           =   4365
         End
         Begin VB.TextBox txtDomici 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1155
            MaxLength       =   50
            TabIndex        =   86
            Top             =   1050
            Width           =   5595
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   1950
            MaskColor       =   &H000000FF&
            Picture         =   "frmReciboCliente.frx":0BE4
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Buscar Cliente"
            Top             =   390
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtCodCliente 
            Height          =   300
            Left            =   1155
            MaxLength       =   40
            TabIndex        =   5
            Top             =   390
            Width           =   750
         End
         Begin VB.ComboBox CboVend 
            Height          =   315
            Left            =   1155
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1365
            Width           =   2805
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   390
            TabIndex        =   84
            Top             =   750
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   420
            TabIndex        =   83
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Loc.:"
            Height          =   195
            Left            =   3420
            TabIndex        =   82
            Top             =   765
            Width           =   360
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   81
            Top             =   1410
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Recibimos de:"
            Height          =   195
            Left            =   90
            TabIndex        =   80
            Top             =   420
            Width           =   1005
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
         Height          =   1845
         Left            =   -74715
         TabIndex        =   65
         Top             =   540
         Width           =   11115
         Begin VB.CommandButton cmdBuscarVen 
            Height          =   300
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmReciboCliente.frx":0EEE
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "Buscar Vendedor"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   300
            TabIndex        =   53
            Top             =   465
            Width           =   855
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   300
            TabIndex        =   55
            Top             =   990
            Width           =   810
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3360
            MaxLength       =   40
            TabIndex        =   57
            Top             =   255
            Width           =   975
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
            TabIndex        =   68
            Tag             =   "Descripción"
            Top             =   255
            Width           =   4620
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1470
            Left            =   10245
            MaskColor       =   &H000000FF&
            Picture         =   "frmReciboCliente.frx":11F8
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Buscar "
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   300
            TabIndex        =   54
            Top             =   727
            Width           =   1035
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
            Left            =   4845
            TabIndex        =   67
            Top             =   600
            Width           =   4620
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3360
            TabIndex        =   58
            Top             =   585
            Width           =   990
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmReciboCliente.frx":399A
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Buscar Cliente"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CheckBox chkTipoRecibo 
            Caption         =   "Tipo de Recibo"
            Height          =   195
            Left            =   300
            TabIndex        =   56
            Top             =   1245
            Width           =   1485
         End
         Begin VB.ComboBox cboRecibo1 
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   1260
            Width           =   2400
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   3360
            TabIndex        =   59
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58392577
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   6090
            TabIndex        =   60
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58392577
            CurrentDate     =   41098
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
            TabIndex        =   73
            Top             =   300
            Width           =   525
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2265
            TabIndex        =   72
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   5055
            TabIndex        =   71
            Top             =   1020
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Index           =   0
            Left            =   2535
            TabIndex        =   70
            Top             =   615
            Width           =   735
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Recibo:"
            Height          =   195
            Left            =   2355
            TabIndex        =   69
            Top             =   1305
            Width           =   915
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   3780
         Left            =   -74730
         TabIndex        =   63
         Top             =   2475
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6668
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Concepto:"
         Height          =   195
         Left            =   240
         TabIndex        =   144
         Top             =   6360
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   79
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Recibo"
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
      Left            =   5730
      TabIndex        =   138
      Top             =   7110
      Width           =   2100
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
      Left            =   180
      TabIndex        =   90
      Top             =   6630
      Width           =   750
   End
End
Attribute VB_Name = "frmReciboCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim TotFac As Double
Dim Estado As Integer
 
Private Function SumaGrilla(Grilla As MSFlexGrid, COLUMNA As Integer) As String
    Dim Suma As Double
    Suma = 0
    For I = 1 To Grilla.Rows - 1
        Suma = Suma + CDbl(Grilla.TextMatrix(I, COLUMNA))
    Next
    SumaGrilla = Valido_Importe(CStr(Suma))
End Function

Private Sub CboVend_LostFocus()
    If txtCodCliente.Text <> "" Then
        GrillaAplicar.HighLight = flexHighlightAlways
        tabComprobantes.Tab = 1
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim j As Integer
    If MsgBox("¿Confirma Impresión del Recibo?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'PONE A LA IMPRESORA  COMO PREDETERMINADA
    Dim X As Printer
    Dim mDriver As String
    mDriver = IMPRESORA
    For Each X In Printers
        If UCase(X.DeviceName) = UCase(mDriver) Then
            ' La define como predeterminada del sistema.
            Set Printer = X
            Exit For
        End If
    Next
'-----------------------------------
    Set_Impresora
    'SeteoImpresora(256, 1, 7, -1, "Roman 10cpi", 10, False, 12220, 7950)
    j = 1
    For j = 1 To 2
        ImprimirFactura CDbl(j)
    Next j
    Printer.EndDoc
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub
Public Sub ImprimirFactura(Fila As Double)
    Dim Renglon As Double
    Dim hayVuelto As Integer
    Dim impVuelto As String
    hayVuelto = 0 'o no hay, 1 si hay
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Imprimiendo..."
    
    'imprimir por duplicado
    If Fila = 2 Then Fila = 15.77
    ImprimirEncabezado Fila
    
    '---- IMPRESION DE LA FACTURA ------------------
    Renglon = 4
    'Printer.FontSize = 6
    
    'CAMBIAR LA GRILLA Y EL FORMATO DE LA SALIDA IMPRESA DEL RECIBO
    
    Imprimir 3, Fila + Renglon + 0.5, False, "PESOS " & UCase(EnLetras(txtTotalValores.Text))
    'Imprimir 4, Fila + Renglon + 1.18, False, txtObservaciones.Text
    justifica_printer 4, 20, Fila + Renglon + 1.18, txtObservaciones.Text

    For I = 1 To grillaValores.Rows - 1
        If grillaValores.TextMatrix(I, 0) = "CHE" Then
            Imprimir 2, Fila + Renglon + 2.8, False, Trim(grillaValores.TextMatrix(I, 4))
            Imprimir 5, Fila + Renglon + 2.8, False, Trim(grillaValores.TextMatrix(I, 3))
            
            Imprimir 13.8, Fila + Renglon + 2.8, False, Trim(grillaValores.TextMatrix(I, 2))
            Imprimir 16.7, Fila + Renglon + 2.8, False, "$" & Trim(grillaValores.TextMatrix(I, 1))
            
           

            Renglon = Renglon + 0.55
        Else
            If grillaValores.TextMatrix(I, 0) = "EFT" Then
                If grillaValores.TextMatrix(I, 3) = "TRANSFERENCIA" Then
                    Imprimir 3, Fila + 10.2, False, "$" & Trim(grillaValores.TextMatrix(I, 1)) & "  " & Trim(grillaValores.TextMatrix(I, 3))
                Else
                    Imprimir 3, Fila + 10.2, False, "$" & Trim(grillaValores.TextMatrix(I, 1))
                End If
                
                
                If grillaValores.TextMatrix(I, 3) = "VUELTO" Then
                    hayVuelto = 1
                    Imprimir 1, Fila + 10.1, False, "VUELTO "
                    'impVuelto = Trim(grillaValores.TextMatrix(I, 1))
                End If
            Else
                If grillaValores.TextMatrix(I, 0) = "FAC" Or grillaValores.TextMatrix(I, 0) = "COMP" Then
                    
                    If Left(grillaValores.TextMatrix(I, 3), 3) = "RET" Then
                        Imprimir 2, Fila + Renglon + 2.8, False, TipoCBT(grillaValores.TextMatrix(I, 5))
                        Imprimir 5, Fila + Renglon + 2.8, False, Trim(grillaValores.TextMatrix(I, 4))
                    
                        Imprimir 13.8, Fila + Renglon + 2.8, False, Format(Trim(grillaValores.TextMatrix(I, 2)), "mm/dd/yyyy")
                        Imprimir 16.7, Fila + Renglon + 2.8, False, "$" & Trim(grillaValores.TextMatrix(I, 1))
                    
                    Else
                        Imprimir 2, Fila + Renglon + 2.8, False, "FACTURA A"
                        Imprimir 5, Fila + Renglon + 2.8, False, Trim(grillaValores.TextMatrix(I, 4))
                        
                        Imprimir 13.8, Fila + Renglon + 2.8, False, Trim(grillaValores.TextMatrix(I, 2))
                        Imprimir 16.7, Fila + Renglon + 2.8, False, "$" & Trim(grillaValores.TextMatrix(I, 1))
                    End If
                   
        
                    Renglon = Renglon + 0.55
                End If
            End If
        End If
    Next I
    
    'IMPRIMO PAGARE
    'HOJA 1

    'EFECTIVO
'    If hayVuelto = 1 Then
'        Imprimir 1, Fila + 9.5, False, "VUELTO " & Valido_Importe(impVuelto)
'    End If
    
    'TOTAL
    Imprimir 2.3, Fila + 11.2, False, "$ " & Valido_Importe(txtTotalValores.Text)



End Sub

Public Sub ImprimirEncabezado(row As Double)
 '-----------IMPRIME EL ENCABEZADO DE RECIBO-------------------

    Dim año As String
    
    'Imprimir 1, row - 1, False, "-ROSA VIEJO-"
    'Imprimir 9.5, row, False, "RECIBO"
    'Imprimir 14.5, row + 1, False, "Numero: " & txtNroSucursal.Text & " - " & txtNroRecibo.Text
    'Imprimir 14.5, row + 1.5, False, "Fecha:  " & FechaRecibo

    año = Year(FechaRecibo)
    Imprimir 15, row + 1, False, Format(Day(FechaRecibo), "00")
    Imprimir 16.5, row + 1, False, Format(Month(FechaRecibo), "00")
    Imprimir 17.8, row + 1, False, Mid(año, 3, 2)
        
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI, L.LOC_DESCRI, C.CLI_CUIT, C.IVA_CODIGO"
    sql = sql & ", P.PRO_DESCRI"
    sql = sql & " FROM CLIENTE C, LOCALIDAD L,"
    sql = sql & " PROVINCIA P"
    sql = sql & " WHERE "
    sql = sql & " C.LOC_CODIGO = L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND CLI_CODIGO=" & XN(txtCodCliente.Text)
'    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
'    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
'    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
'    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    

    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        'Hoja 1
        Imprimir 3.5, row + 1.8, False, Trim(Rec1!CLI_RAZSOC)
        Imprimir 3.5, row + 2.5, False, Trim(IIf(IsNull(Rec1!CLI_DOMICI), "", Rec1!CLI_DOMICI))
        Imprimir 3.5, row + 3.2, False, Trim(IIf(IsNull(Rec1!LOC_DESCRI), "", Rec1!LOC_DESCRI)) & " - " & Trim(IIf(IsNull(Rec1!PRO_DESCRI), "", Rec1!PRO_DESCRI))
        Imprimir 12, row + 2.5, False, Trim(IIf(IsNull(Rec1!CLI_CUIT), "", Format(Rec1!CLI_CUIT, "##-########-#")))
                
        Select Case Rec1!IVA_CODIGO
            Case 1
                Imprimir 5.56, row + 3.9, False, "X"
            Case 3
                Imprimir 16.7, row + 3.9, False, "X"
            Case 4
                Imprimir 12.1, row + 3.9, False, "X"
            Case 5
                Imprimir 9.1, row + 3.9, False, "X"
            End Select
        
        
    End If
    Rec1.Close
'    'busco forma de pago
'    sql = "SELECT FP.FPG_DESCRI"
'    sql = sql & " FROM FACTURA_CLIENTE FC, FORMA_PAGO FP"
'    sql = sql & " WHERE FC.FPG_CODIGO=FP.FPG_CODIGO"
'    sql = sql & " AND FC.TCO_CODIGO = " & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
'    sql = sql & " AND FC.FCL_NUMERO = " & XN(txtNrorECIBO.Text)
'    sql = sql & " AND FC.FCL_SUCURSAL = " & XN(txtNroSucursal.Text)
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'    If rec.EOF = False Then
'        Imprimir 1, row + 2, False, "Forma de Pago:" & rec!FPG_DESCRI
'
'    End If
'
'    rec.Close
    'Hoja 1
    'Imprimir 14.5, row + 2, False, "Vendedor: " & cboVendedor.Text
'    Imprimir 1, row + 4, False, "CODIGO"
'    Imprimir 3, row + 4, False, "DETALLE"
'    Imprimir 9.5, row + 4, False, "TALLE"
'    Imprimir 11, row + 4, False, "COLOR"
'    Imprimir 14, row + 4, False, "CANTIDAD"
'    Imprimir 16.2, row + 4, False, "PRECIO"
'    Imprimir 17.8, row + 4, False, "TOTAL"
    
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
Private Sub chkTipoRecibo_Click()
    If chkTipoRecibo.Value = Checked Then
        cboRecibo1.Enabled = True
        cboRecibo1.ListIndex = 0
    Else
        cboRecibo1.Enabled = False
        cboRecibo1.ListIndex = -1
    End If
End Sub

Private Sub chkTipoRecibo_LostFocus()
    If chkTipoRecibo.Value = Checked And chkCliente.Value = Unchecked _
        And chkVendedor.Value = Unchecked _
        And chkFecha.Value = Unchecked Then cboRecibo1.SetFocus
End Sub

Private Sub chkVendedor_Click()
    If chkVendedor.Value = Checked Then
        txtVendedor.Enabled = True
        cmdBuscarVen.Enabled = True
    Else
        txtVendedor.Enabled = False
        cmdBuscarVen.Enabled = False
    End If
End Sub

Private Sub cmaAceptarACta_Click()
    txtSaldoACta.Text = ""
    txtImporteACta.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdAceptarCheques_Click()
    Dim I, j As Integer
    'recorro grilla para saber si hay 6 cheques cargados (limite de impresion de cheques en recibo)
    For I = 1 To grillaValores.Rows - 1
        If grillaValores.TextMatrix(I, 0) = "CHE" Then
            j = j + 1
        End If
    Next
    If j < 6 Then
    
        If GrillaCheques.Rows > 1 Then
            'CARGO EN GRILLA VALORES
            For I = 1 To GrillaCheques.Rows - 1
                grillaValores.AddItem "CHE" & Chr(9) & GrillaCheques.TextMatrix(I, 6) & Chr(9) & _
                                      GrillaCheques.TextMatrix(I, 5) & Chr(9) & GrillaCheques.TextMatrix(I, 8) _
                                      & Chr(9) & GrillaCheques.TextMatrix(I, 4) & Chr(9) & GrillaCheques.TextMatrix(I, 7)
            Next
            txtTotalValores.Text = Valido_Importe(CStr(SumaGrilla(grillaValores, 1)))
            grillaValores.HighLight = flexHighlightAlways
            GrillaCheques.Rows = 1
            txtTotalCheques.Text = ""
            tabValores.Tab = 0
        End If
    Else
        MsgBox "Ha superado el numero de cheques por recibo, cargue un nuevo recibo", vbInformation, TIT_MSGBOX
    End If
End Sub

Private Sub cmdAceptarComprobante_Click()
    If GrillaComp.Rows > 1 Then
        'CARGO EN GRILLA VALORES
        For I = 1 To GrillaComp.Rows - 1
            grillaValores.AddItem "FAC" & Chr(9) & GrillaComp.TextMatrix(I, 3) & Chr(9) & GrillaComp.TextMatrix(I, 2) _
                                   & Chr(9) & GrillaComp.TextMatrix(I, 0) & Chr(9) & GrillaComp.TextMatrix(I, 1) & Chr(9) & _
                                   GrillaComp.TextMatrix(I, 4)
        Next
        txtTotalValores.Text = Valido_Importe(CStr(SumaGrilla(grillaValores, 1)))
        grillaValores.HighLight = flexHighlightAlways
        GrillaComp.Rows = 1
        txtNroComprobantes.Text = ""
        txtNroCompSuc.Text = ""
        fechaComprobantes.Value = Null
        cboComprobantes.ListIndex = 0
        txtTotalComprobante.Text = ""
        tabValores.Tab = 0
    End If
End Sub

Private Sub cmdAceptarComprobantes_Click()
    Dim obserPT, obserSA, obserAC As String
    Dim obserFACPT, obserFACSA, obserFACAC As String
    obserPT = ""
    obserSA = ""
    obserAC = ""
    obserFACPT = ""
    obserFACSA = ""
    obserFACAC = ""
    
    txtSaldo.Text = ""
    txtImporteApagar.Text = ""
    tabComprobantes.Tab = 0
    
    'Armo la observacion automatica para una FAC a pagar
    If GrillaAplicar1.Rows = 2 Then
        If GrillaAplicar1.TextMatrix(1, 1) = GrillaAplicar1.TextMatrix(1, 6) Then
            If GrillaAplicar1.TextMatrix(1, 4) = "0,00" Then
                txtObservaciones.Text = "PAGO TOTAL SEGUN " & GrillaAplicar1.TextMatrix(1, 0) & " " & GrillaAplicar1.TextMatrix(1, 2)
            Else
                txtObservaciones.Text = "PAGO A CUENTA SEGUN " & GrillaAplicar1.TextMatrix(1, 0) & " " & GrillaAplicar1.TextMatrix(1, 2)
            End If
        Else
            If GrillaAplicar1.TextMatrix(1, 4) = "0,00" Then
                txtObservaciones.Text = "PAGO SALDO TOTAL SEGUN " & GrillaAplicar1.TextMatrix(1, 0) & " " & GrillaAplicar1.TextMatrix(1, 2)
            Else
                txtObservaciones.Text = "PAGO A CUENTA SEGUN " & GrillaAplicar1.TextMatrix(1, 0) & " " & GrillaAplicar1.TextMatrix(1, 2)
            End If
        End If
    Else
        If GrillaAplicar1.Rows > 2 Then
            For I = 1 To GrillaAplicar1.Rows - 1
                If GrillaAplicar1.TextMatrix(I, 1) = GrillaAplicar1.TextMatrix(I, 6) Then
                    If GrillaAplicar1.TextMatrix(I, 4) = "0,00" Then
                        obserPT = "PAGO TOTAL SEGUN "
                        obserFACPT = obserFACPT & GrillaAplicar1.TextMatrix(I, 0) & " " & GrillaAplicar1.TextMatrix(I, 2) & "  "
                    Else
                        obserAC = "PAGO A CUENTA SEGUN "
                        obserFACAC = obserFACAC & GrillaAplicar1.TextMatrix(I, 0) & " " & GrillaAplicar1.TextMatrix(I, 2) & "  "
                    End If
                Else
                    If GrillaAplicar1.TextMatrix(I, 4) = "0,00" Then
                        obserSA = "PAGO SALDO TOTAL SEGUN "
                        obserFACSA = obserFACSA & GrillaAplicar1.TextMatrix(I, 0) & " " & GrillaAplicar1.TextMatrix(I, 2) & "  "
                    Else
                        obserAC = "PAGO A CUENTA SEGUN "
                        obserFACAC = obserFACAC & GrillaAplicar1.TextMatrix(I, 0) & " " & GrillaAplicar1.TextMatrix(I, 2) & "  "
                    End If
                End If
            Next
        End If
        If obserPT <> "" Then
            If obserAC <> "" Then
                obserAC = " y " & obserAC
                obserAC = " " & obserAC
            End If
            If obserSA <> "" Then
                obserSA = " y " & obserSA
                obserSA = " " & obserSA
            End If
        Else
            If obserSA <> "" Then
                If obserAC <> "" Then
                    obserAC = " y " & obserAC
                Else
                    obserAC = " " & obserAC
                End If
            End If
        End If
        txtObservaciones.Text = obserPT & " " & obserFACPT & obserSA & " " & obserFACSA & obserAC & " " & obserFACAC
        
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


Private Sub cmdAceptarFacturas_Click()
    cmdAgregarCHE.SetFocus
End Sub

Private Sub cmdAceptarMoneda_Click()
    
    If GrillaEfectivo.Rows > 1 Then
        'CARGO EN GRILLA VALORES
        For I = 1 To GrillaEfectivo.Rows - 1
            grillaValores.AddItem "EFT" & Chr(9) & GrillaEfectivo.TextMatrix(I, 1) & Chr(9) & "" _
                                   & Chr(9) & GrillaEfectivo.TextMatrix(I, 0) & Chr(9) & "" & Chr(9) & _
                                   GrillaEfectivo.TextMatrix(I, 2)
        Next
        txtTotalValores.Text = Valido_Importe(CStr(SumaGrilla(grillaValores, 1)))
        grillaValores.HighLight = flexHighlightAlways
        GrillaEfectivo.Rows = 1
        txtTotalEfectivo.Text = ""
        tabValores.Tab = 0
    End If
End Sub

Private Sub cmdAceptarValores_Click()
    If CmdGrabar.Enabled = True Then
        CmdGrabar.SetFocus
    Else
        CmdNuevo.SetFocus
    End If
End Sub

Private Sub cmdAgregarACta_Click()
    If GrillaAFavor.Rows > 1 Then
        If grillaValores.Rows > 1 Then
            For I = 1 To grillaValores.Rows - 1
                If grillaValores.TextMatrix(I, 0) = "COMP" Then
                    If GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 5) = grillaValores.TextMatrix(I, 5) _
                        And CLng(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 1)) = CLng(grillaValores.TextMatrix(I, 4)) _
                        And CDate(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 2)) = CDate(grillaValores.TextMatrix(I, 2)) Then
                       MsgBox "El Valor ya fue ingresado", vbInformation, TIT_MSGBOX
                       txtSaldoACta.Text = ""
                       txtImporteACta.Text = ""
                       GrillaAFavor.SetFocus
                       Exit Sub
                    End If
                End If
            Next
        End If
'        If GrillaAFavor.CellForeColor = vbBlack Then
'            Call CambiaColorAFilaDeGrilla(GrillaAFavor, GrillaAFavor.RowSel, vbRed)
'        Else
'            Call CambiaColorAFilaDeGrilla(GrillaAFavor, GrillaAFavor.RowSel, vbBlack)
'        End If
                
        'CARGO EN GRILLA VALORES
        grillaValores.AddItem "FAC" & Chr(9) & Valido_Importe(txtImporteACta) & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 2) _
                                & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 0) & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 1) & Chr(9) & _
                                GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 5)

        'ARREGLO EL SALDO DEL DINERO A CTA
        GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4) = Valido_Importe(CStr(CDbl(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4)) - CDbl(txtImporteACta.Text)))
        
        txtTotalValores.Text = Valido_Importe(CStr(SumaGrilla(grillaValores, 1)))
        grillaValores.HighLight = flexHighlightAlways

        txtSaldoACta.Text = ""
        txtImporteACta.Text = ""
        GrillaAFavor.SetFocus
    End If
End Sub

Private Sub cmdAgregarCHE_Click()
    tabValores.Tab = 1
End Sub

Private Sub cmdAgregarCheque_Click()
    If GrillaCheques.Rows = 7 Then
        MsgBox "No se aceptan mas de 6 cheques por Recibo", vbExclamation, TIT_MSGBOX
        Exit Sub
    Else
        
    

        If TxtCheNumero.Text = "" Then
            MsgBox "Debe ingresar el número del cheque", vbExclamation, TIT_MSGBOX
            TxtCheNumero.SetFocus
            Exit Sub
        End If
        If TxtBanco.Text = "" Then
            MsgBox "Debe ingresar el código del banco", vbExclamation, TIT_MSGBOX
            TxtBanco.SetFocus
            Exit Sub
        End If
        If txtlocalidad.Text = "" Then
            MsgBox "Debe ingresar el código del banco", vbExclamation, TIT_MSGBOX
            txtlocalidad.SetFocus
            Exit Sub
        End If
        If TxtSucursal.Text = "" Then
            MsgBox "Debe ingresar el código del banco", vbExclamation, TIT_MSGBOX
            TxtSucursal.SetFocus
            Exit Sub
        End If
        If txtcodigo.Text = "" Then
            MsgBox "Debe ingresar el código del banco", vbExclamation, TIT_MSGBOX
            txtcodigo.SetFocus
            Exit Sub
        End If
        'VALIDO QUE EL CHEQUE NO SE HAYA CARGADO
        If GrillaCheques.Rows > 1 Then
            If ValidoIngCheques = False Then
                MsgBox "El Cheque ya fue ingresado", vbCritical, TIT_MSGBOX
                TxtCheNumero.Text = ""
                TxtCheNumero.SetFocus
                Exit Sub
            End If
        End If
        
        'CARGO GRILLA
        GrillaCheques.AddItem TxtBanco.Text & Chr(9) & txtlocalidad.Text & Chr(9) & _
                              TxtSucursal.Text & Chr(9) & txtcodigo.Text & Chr(9) & _
                              TxtCheNumero.Text & Chr(9) & TxtCheFecVto.Value & Chr(9) & _
                              Valido_Importe(TxtCheImport.Text) & Chr(9) & TxtCodInt.Text & Chr(9) & TxtBanDescri.Text
        
        
        GrillaCheques.HighLight = flexHighlightAlways
        txtTotalCheques.Text = Valido_Importe(CStr(SumaGrilla(GrillaCheques, 6)))
        LimpiarCheques
        cmdAgregarCheque.Enabled = False
        TxtCheNumero.SetFocus
    End If
End Sub

Private Function ValidoIngCheques() As Boolean
    For I = 1 To GrillaCheques.Rows - 1
        If TxtCodInt.Text = GrillaCheques.TextMatrix(I, 7) And _
           TxtCheNumero.Text = GrillaCheques.TextMatrix(I, 4) Then
           
           ValidoIngCheques = False
           Exit Function
        End If
    Next
    ValidoIngCheques = True
End Function

Private Sub LimpiarCheques()
    TxtBanco.Text = ""
    txtlocalidad.Text = ""
    TxtSucursal.Text = ""
    txtcodigo.Text = ""
    TxtCheNumero.Text = ""
    TxtCheFecVto.Value = Null
    TxtCheImport.Text = ""
    TxtCodInt.Text = ""
    TxtBanDescri.Text = ""
    frameBanco.Enabled = False
    cmdAgregarCheque.Enabled = False
End Sub

Private Sub cmdAgregarCOMP_Click()
    tabValores.Tab = 3
End Sub

Private Sub cmdAgregarComprobante_Click()
    
    If cboComprobantes.ListIndex = -1 Then
        MsgBox "Debe seleccionar un tipo de Documento", vbCritical, TIT_MSGBOX
        cboComprobantes.SetFocus
        Exit Sub
    End If
    If txtNroCompSuc.Text = "" Or txtNroComprobantes.Text = "" Then
        MsgBox "Debe ingresar el número del Comprobante", vbCritical, TIT_MSGBOX
        txtNroCompSuc.SetFocus
        Exit Sub
    End If
    If IsNull(fechaComprobantes.Value) Then
        MsgBox "Debe ingresar la fecha del Documento", vbCritical, TIT_MSGBOX
        fechaComprobantes.SetFocus
        Exit Sub
    End If
    If txtImporteComprobante.Text = "" Then
        MsgBox "Debe ingresar el importe del Documento", vbCritical, TIT_MSGBOX
        txtImporteComprobante.SetFocus
        Exit Sub
    End If
    
    'VALIDO QUE EL COMPROBANTE NO SE HAYA CARGADO
    If GrillaComp.Rows > 1 Then
        If ValidoIngFactura(cboComprobantes, GrillaComp, txtNroComprobantes, txtNroCompSuc) = False Then
            MsgBox "El Documento ya fue ingresado", vbCritical, TIT_MSGBOX
            txtNroComprobantes.Text = ""
            txtNroCompSuc.Text = ""
            fechaComprobantes.Value = Null
            cboComprobantes.SetFocus
            Exit Sub
        End If
    End If
    'VALIDO QUE NO INGRESE DE NUEVO EL DOCUMENTO
    If grillaValores.Rows > 1 Then
        For I = 1 To grillaValores.Rows - 1
            If grillaValores.TextMatrix(I, 0) = "COMP" Then
                If cboComprobantes.ItemData(cboComprobantes.ListIndex) = grillaValores.TextMatrix(I, 5) And _
                   txtNroComprobantes = Right(grillaValores.TextMatrix(I, 4), 8) And _
                   txtNroCompSuc = Left(grillaValores.TextMatrix(I, 4), 4) Then
                   
                   MsgBox "El Documento ya fue ingresado", vbCritical, TIT_MSGBOX
                   txtNroComprobantes.Text = ""
                   txtNroCompSuc.Text = ""
                   fechaComprobantes.Value = Null
                   cboComprobantes.SetFocus
                   Exit Sub
                End If
            End If
        Next
    End If
    
    'CARGO GRILLA
    GrillaComp.AddItem BuscarTipoDocAbre(CStr(cboComprobantes.ItemData(cboComprobantes.ListIndex))) _
                       & Chr(9) & txtNroCompSuc & "-" & txtNroComprobantes & Chr(9) & _
                       fechaComprobantes & Chr(9) & txtImporteComprobante.Text & Chr(9) & _
                       cboComprobantes.ItemData(cboComprobantes.ListIndex)

                           
    GrillaComp.HighLight = flexHighlightAlways
    txtTotalComprobante.Text = Valido_Importe(CStr(SumaGrilla(GrillaComp, 3)))
    txtNroComprobantes.Text = ""
    txtNroCompSuc.Text = ""
    fechaComprobantes.Value = Null
    cboComprobantes.SetFocus
End Sub

Private Function BuscarTipoDocAbre(Codigo As String) As String
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT TCO_ABREVIA"
    sql = sql & " FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_CODIGO=" & XN(Codigo)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscarTipoDocAbre = Rec1!TCO_ABREVIA
    Else
        BuscarTipoDocAbre = ""
    End If
    Rec1.Close
End Function

Private Sub cmdAgregarEFT_Click()
    tabValores.Tab = 2
End Sub

Private Sub cmdAgregarFactura_Click()
    tabComprobantes.Tab = 1
End Sub

Private Sub cmdAgregarFacturas_Click()
    
    If GrillaAplicar1.Rows > 1 Then
        For I = 1 To GrillaAplicar1.Rows - 1
            If GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 0) = GrillaAplicar1.TextMatrix(I, 0) _
                And GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 1) = GrillaAplicar1.TextMatrix(I, 2) _
                And CDate(GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 2)) = CDate(GrillaAplicar1.TextMatrix(I, 3)) Then
               MsgBox "La Factura ya fue elegida", vbInformation, TIT_MSGBOX
               txtSaldo.Text = ""
               txtImporteApagar.Text = ""
               GrillaAplicar.SetFocus
               Exit Sub
            End If
            ' armo la observacion
            
        Next
    End If
'    If GrillaAplicar.CellForeColor = vbBlack Then
'        Call CambiaColorAFilaDeGrilla(GrillaAplicar, GrillaAplicar.RowSel, vbRed)
'    Else
'        Call CambiaColorAFilaDeGrilla(GrillaAplicar, GrillaAplicar.RowSel, vbBlack)
'    End If
    GrillaAplicar1.AddItem GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 0) & Chr(9) & _
                           Valido_Importe(txtImporteApagar.Text) & Chr(9) & _
                           GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 1) & Chr(9) & _
                           GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 2) & Chr(9) & _
                           Valido_Importe(CStr(CDbl(txtSaldo.Text) - CDbl(txtImporteApagar.Text))) & Chr(9) & _
                           GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 5) & Chr(9) & _
                           GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 3)
    
    GrillaAplicar1.HighLight = flexHighlightAlways
    txtSaldo.Text = ""
    txtImporteApagar.Text = ""
    GrillaAplicar.SetFocus
    
    
End Sub

Private Sub cmdAgregarEfectivo_Click()
    'VALIDO QUE EL CHEQUE NO SE HAYA CARGADO
    If GrillaEfectivo.Rows > 1 Then
        If ValidoIngMoneda = False Then
            MsgBox "La Moneda ya fue ingresada", vbCritical, TIT_MSGBOX
            txtEftImporte.Text = ""
            cboMoneda.SetFocus
            Exit Sub
        End If
    End If
    'CARGO GRILLA
    If cboMoneda.Text = "VUELTO" Then
        txtEftImporte.Text = -Valido_Importe(txtEftImporte)
    End If
       GrillaEfectivo.AddItem cboMoneda.Text & Chr(9) & Valido_Importe(txtEftImporte.Text) _
                                & Chr(9) & cboMoneda.ItemData(cboMoneda.ListIndex)
                                                   
    GrillaEfectivo.HighLight = flexHighlightAlways
    txtTotalEfectivo.Text = Valido_Importe(CStr(SumaGrilla(GrillaEfectivo, 1)))
    txtEftImporte.Text = ""
    cboMoneda.SetFocus
End Sub

Private Function ValidoIngMoneda() As Boolean
    For I = 1 To GrillaEfectivo.Rows - 1
        If cboMoneda.ItemData(cboMoneda.ListIndex) = GrillaEfectivo.TextMatrix(I, 2) Then
           
           ValidoIngMoneda = False
           Exit Function
        End If
    Next
    ValidoIngMoneda = True
End Function

Private Function ValidoIngFactura(Combo As ComboBox, Grilla As MSFlexGrid, NROFAC As String, NroSuc As String) As Boolean
    For I = 1 To Grilla.Rows - 1
        If Combo.ItemData(Combo.ListIndex) = Grilla.TextMatrix(I, 4) And _
           NROFAC = Right(Grilla.TextMatrix(I, 1), 8) And _
           NroSuc = Left(Grilla.TextMatrix(I, 1), 4) Then
           
           ValidoIngFactura = False
           Exit Function
        End If
    Next
    ValidoIngFactura = True
End Function

Private Sub cmdAgregarVALCTA_Click()
    tabValores.Tab = 4
End Sub

Private Sub CmdBanco_Click()
     ABMBanco.Show vbModal
End Sub

Private Sub cmdBuscaCheque_Click()
    Dim codint As Integer
    frmBuscar.TipoBusqueda = 6
    frmBuscar.Show vbModal
    'TxtCheNumero.Text = frmBuscar.grdBuscar.Col
    TxtCheNumero.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
    TxtBanco.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 5)
    txtlocalidad.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 6)
    TxtSucursal.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 7)
    txtcodigo.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 8)
    TxtCheImport.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
    TxtCheFecVto.Value = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2)
    TxtBanDescri.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
    TxtCodInt.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)
    
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT RC.REC_NUMERO, RC.REC_SUCURSAL,"
    sql = sql & " RC.REC_FECHA, RC.TCO_CODIGO, TC.TCO_ABREVIA, C.CLI_RAZSOC, V.VEN_NOMBRE"
    sql = sql & " FROM RECIBO_CLIENTE RC, CLIENTE C, VENDEDOR V, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " RC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND RC.VEN_CODIGO=V.VEN_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND RC.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND RC.REC_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND RC.REC_FECHA<=" & XDQ(FechaHasta)
    If chkTipoRecibo.Value = Checked Then sql = sql & " AND RC.TCO_CODIGO=" & XN(cboRecibo1.ItemData(cboRecibo1.ListIndex))
    sql = sql & " ORDER BY RC.REC_SUCURSAL, RC.REC_NUMERO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            GrdModulos.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!REC_SUCURSAL, "0000") & "-" & Format(Rec1!REC_NUMERO, "00000000") _
                               & Chr(9) & Rec1!REC_FECHA & Chr(9) & Rec1!CLI_RAZSOC _
                               & Chr(9) & Rec1!VEN_NOMBRE & Chr(9) & Rec1!TCO_CODIGO _
                               & Chr(9) & ""
            Rec1.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        chkCliente.SetFocus
    End If
    Rec1.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCodCliente.Text = frmBuscar.grdBuscar.Text
        txtCodCliente_LostFocus
        'txtCliRazSoc.SetFocus
    Else
        txtCodCliente.SetFocus
    End If
    txtCodCliente.SetFocus
End Sub

Private Sub cmdBuscarVen_Click()
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

Private Sub cmdCancelarCheques_Click()
    GrillaCheques.Rows = 1
    txtTotalCheques.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdCancelarComprobante_Click()
    GrillaComp.Rows = 1
    txtNroComprobantes.Text = ""
    txtNroCompSuc.Text = ""
    fechaComprobantes.Value = Null
    cboComprobantes.ListIndex = 0
    txtTotalComprobante.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdCancelarMoneda_Click()
    GrillaEfectivo.Rows = 1
    txtTotalEfectivo.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdGrabar_Click()
    'Dim vuelto As String
    Set Rec1 = New ADODB.Recordset
    If ValidarRecibo = False Then Exit Sub
    If MsgBox("¿Confirma Recibo?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayError
    DBConn.BeginTrans
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    sql = "SELECT EST_CODIGO"
    sql = sql & " FROM RECIBO_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    sql = sql & " AND REC_NUMERO=" & XN(txtNroRecibo.Text)
    sql = sql & " AND REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = True Then
        
        'CABEZA DEL RECIBO
        sql = "INSERT INTO RECIBO_CLIENTE ("
        sql = sql & " TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
        sql = sql & " EST_CODIGO, CLI_CODIGO, VEN_CODIGO, REC_FECHA_RENDICION,"
        sql = sql & " REC_NUMEROTXT, REC_TOTAL, REC_OBSER)"
        sql = sql & " VALUES ("
        sql = sql & XN(cboRecibo.ItemData(cboRecibo.ListIndex)) & ", "
        sql = sql & XN(txtNroRecibo.Text) & ","
        sql = sql & XN(txtNroSucursal.Text) & ","
        sql = sql & XDQ(FechaRecibo.Value) & ","
        sql = sql & "3," 'ESTADO DEFINITIVO
        sql = sql & XN(txtCodCliente.Text) & ","
        sql = sql & CboVend.ItemData(CboVend.ListIndex) & ","
        sql = sql & XDQ(FechaRendicion.Value) & ","
        sql = sql & XS(Format(txtNroRecibo.Text, "00000000")) & ","
        sql = sql & XN(txtTotalAplicar) & ","
        sql = sql & XS(txtObservaciones.Text) & ")"
        DBConn.Execute sql
        
        'DETALLE DEL RECIBO
        For I = 1 To grillaValores.Rows - 1
            sql = "INSERT INTO DETALLE_RECIBO_CLIENTE"
            sql = sql & " (TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
            sql = sql & " DRE_NROITEM, BAN_CODINT, CHE_NUMERO,MON_CODIGO,"
            sql = sql & " DRE_MONIMP, DRE_TCO_CODIGO, DRE_COMFECHA, DRE_COMNUMERO,"
            sql = sql & " DRE_COMSUCURSAL, DRE_COMIMP,DRE_VUELTO)"
            sql = sql & " VALUES ("
            sql = sql & XN(cboRecibo.ItemData(cboRecibo.ListIndex)) & ","
            sql = sql & XN(txtNroRecibo.Text) & ","
            sql = sql & XN(txtNroSucursal.Text) & ","
            sql = sql & XDQ(FechaRecibo.Value) & ","
            sql = sql & XN(CStr(I)) & ","
            If grillaValores.TextMatrix(I, 0) = "CHE" Then
                sql = sql & XN(grillaValores.TextMatrix(I, 5)) & ","
                sql = sql & XS(grillaValores.TextMatrix(I, 4)) & "," 'NUMERO DE CHEQUE
            Else
                sql = sql & "NULL,NULL,"
            End If
            If grillaValores.TextMatrix(I, 0) = "EFT" Then
                sql = sql & XN(grillaValores.TextMatrix(I, 5)) & ","
                sql = sql & XN(grillaValores.TextMatrix(I, 1)) & ","
            Else
                sql = sql & "NULL,NULL,"
            End If
            If grillaValores.TextMatrix(I, 0) = "COMP" Or grillaValores.TextMatrix(I, 0) = "A-CTA" Or grillaValores.TextMatrix(I, 0) = "FAC" Then
                sql = sql & XN(grillaValores.TextMatrix(I, 5)) & ","
                sql = sql & XDQ(grillaValores.TextMatrix(I, 2)) & ","
                sql = sql & XN(Right(grillaValores.TextMatrix(I, 4), 8)) & "," 'NUMERO COMPROBANTE
                sql = sql & XN(Left(grillaValores.TextMatrix(I, 4), 4)) & ","  'NUMERO SUCURSAL
                sql = sql & XN(grillaValores.TextMatrix(I, 1)) & ","
            Else
                sql = sql & "NULL,NULL,NULL,NULL,NULL,"
            End If
            If grillaValores.TextMatrix(I, 0) = "VUE" Then
                'sql = sql & XN(grillaValores.TextMatrix(I, 5)) & ","
                sql = sql & XN(grillaValores.TextMatrix(I, 1)) & ")"
            Else
                sql = sql & "NULL)"
            End If
            DBConn.Execute sql
        Next
        'FACTURAS CANCELADAS EN EL RECIBO
        For I = 1 To GrillaAplicar1.Rows - 1
            sql = "INSERT INTO FACTURAS_RECIBO_CLIENTE"
            sql = sql & " (TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
            sql = sql & " FCL_TCO_CODIGO, FCL_NUMERO, FCL_SUCURSAL,"
            sql = sql & " FCL_FECHA, REC_IMPORTE)"
            sql = sql & " VALUES ("
            sql = sql & XN(cboRecibo.ItemData(cboRecibo.ListIndex)) & ","
            sql = sql & XN(txtNroRecibo.Text) & ","
            sql = sql & XN(txtNroSucursal.Text) & ","
            sql = sql & XDQ(FechaRecibo) & ","
            sql = sql & XN(GrillaAplicar1.TextMatrix(I, 5)) & ","           'TIPO FACTURA
            sql = sql & XN(Right(GrillaAplicar1.TextMatrix(I, 2), 8)) & "," 'NUMERO FACTURA
            sql = sql & XN(Left(GrillaAplicar1.TextMatrix(I, 2), 4)) & ","  'NUMERO SUCURSAL
            sql = sql & XDQ(GrillaAplicar1.TextMatrix(I, 3)) & ","          'FECHA FACTURA
            sql = sql & XN(GrillaAplicar1.TextMatrix(I, 1)) & ")"           'IMPORTE
            DBConn.Execute sql
        Next
        'ACTUALIZO EL SALDO DE LAS FACTURAS ELEGIDAS
        For I = 1 To GrillaAplicar1.Rows - 1
            sql = "UPDATE FACTURA_CLIENTE"
            sql = sql & " SET FCL_SALDO=" & XN(GrillaAplicar1.TextMatrix(I, 4))
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & XN(GrillaAplicar1.TextMatrix(I, 5))
            sql = sql & " AND FCL_NUMERO=" & XN(Right(GrillaAplicar1.TextMatrix(I, 2), 8))  'NUMERO FACTURA
            sql = sql & " AND FCL_SUCURSAL=" & XN(Left(GrillaAplicar1.TextMatrix(I, 2), 4)) 'NUMERO SUCURSAL
            DBConn.Execute sql
            
            If GrillaAplicar1.TextMatrix(I, 0) = "ND-A" Then
                sql = "UPDATE NOTA_DEBITO_CLIENTE"
                sql = sql & " SET NDC_SALDO=" & XN(GrillaAplicar1.TextMatrix(I, 4))
                sql = sql & " WHERE"
                sql = sql & " TCO_CODIGO=" & XN(GrillaAplicar1.TextMatrix(I, 5))
                sql = sql & " AND NDC_NUMERO=" & XN(Right(GrillaAplicar1.TextMatrix(I, 2), 8))  'NUMERO FACTURA
                sql = sql & " AND NDC_SUCURSAL=" & XN(Left(GrillaAplicar1.TextMatrix(I, 2), 4)) 'NUMERO SUCURSAL
                DBConn.Execute sql
            End If
        Next
        'ACTUALIZO EL DINERO A CUENTA (RECIBO_CLIENTE_SALDO)
        For I = 1 To GrillaAFavor.Rows - 1
            sql = "UPDATE RECIBO_CLIENTE_SALDO"
            sql = sql & " SET REC_SALDO=" & XN(GrillaAFavor.TextMatrix(I, 4))
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & XN(GrillaAFavor.TextMatrix(I, 5))
            sql = sql & " AND REC_NUMERO=" & XN(Right(GrillaAFavor.TextMatrix(I, 1), 8)) 'NUMERO RECIBO
            sql = sql & " AND REC_SUCURSAL=" & XN(Left(GrillaAFavor.TextMatrix(I, 1), 4)) 'NUMERO SUCURSAL
            DBConn.Execute sql
        Next
        'VERIFICO SI HAY DINERO A CUENTA
        'If CDbl(txtTotalAplicar.Text) < CDbl(txtTotalValores.Text) Then
        '    vuelto = "- " & CDbl(txtTotalAplicar.Text) < CDbl(txtTotalValores.Text)
            
        '    grillaValores.AddItem "VUE" & Chr(9) & "Vuelto en efectivo" & Chr(9) & _
                                  GrillaCheques.TextMatrix(I, 5) & Chr(9) & GrillaCheques.TextMatrix(I, 8) _
                                  & Chr(9) & GrillaCheques.TextMatrix(I, 4) & Chr(9) & GrillaCheques.TextMatrix(I, 7)
            'sql = "INSERT INTO RECIBO_CLIENTE_SALDO"
            'sql = sql & " (TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
            'sql = sql & " REC_TOTSALDO, REC_SALDO)"
            'sql = sql & " VALUES ("
            'sql = sql & cboRecibo.ItemData(cboRecibo.ListIndex) & ","
            'sql = sql & XN(txtNroRecibo) & ","
            'sql = sql & XN(txtNroSucursal) & ","
            'sql = sql & XDQ(FechaRecibo) & ","
            'sql = sql & XN(CStr(CDbl(txtTotalValores.Text) - CDbl(txtTotalAplicar.Text))) & ","
            'sql = sql & XN(CStr(CDbl(txtTotalValores.Text) - CDbl(txtTotalAplicar.Text))) & ")"
            'DBConn.Execute sql
        'End If
        
        'ACTUALIZO LA CUENTA CORRIENTE DEL CLIENTE
        DBConn.Execute AgregoCtaCteCliente(txtCodCliente.Text, CStr(cboRecibo.ItemData(cboRecibo.ListIndex)) _
                                            , txtNroRecibo, txtNroSucursal, _
                                            FechaRecibo, txtTotalValores.Text, "H", CStr(Date))
                                                
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO AL RECIBO QUE CORRESPONDA
        sql = "SELECT * FROM PARAMETROS"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
                Select Case cboRecibo.ItemData(cboRecibo.ListIndex)
                    Case 10
                        sql = "UPDATE PARAMETROS SET RECIBO_A=" & XN(txtNroRecibo)
                        DBConn.Execute sql
                    Case 11
                        sql = "UPDATE PARAMETROS SET RECIBO_B=" & XN(txtNroRecibo)
                        DBConn.Execute sql
                End Select
        End If
        Rec1.Close
        
        DBConn.CommitTrans
        
    Else 'SI EXISTE
        MsgBox "El Recibo ya fue Registrado", vbCritical, TIT_MSGBOX
        DBConn.CommitTrans
    End If
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    rec.Close
    
   ' If MsgBox("¿Desea imprimir el recibo Recibo?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    cmdImprimir_Click
    CmdNuevo_Click
    Exit Sub
    
HayError:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarRecibo() As Boolean
Dim vuelto As String
    If txtNroSucursal.Text = "" Or txtNroRecibo.Text = "" Then
        MsgBox "Debe ingresar el número de Recibo", vbCritical, TIT_MSGBOX
        txtNroSucursal.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If IsNull(FechaRecibo.Value) Then
        MsgBox "Debe ingresar la fecha del Recibo", vbCritical, TIT_MSGBOX
        FechaRecibo.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If IsNull(FechaRendicion.Value) Then
        MsgBox "Debe ingresar la fecha de rendicion del Recibo", vbCritical, TIT_MSGBOX
        FechaRendicion.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If txtCodCliente.Text = "" Then
        MsgBox "Debe ingresar un Cliente", vbCritical, TIT_MSGBOX
        txtCodCliente.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If CboVend.ListIndex = -1 Then
        MsgBox "Debe seleccionar un Vendedor", vbCritical, TIT_MSGBOX
        CboVend.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If grillaValores.Rows = 1 Then
        MsgBox "Debe ingresar Valores Recibidos", vbCritical, TIT_MSGBOX
        cmdAgregarCHE.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If GrillaAplicar.Rows = 1 Then
        MsgBox "Debe ingresar una Factura", vbCritical, TIT_MSGBOX
        cmdAgregarFactura.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If CDbl(txtTotalAplicar.Text) > CDbl(txtTotalValores.Text) Then
        MsgBox "El Total de Facturas supera al Total de Valores Recibidos", vbCritical, TIT_MSGBOX
        cmdAgregarCHE.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If CDbl(txtTotalAplicar.Text) < CDbl(txtTotalValores.Text) Then
        If MsgBox("El Total de Valores Recibidos supera al Total de Facturas," & Chr(13) & _
                "deja el importe (" & Format(CStr(CDbl(txtTotalValores.Text) - CDbl(txtTotalAplicar.Text)), "#,##0.00") & _
                ") como dinero a cuenta", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
                
            vuelto = "-" & CDbl(txtTotalValores.Text) - CDbl(txtTotalAplicar.Text)
            
            grillaValores.AddItem "VUE" & Chr(9) & Valido_Importe(vuelto) & Chr(9) & _
                                   "" & Chr(9) & "VUELTO EN PESOS"
                
            txtTotalValores = SumaGrilla(grillaValores, 1)
            cmdAgregarFactura.SetFocus
            ValidarRecibo = True
        Else
            ValidarRecibo = False
            Exit Function
        End If
    End If
    ValidarRecibo = True
End Function

Private Sub CmdNuevo_Click()
    Estado = 1
    CmdGrabar.Enabled = True
    FrameRecibo.Enabled = True
    FrameRemito.Enabled = True
    TxtCheNumero.Text = ""
    GrillaCheques.Rows = 1
    GrillaCheques.HighLight = flexHighlightNever
    txtEftImporte.Text = ""
    GrillaEfectivo.Rows = 1
    GrillaEfectivo.HighLight = flexHighlightNever
    GrillaAplicar.Rows = 1
    GrillaAplicar.HighLight = flexHighlightNever
    GrillaAplicar1.Rows = 1
    GrillaAplicar1.HighLight = flexHighlightNever
    GrillaComp.Rows = 1
    GrillaComp.HighLight = flexHighlightNever
    grillaValores.Rows = 1
    grillaValores.HighLight = flexHighlightNever
    
    FechaRecibo.Value = Date
    txtCodCliente.Text = ""
    txtNroRecibo.Text = ""
    txtNroSucursal.Text = ""
    FechaRendicion.Value = Date
    cboRecibo.ListIndex = 0
    CboVend.ListIndex = 0
    txtTotalCheques.Text = ""
    txtTotalEfectivo.Text = ""
    txtTotalValores.Text = ""
    txtTotalAplicar.Text = ""
    txtTotalComprobante.Text = ""
    tabValores.Tab = 0
    tabComprobantes.Tab = 0
    
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRecibo) 'ESTADO PENDIENTE
    tabDatos.Tab = 0
    cboRecibo.SetFocus
    
    txtObservaciones.Text = ""
    cmdImprimir.Visible = False
End Sub
Public Function TipoCBT(Codigo As Integer) As String
    Set Rec3 = New ADODB.Recordset
    TipoCBT = ""
    sql = "SELECT TCO_DESCRI, TCO_ABREVIA"
    sql = sql & " FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE"
    sql = sql & " TCO_CODIGO=" & Codigo
    Rec3.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec3.EOF = False Then
        TipoCBT = ChkNull(Rec3!TCO_ABREVIA)
            
    End If
    Rec3.Close
    
End Function

Private Sub cmdNuevoCheque_Click()
    FrmCargaCheques.Show vbModal
    TxtCheNumero.SetFocus
    If TxtCheNumero.Text <> "" Then
        TxtCheNumero_LostFocus
        cmdAgregarCheque_Click
    End If
End Sub

Private Sub cmdQuitarComprobantes_Click()
    If GrillaAplicar1.Rows > 1 Then
        If MsgBox("¿Seguro que desea eliminar?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            If GrillaAplicar1.Rows > 2 Then
                GrillaAplicar1.RemoveItem GrillaAplicar1.RowSel
                txtTotalAplicar.Text = SumaGrilla(GrillaAplicar1, 4)
            Else
                GrillaAplicar1.Rows = 1
                txtTotalAplicar.Text = ""
                GrillaAplicar1.HighLight = flexHighlightNever
            End If
        End If
    End If
End Sub

Private Sub cmdQuitarVal_Click()
    If grillaValores.Rows > 1 Then
        If MsgBox("¿Seguro que desea eliminar?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            If grillaValores.Rows > 2 Then
                If grillaValores.TextMatrix(grillaValores.RowSel, 0) = "A-CTA" Then
                    QuitoDineroACta
                End If
                grillaValores.RemoveItem grillaValores.RowSel
                txtTotalValores.Text = SumaGrilla(grillaValores, 1)
            Else
                If grillaValores.TextMatrix(grillaValores.RowSel, 0) = "A-CTA" Then
                    QuitoDineroACta
                End If
                grillaValores.Rows = 1
                txtTotalValores.Text = ""
                grillaValores.HighLight = flexHighlightNever
            End If
        End If
    End If
End Sub

Private Sub QuitoDineroACta()
    For I = 1 To GrillaAFavor.Rows - 1
        If GrillaAFavor.TextMatrix(I, 5) = grillaValores.TextMatrix(grillaValores.RowSel, 5) _
            And CLng(GrillaAFavor.TextMatrix(I, 1)) = CLng(grillaValores.TextMatrix(grillaValores.RowSel, 4)) _
            And CDate(GrillaAFavor.TextMatrix(I, 2)) = CDate(grillaValores.TextMatrix(grillaValores.RowSel, 2)) Then
            
            'ARREGLO EL SALDO DEL DINERO A CTA
            GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4) = Valido_Importe(CStr(CDbl(GrillaAFavor.TextMatrix(I, 4)) + CDbl(grillaValores.TextMatrix(grillaValores.RowSel, 1))))
           Exit For
        End If
    Next
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmReciboCliente = Nothing
        Unload Me
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub FechaRecibo_LostFocus()
    If IsNull(FechaRecibo.Value) Then
        FechaRecibo.Value = Date
    End If
End Sub

Private Sub FechaRendicion_LostFocus()
    If IsNull(FechaRendicion.Value) Then
        FechaRendicion.Value = Date
    End If
End Sub


Private Sub Form_Activate()
    'recibo automatico
    If frmFacturaCliente.TxtCodigoCli.Text <> "" Then
        txtCodCliente.Text = frmFacturaCliente.TxtCodigoCli.Text
        txtNroSucursal_LostFocus
        txtNroRecibo_LostFocus
        txtCodCliente_LostFocus
        CboVend_LostFocus
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
    Set Rec2 = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
   
    FechaRecibo.Value = Date
    Call Centrar_pantalla(Me)
    tabDatos.Tab = 0
    tabValores.Tab = 0
    tabComprobantes.Tab = 0
    'CONFIGURO GRILLAS
    ConfiguroGrillas
    
    'CARGO COMBO CON LOS TIPOS DE RECIBO
    LlenarComboRecibo
    'CARGO COMBO CON LOS VENDEDORES
    LLenarComboVendedor
    'CARGO COMBO CON LAS PROVINCIAS
    LLenarComboMoneda
    'CARGO COMBO CON LAS FACTURAS
    LlenarComboFactura
    
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRecibo) 'ESTADO PENDIENTE
    Estado = 1
    '------------------------
    frameBanco.Enabled = False
    cmdAgregarCheque.Enabled = False
    cmdAgregarEfectivo.Enabled = False
    FechaRendicion.Value = Date
    txtNroRecibo.Enabled = True
    cmdAgregarFacturas.Enabled = False
    lblEstado.Caption = ""
    
    If txtCodCliente.Text <> "" Then TxtCodigo_LostFocus
End Sub
Private Sub LlenarComboFactura()
    'CARGO NOTAS DE CREDITO (POR AHORA QUEDAN AFUERA)
'    sql = "SELECT * FROM TIPO_COMPROBANTE"
'    sql = sql & " WHERE TCO_DESCRI LIKE 'NOTA DE CRED%'"
'    sql = sql & " ORDER BY TCO_DESCRI"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            cboComprobantes.AddItem rec!TCO_DESCRI
'            cboComprobantes.ItemData(cboComprobantes.NewIndex) = rec!TCO_CODIGO
'            rec.MoveNext
'        Loop
'        cboComprobantes.ListIndex = 0
'    End If
'    rec.Close
    'CARGO COMPRONATES DE RETENCION
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'RETENCION%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboComprobantes.AddItem rec!TCO_DESCRI
            cboComprobantes.ItemData(cboComprobantes.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboComprobantes.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub ConfiguroGrillas()
    'GRILLA CHEQUES
    GrillaCheques.FormatString = "^Bco|^Loc|^Suc|^Cod|^Nro Cheque|^Fec Vto|>Importe|COD INTERNO BANCO|DECRI BANCO"
    GrillaCheques.ColWidth(0) = 500   'BCO
    GrillaCheques.ColWidth(1) = 500   'LOC
    GrillaCheques.ColWidth(2) = 500   'SUC
    GrillaCheques.ColWidth(3) = 700   'COD
    GrillaCheques.ColWidth(4) = 1100  'NRO CHEQUE
    GrillaCheques.ColWidth(5) = 1000  'FEC VTO
    GrillaCheques.ColWidth(6) = 1000  'IMPORTE
    GrillaCheques.ColWidth(7) = 0     'COD INTERNO BANCO
    GrillaCheques.ColWidth(8) = 0     'DESCRI BANCO
    GrillaCheques.Rows = 1
    'GRILLA EFECTIVO
    GrillaEfectivo.FormatString = "Moneda|>Importe|codigo moneda"
    GrillaEfectivo.ColWidth(0) = 1900 'MONEDA
    GrillaEfectivo.ColWidth(1) = 1000 'IMPORTE
    GrillaEfectivo.ColWidth(2) = 0    'CODIGO MONEDA
    GrillaEfectivo.Rows = 1
    'GRILLA Aplicar A
    GrillaAplicar.FormatString = "Comp.|^Número|^Fecha|>Total|>Saldo|codigo comprobante"
    GrillaAplicar.ColWidth(0) = 850 'COMPROBANTE
    GrillaAplicar.ColWidth(1) = 1300 'NUMERO
    GrillaAplicar.ColWidth(2) = 1000 'FECHA
    GrillaAplicar.ColWidth(3) = 1000 'TOTAL
    GrillaAplicar.ColWidth(4) = 1000 'SALDO
    GrillaAplicar.ColWidth(5) = 0    'CODIGO COMPROBANTE
    GrillaAplicar.Rows = 1
    
    'GRILLA BUSQUEDA
    GrdModulos.FormatString = "^Tipo Rec|^Nro Recibo|^Fecha Recibo|Cliente|Vendedor|" _
                              & "TIPO RECIBO|REPRESENTADA"
    GrdModulos.ColWidth(0) = 1000 'TIPO_RECIBO
    GrdModulos.ColWidth(1) = 1400 'NRO RECIBO
    GrdModulos.ColWidth(2) = 1200 'FECHA_RECIBO
    GrdModulos.ColWidth(3) = 3500 'CLIENTE
    GrdModulos.ColWidth(4) = 3500 'VENDEDOR
    GrdModulos.ColWidth(5) = 0    'TIPO RECIBO (TCO_CODIGO)
    GrdModulos.ColWidth(6) = 0    'REPRESENTADA
    GrdModulos.Cols = 7
    GrdModulos.Rows = 1
    'GRILLA VALORES
    grillaValores.FormatString = "|>Importe|^Fecha|Descripción|^Número|codigo"
    grillaValores.ColWidth(0) = 550  'TIPO DE VALOR (CHE,EFT...)
    grillaValores.ColWidth(1) = 1000 'IMPORTE
    grillaValores.ColWidth(2) = 1000 'FECHA
    grillaValores.ColWidth(3) = 2500 'DESCRIPCION
    grillaValores.ColWidth(4) = 1300 'NUMERO
    grillaValores.ColWidth(5) = 0    'CODIGO
    grillaValores.Rows = 1
    'GRILLA APLICAR1
    GrillaAplicar1.FormatString = "Comp.|>Importe|^Número|^Fecha|>Saldo|codigo comprobante|TOTAL"
    GrillaAplicar1.ColWidth(0) = 900 'COMPROBANTE
    GrillaAplicar1.ColWidth(1) = 1000 'IMPORTE
    GrillaAplicar1.ColWidth(2) = 1300 'NUMERO
    GrillaAplicar1.ColWidth(3) = 1000 'FECHA
    GrillaAplicar1.ColWidth(4) = 1000 'IMPORTE
    GrillaAplicar1.ColWidth(5) = 0    'CODIGO COMPROBANTE
    GrillaAplicar1.ColWidth(6) = 1000    'TOTAL
    GrillaAplicar1.Rows = 1
    'GRILLA COMPROBANTES
    GrillaComp.FormatString = "Comp.|^Número|^Fecha|>Importe|codigo comprobante"
    GrillaComp.ColWidth(0) = 1100 'COMPROBANTE
    GrillaComp.ColWidth(1) = 1300 'NUMERO
    GrillaComp.ColWidth(2) = 1000 'FECHA
    GrillaComp.ColWidth(3) = 1000 'IMPORTE
    GrillaComp.ColWidth(4) = 0    'CODIGO COMPROBANTE
    GrillaComp.Rows = 1
    'GRILLA AFAVOR
    GrillaAFavor.FormatString = "Comp.|^Número|^Fecha|>Total|>Saldo|codigo comprobante"
    GrillaAFavor.ColWidth(0) = 850  'COMPROBANTE
    GrillaAFavor.ColWidth(1) = 1300 'NUMERO
    GrillaAFavor.ColWidth(2) = 1000 'FECHA
    GrillaAFavor.ColWidth(3) = 1000 'TOTAL
    GrillaAFavor.ColWidth(4) = 1000 'SALDO
    GrillaAFavor.ColWidth(5) = 0    'CODIGO COMPROBANTE
    GrillaAFavor.Rows = 1
    GrillaAFavor.HighLight = flexHighlightNever
End Sub

Private Sub LlenarComboRecibo()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'RECIB%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboRecibo.AddItem rec!TCO_DESCRI
            cboRecibo.ItemData(cboRecibo.NewIndex) = rec!TCO_CODIGO
            cboRecibo1.AddItem rec!TCO_DESCRI
            cboRecibo1.ItemData(cboRecibo1.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboRecibo.ListIndex = 0
        cboRecibo1.ListIndex = -1
    End If
    rec.Close
End Sub

Private Sub LLenarComboVendedor()
    sql = "SELECT * FROM VENDEDOR ORDER BY VEN_NOMBRE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            CboVend.AddItem rec!VEN_NOMBRE
            CboVend.ItemData(CboVend.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        CboVend.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LLenarComboMoneda()
    sql = "SELECT * FROM MONEDA ORDER BY MON_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboMoneda.AddItem rec!MON_DESCRI
            cboMoneda.ItemData(cboMoneda.NewIndex) = rec!MON_CODIGO
            rec.MoveNext
        Loop
        cboMoneda.ListIndex = 0
    End If
    rec.Close
End Sub

Private Function BuscoUltimoRecibo(TipoRec As Integer) As String
    'ACA BUSCA EL NUMERO DE REMITO SIGUIENTE AL ULTIMO CARGADO
    sql = "SELECT (RECIBO_A) + 1 AS REC_A, (RECIBO_B) + 1 AS REC_B"
    sql = sql & " FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Select Case TipoRec
            Case 10
                BuscoUltimoRecibo = IIf(IsNull(rec!REC_A), 1, rec!REC_A)
            Case 11
                BuscoUltimoRecibo = IIf(IsNull(rec!REC_B), 1, rec!REC_B)
            Case 12
                MsgBox "No hay Recibos del tipo C", vbExclamation, TIT_MSGBOX
                cboRecibo.SetFocus
        End Select
    End If
    rec.Close
End Function

Private Sub GrdModulos_DblClick()
     If GrdModulos.Rows > 1 Then
        CmdNuevo_Click
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 5)), cboRecibo)
        txtNroRecibo.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        FechaRecibo.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
        tabDatos.Tab = 0
        Call BuscarRecibo(GrdModulos.TextMatrix(GrdModulos.RowSel, 5), Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8) _
                        , Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
        ', GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
        cmdImprimir.Visible = True
     End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub GrillaAFavor_DblClick()
    If GrillaAFavor.Rows > 1 Then
        txtSaldoACta.Text = Valido_Importe(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4))
        txtImporteACta.Text = Valido_Importe(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4))
        txtImporteACta.SetFocus
    End If
End Sub

Private Sub GrillaAFavor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GrillaAFavor.Rows > 1 Then
           GrillaAFavor_DblClick
        End If
    End If
End Sub

Private Sub GrillaAplicar_DblClick()
    If GrillaAplicar.Rows > 1 Then
        txtSaldo.Text = Valido_Importe(GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 4))
        txtImporteApagar.Text = Valido_Importe(GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 4))
        txtImporteApagar.SetFocus
    End If
End Sub

Private Sub GrillaAplicar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GrillaAplicar.Rows > 1 Then
           GrillaAplicar_DblClick
        End If
    End If
End Sub

Private Sub GrillaCheques_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If GrillaCheques.Rows > 2 Then
           GrillaCheques.RemoveItem GrillaCheques.RowSel
        Else
           GrillaCheques.Rows = 1
           GrillaCheques.HighLight = flexHighlightNever
           TxtCheNumero.SetFocus
        End If
        txtTotalCheques.Text = SumaGrilla(GrillaCheques, 6)
        txtTotalValores.Text = Valido_Importe(CStr(CDbl(SumaGrilla(GrillaCheques, 6)) + CDbl(SumaGrilla(GrillaEfectivo, 1))))
    End If
End Sub

Private Sub GrillaEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If GrillaEfectivo.Rows > 2 Then
           GrillaEfectivo.RemoveItem GrillaEfectivo.RowSel
        Else
           GrillaEfectivo.Rows = 1
           GrillaEfectivo.HighLight = flexHighlightNever
           cboMoneda.SetFocus
        End If
        txtTotalEfectivo.Text = SumaGrilla(GrillaEfectivo, 1)
        txtTotalValores.Text = Valido_Importe(CStr(CDbl(SumaGrilla(GrillaCheques, 6)) + CDbl(SumaGrilla(GrillaEfectivo, 1))))
    End If
End Sub

Private Sub Label28_Click()
End Sub

Private Sub tabComprobantes_Click(PreviousTab As Integer)
    If tabComprobantes.Tab = 1 Then
        GrillaAplicar.SetFocus
    End If
    If tabComprobantes.Tab = 0 Then
        If Me.tabComprobantes.Visible = True Then cmdAgregarFactura.SetFocus
        If GrillaAplicar1.Rows > 1 Then
           txtTotalAplicar.Text = Valido_Importe(SumaGrilla(GrillaAplicar1, 1))
        End If
    End If
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    'LimpiarBusqueda
    txtCliente.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    txtVendedor.Enabled = False
    cboRecibo1.Enabled = False
    cmdBuscarCli.Enabled = False
    cmdBuscarVen.Enabled = False
    CmdGrabar.Enabled = False
    If Me.Visible = True Then chkCliente.SetFocus
  Else
    CmdGrabar.Enabled = True
  End If
End Sub

Private Sub LimpiarBusqueda()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    txtVendedor.Text = ""
    txtDesVen.Text = ""
    GrdModulos.Rows = 1
    cboRecibo1.ListIndex = -1
    chkCliente.Value = Unchecked
    chkFecha.Value = Unchecked
    chkVendedor.Value = Unchecked
    chkTipoRecibo.Value = Unchecked
End Sub

Private Sub tabValores_Click(PreviousTab As Integer)
    If tabValores.Tab = 0 Then
        If Me.tabValores.Visible = True Then cmdAgregarCHE.SetFocus
    End If
    If tabValores.Tab = 1 Then
        TxtCheNumero.SetFocus
    End If
    If tabValores.Tab = 2 Then
        cboMoneda.SetFocus
    End If
    If tabValores.Tab = 3 Then
        cboComprobantes.SetFocus
    End If
    If tabValores.Tab = 4 Then
        GrillaAFavor.SetFocus
    End If
End Sub

Private Sub TxtBANCO_GotFocus()
    SelecTexto TxtBanco
End Sub

Private Sub TxtBANCO_LostFocus()
    If Len(TxtBanco.Text) < 3 Then TxtBanco.Text = CompletarConCeros(TxtBanco.Text, 3)
End Sub

Private Sub TxtCheNumero_Change()
    If TxtCheNumero.Text = "" Then
        LimpiarCheques
    Else
        frameBanco.Enabled = True
        cmdAgregarCheque.Enabled = True
    End If
End Sub

Private Sub TxtCheNumero_GotFocus()
    SelecTexto TxtCheNumero
End Sub

Private Sub TxtCheNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCheNumero_LostFocus()
    If TxtCheNumero.Text <> "" Then
        If Len(TxtCheNumero.Text) < 8 Then TxtCheNumero.Text = CompletarConCeros(TxtCheNumero.Text, 8)
    'sql = "SELECT * FROM CHEQUE WHERE "
        sql = "SELECT DISTINCT CE.CHE_NUMERO, CH.CHE_IMPORT, CH.CHE_FECVTO, CE.BAN_CODINT, B.BAN_BANCO, B.BAN_LOCALIDAD,"
        sql = sql & " B.BAN_SUCURSAL, B.BAN_CODIGO, B.BAN_NOMCOR,CE.CES_DESCRI,B.BAN_DESCRI"
        sql = sql & " FROM CHEQUE_ESTADOS CE, CHEQUE CH, BANCO B,ESTADO_CHEQUE E"
        sql = sql & " Where "
        sql = sql & " CE.CHE_NUMERO = CH.CHE_NUMERO And "
        sql = sql & " CE.BAN_CODINT = CH.BAN_CODINT And "
        sql = sql & " CH.BAN_CODINT=B.BAN_CODINT  "
        'sql = sql & " CE.ECH_CODIGO= E.ECH_CODIGO AND" '
        'sql = sql & " E.ECH_CODIGO=7" ' 7-entregado
        sql = sql & " AND CH.CHE_NUMERO LIKE '%" & Trim(TxtCheNumero) & "%'"  'CODIGO (1) ES CHEQUE EN CARTERA
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            TxtCheNumero.Text = rec!CHE_NUMERO
            TxtBanco.Text = rec!BAN_BANCO
            txtlocalidad.Text = rec!BAN_LOCALIDAD
            TxtSucursal.Text = rec!BAN_SUCURSAL
            txtcodigo.Text = rec!BAN_CODIGO
            TxtCheImport.Text = rec!che_import
            TxtCheFecVto.Value = rec!CHE_FECVTO
            TxtBanDescri.Text = rec!BAN_DESCRI
            TxtCodInt.Text = rec!BAN_CODINT
        End If
        rec.Close
    End If
    
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
            'txtCliente.SetFocus
        End If
        rec.Close
    End If
    If chkFecha.Value = Unchecked _
        And chkVendedor.Value = Unchecked And chkTipoRecibo.Value = Unchecked _
        And ActiveControl.Name <> "cmdBuscarCli" _
        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
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

Private Sub txtCodCliente_Change()
    If txtCodCliente.Text = "" Then
        txtCliRazSoc.Text = ""
        txtProvincia.Text = ""
        txtCliLocalidad.Text = ""
        txtDomici.Text = ""
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
        Set rec = New ADODB.Recordset
        sql = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, P.PRO_DESCRI, L.LOC_DESCRI, C.PROV_CODIGO"
        sql = sql & " FROM CLIENTE C,  PROVINCIA P, LOCALIDAD L"
        sql = sql & " WHERE"
        sql = sql & " C.CLI_CODIGO=" & XN(txtCodCliente)
        sql = sql & " AND C.PRO_CODIGO=P.PRO_CODIGO"
        sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
        sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtCliRazSoc.Text = rec!CLI_RAZSOC
            txtProvincia.Text = rec!PRO_DESCRI
            txtCliLocalidad.Text = rec!LOC_DESCRI
            txtDomici.Text = IIf(IsNull(rec!CLI_DOMICI), "", rec!CLI_DOMICI)
            If Estado = 1 Then
                If BuscarFactura(txtCodCliente) = False Then
                    MsgBox "El cliente No tiene Facturas pendiente de Pago", vbExclamation, TIT_MSGBOX
                    txtCodCliente.Text = ""
                    txtCodCliente.SetFocus
                Else
                    If Not IsNull(rec!PROV_CODIGO) Then
                        Call BuscarSaldosAFavor(rec!PROV_CODIGO)
                    End If
                End If
            End If
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCliRazSoc.Text = ""
            txtCodCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto txtcodigo
End Sub

Private Sub txtEftImporte_Change()
    If txtEftImporte.Text = "" Then
        cmdAgregarEfectivo.Enabled = False
    Else
        cmdAgregarEfectivo.Enabled = True
    End If
End Sub

Private Sub txtEftImporte_GotFocus()
    SelecTexto txtEftImporte
End Sub

Private Sub txtEftImporte_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtEftImporte, KeyAscii)
End Sub

Private Sub txtEftImporte_LostFocus()
    If txtEftImporte.Text <> "" Then
        txtEftImporte.Text = Valido_Importe(txtEftImporte.Text)
        cmdAgregarEfectivo.Enabled = True
        cmdAgregarEfectivo.SetFocus
    End If
End Sub

Private Sub txtImporteACta_Change()
    If txtSaldoACta.Text <> "" And txtImporteACta.Text <> "" Then
        cmdAgregarACta.Enabled = True
    Else
        cmdAgregarACta.Enabled = False
    End If
End Sub

Private Sub txtImporteACta_GotFocus()
    SelecTexto txtImporteACta
End Sub

Private Sub txtImporteACta_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImporteACta, KeyAscii)
End Sub

Private Sub txtImporteACta_LostFocus()
    If txtSaldoACta.Text <> "" Then
        If txtImporteACta.Text = "" Then
            txtImporteACta.Text = txtSaldoACta.Text
        ElseIf CDbl(txtImporteACta.Text) > CDbl(txtSaldoACta.Text) Then
            MsgBox "Importe mayor al Saldo", vbCritical, TIT_MSGBOX
            txtImporteACta.Text = txtSaldoACta.Text
            txtImporteACta.SetFocus
        End If
        txtImporteACta.Text = Valido_Importe(txtImporteACta)
    End If
End Sub

Private Sub txtImporteApagar_Change()
    If txtSaldo.Text <> "" And txtImporteApagar.Text <> "" Then
        cmdAgregarFacturas.Enabled = True
    Else
        cmdAgregarFacturas.Enabled = False
    End If
End Sub

Private Sub txtImporteApagar_GotFocus()
    SelecTexto txtImporteApagar
End Sub

Private Sub txtImporteApagar_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImporteApagar, KeyAscii)
End Sub

Private Sub txtImporteApagar_LostFocus()
    If txtSaldo.Text <> "" Then
        If txtImporteApagar.Text = "" Then
            txtImporteApagar.Text = txtSaldo.Text
        ElseIf CDbl(txtImporteApagar.Text) > CDbl(txtSaldo.Text) Then
            MsgBox "Importe mayor al Saldo", vbCritical, TIT_MSGBOX
            txtImporteApagar.Text = txtSaldo.Text
            txtImporteApagar.SetFocus
        End If
        txtImporteApagar.Text = Valido_Importe(txtImporteApagar)
    End If
End Sub

Private Sub txtImporteComprobante_GotFocus()
    SelecTexto txtImporteComprobante
End Sub

Private Sub txtImporteComprobante_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImporteComprobante, KeyAscii)
End Sub

Private Sub txtImporteComprobante_LostFocus()
    If txtImporteComprobante.Text <> "" Then txtImporteComprobante = Valido_Importe(txtImporteComprobante.Text)
End Sub

Private Sub TxtLOCALIDAD_GotFocus()
    SelecTexto txtlocalidad
End Sub

Private Sub Txtlocalidad_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtLOCALIDAD_LostFocus()
    If Len(txtlocalidad.Text) < 3 Then txtlocalidad.Text = CompletarConCeros(txtlocalidad.Text, 3)
End Sub

Private Sub txtNroComprobantes_Change()
    If txtNroComprobantes.Text = "" Then
        txtImporteComprobante.Text = ""
        txtImporteComprobante.Enabled = False
        cmdAgregarComprobante.Enabled = False
    Else
        txtImporteComprobante.Enabled = True
        cmdAgregarComprobante.Enabled = True
    End If
End Sub

Private Function BuscarFactura(CodCli As String) As Boolean
        GrillaAplicar.Rows = 1
        Set Rec1 = New ADODB.Recordset
        sql = "SELECT FCL_NUMERO, FCL_SUCURSAL, FCL_FECHA, FCL_TOTAL, FCL_SALDO"
        sql = sql & " ,TCO_CODIGO, TCO_ABREVIA"
        sql = sql & " FROM SALDO_FACTURAS_CLIENTE_V"
        sql = sql & " WHERE "
        sql = sql & " CLI_CODIGO=" & XN(CodCli)
        sql = sql & " ORDER BY FCL_FECHA DESC"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                If Rec1!FCL_SALDO > 0 Then
                    GrillaAplicar.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!FCL_SUCURSAL, "0000") & "-" & Format(Rec1!FCL_NUMERO, "00000000") _
                                    & Chr(9) & Rec1!FCL_FECHA & Chr(9) & Valido_Importe(Rec1!FCL_TOTAL) _
                                    & Chr(9) & Valido_Importe(Rec1!FCL_SALDO) & Chr(9) & Rec1!TCO_CODIGO
                End If
                Rec1.MoveNext
            Loop
        End If
        Rec1.Close
        
        'BUSCAR NOTA DE DEBITO
        sql = "SELECT N.NDC_NUMERO, N.NDC_SUCURSAL, N.NDC_FECHA, N.NDC_TOTAL, N.NDC_SALDO"
        sql = sql & " ,T.TCO_CODIGO, T.TCO_ABREVIA"
        sql = sql & " FROM NOTA_DEBITO_CLIENTE N, TIPO_COMPROBANTE T"
        sql = sql & " WHERE "
        sql = sql & " T.TCO_CODIGO = N.TCO_CODIGO AND"
        sql = sql & " N.CLI_CODIGO=" & XN(CodCli)
        sql = sql & " ORDER BY N.NDC_FECHA DESC"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                If Rec1!NDC_SALDO > 0 Then
                    GrillaAplicar.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!NDC_SUCURSAL, "0000") & "-" & Format(Rec1!NDC_NUMERO, "00000000") _
                                    & Chr(9) & Rec1!NDC_FECHA & Chr(9) & Valido_Importe(Rec1!NDC_TOTAL) _
                                    & Chr(9) & Valido_Importe(Rec1!NDC_SALDO) & Chr(9) & Rec1!TCO_CODIGO
                End If
                Rec1.MoveNext
            Loop
        End If
        Rec1.Close
        
        If GrillaAplicar.Rows > 1 Then
            BuscarFactura = True
        Else
            BuscarFactura = False
        End If
        
        
        
End Function

Private Sub BuscarSaldosAFavor(CodCli As String)
'        GrillaAFavor.Rows = 1
'        Set Rec1 = New ADODB.Recordset
'        sql = "SELECT RS.TCO_CODIGO, RS.REC_NUMERO, RS.REC_SUCURSAL, RS.REC_FECHA, RS.REC_TOTSALDO"
'        sql = sql & " ,RS.REC_SALDO, T.TCO_ABREVIA"
'        sql = sql & " FROM RECIBO_CLIENTE_SALDO RS, RECIBO_CLIENTE R,TIPO_COMPROBANTE T"
'        sql = sql & " WHERE "
'        sql = sql & " RS.TCO_CODIGO=T.TCO_CODIGO"
'        sql = sql & " AND RS.TCO_CODIGO=R.TCO_CODIGO"
'        sql = sql & " AND RS.REC_NUMERO=R.REC_NUMERO"
'        sql = sql & " AND RS.REC_SUCURSAL=R.REC_SUCURSAL"
'        sql = sql & " AND RS.REC_SALDO > 0"
'        sql = sql & " AND R.CLI_CODIGO=" & XN(CodCli)
'        sql = sql & " ORDER BY RS.TCO_CODIGO,RS.REC_SUCURSAL,RS.REC_NUMERO, RS.REC_FECHA"
'        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'        If Rec1.EOF = False Then
'            GrillaAFavor.HighLight = flexHighlightAlways
'            Do While Rec1.EOF = False
'                If Rec1!REC_SALDO > 0 Then
'                    GrillaAFavor.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!REC_SUCURSAL, "0000") & "-" & Format(Rec1!REC_NUMERO, "00000000") _
'                                    & Chr(9) & Rec1!REC_FECHA & Chr(9) & Valido_Importe(Rec1!REC_TOTSALDO) _
'                                    & Chr(9) & Valido_Importe(Rec1!REC_SALDO) & Chr(9) & Rec1!TCO_CODIGO
'                End If
'                Rec1.MoveNext
'            Loop
'        End If
'        Rec1.Close

        GrillaAFavor.Rows = 1
        Set Rec1 = New ADODB.Recordset
        sql = "SELECT FP.*, TC.*"
        sql = sql & " FROM FACTURA_PROVEEDOR FP,PROVEEDOR P, TIPO_COMPROBANTE TC"
        sql = sql & " WHERE "
        sql = sql & " FP.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND FP.PROV_CODIGO=P.PROV_CODIGO"
        sql = sql & " AND P.PROV_CODIGO=" & XN(CodCli)
        sql = sql & " ORDER BY TC.TCO_DESCRI,FP.FPR_NROSUC,FP.FPR_NUMERO, FP.FPR_FECHA"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec1.EOF = False Then
            GrillaAFavor.HighLight = flexHighlightAlways
            Do While Rec1.EOF = False
                If Rec1!FPR_SALDO > 0 Then
                    GrillaAFavor.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!FPR_NROSUC, "0000") & "-" & Format(Rec1!FPR_NUMERO, "00000000") _
                                    & Chr(9) & Rec1!FPR_FECHA & Chr(9) & Valido_Importe(Rec1!FPR_TOTAL) _
                                    & Chr(9) & Valido_Importe(Rec1!FPR_SALDO) & Chr(9) & Rec1!TCO_CODIGO
                End If
                Rec1.MoveNext
            Loop
        End If
        Rec1.Close

End Sub

Private Sub txtNroComprobantes_LostFocus()
    If txtNroComprobantes.Text <> "" Then
        txtNroComprobantes.Text = Format(txtNroComprobantes.Text, "00000000")
        If BuscoComprobanteEnRecibo = False Then
            MsgBox "El comprobante de Retención ya fue cargado a un Recibo", vbInformation, TIT_MSGBOX
            txtNroComprobantes.Text = ""
            txtNroComprobantes.SetFocus
        End If
    End If
End Sub

Private Function BuscoComprobanteEnRecibo() As Boolean
    Set Rec2 = New ADODB.Recordset
    
    sql = "SELECT DR.REC_NUMERO"
    sql = sql & " FROM DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE RC"
    sql = sql & " WHERE"
    sql = sql & " DR.DRE_TCO_CODIGO=" & XN(cboComprobantes.ItemData(cboComprobantes.ListIndex))
    sql = sql & " AND DR.DRE_COMNUMERO=" & XN(txtNroComprobantes)
    sql = sql & " AND DR.DRE_COMSUCURSAL=" & XN(txtNroCompSuc)
    sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCodCliente)
    sql = sql & " AND DR.REC_NUMERO=RC.REC_NUMERO"
    sql = sql & " AND DR.REC_SUCURSAL=RC.REC_SUCURSAL"
    sql = sql & " AND DR.TCO_CODIGO=RC.TCO_CODIGO"
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Rec2.EOF = False Then
        BuscoComprobanteEnRecibo = False
    Else
        BuscoComprobanteEnRecibo = True
    End If
    Rec2.Close
    
End Function

Private Sub txtNroCompSuc_GotFocus()
    SelecTexto txtNroCompSuc
End Sub

Private Sub txtNroCompSuc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroCompSuc_LostFocus()
    If txtNroCompSuc.Text <> "" Then
        txtNroCompSuc.Text = Format(txtNroCompSuc.Text, "0000")
    End If
End Sub

'Private Sub BuscarNotaCredito()
'
'    Set rec = New ADODB.Recordset
'
'    sql = "SELECT NCC_FECHA, NCC_TOTAL"
'    sql = sql & " FROM NOTA_CREDITO_CLIENTE"
'    sql = sql & " WHERE"
'    sql = sql & " TCO_CODIGO=" & cboComprobantes.ItemData(cboComprobantes.ListIndex)
'    If FechaComprobantes.value <> "" Then
'        sql = sql & " AND NCC_FECHA=" & XDQ(FechaComprobantes.value)
'    End If
'    sql = sql & " AND NCC_NUMERO=" & XN(txtNroComprobantes)
'    sql = sql & " AND CLI_CODIGO=" & XN(txtCodCliente)
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        If BuscoComprobanteEnRecibo = False Then
'            MsgBox "La Nota de Crédito ya fue cargada a un Recibo", vbInformation, TIT_MSGBOX
'            txtNroComprobantes.Text = ""
'            txtNroComprobantes.SetFocus
'        Else
'            FechaComprobantes.value = rec!NCC_FECHA
'            txtImporteComprobante.Text = Valido_Importe(rec!NCC_TOTAL)
'            txtImporteComprobante.SetFocus
'        End If
'    Else
'        MsgBox "La Nota de Crédito no pertenece al Cliente seleccionado", vbExclamation, TIT_MSGBOX
'        txtNroComprobantes.Text = ""
'        cboComprobantes.SetFocus
'    End If
'    rec.Close
'End Sub

Private Sub txtNroRecibo_Change()
    If txtNroRecibo.Text = "" Then
        FechaRecibo.Value = Date
    End If
End Sub

Private Sub txtNroRecibo_GotFocus()
    SelecTexto txtNroRecibo
End Sub

Private Sub txtNroRecibo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroRecibo_LostFocus()
    If txtNroRecibo.Text = "" Then
        'BUSCO EL NUMERO DE RECIBO QUE CORRESPONDE
        txtNroRecibo.Text = Format(BuscoUltimoRecibo(cboRecibo.ItemData(cboRecibo.ListIndex)), "00000000")
    Else
        If txtNroSucursal.Text = "" Then
            txtNroSucursal.Text = Sucursal
        End If
        txtNroRecibo.Text = Format(txtNroRecibo.Text, "00000000")
        Call BuscarRecibo(CStr(cboRecibo.ItemData(cboRecibo.ListIndex)), _
                          txtNroRecibo, txtNroSucursal)
    End If
End Sub

Private Sub BuscarRecibo(TipoRec As String, NroRec As String, NroSuc As String)
    Set Rec2 = New ADODB.Recordset
    sql = "SELECT * FROM RECIBO_CLIENTE"
    sql = sql & " WHERE"
    sql = sql & " TCO_CODIGO=" & XN(TipoRec)
    sql = sql & " AND REC_NUMERO=" & XN(NroRec)
    sql = sql & " AND REC_SUCURSAL=" & XN(NroSuc)
    
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        If Rec2.RecordCount > 2 Then
            Rec2.Close
            tabDatos.Tab = 1
            Exit Sub
        End If
        'CABEZA DEL RECIDO
        FechaRecibo.Value = Rec2!REC_FECHA
        FechaRendicion.Value = Rec2!REC_FECHA_RENDICION
        'CARGO ESTADO
        Call BuscoEstado(CInt(Rec2!EST_CODIGO), lblEstadoRecibo)
        Estado = CInt(Rec2!EST_CODIGO)
        txtCodCliente.Text = Rec2!CLI_CODIGO
        txtCodCliente_LostFocus
        'BUSCA VENDEDOR
        Call BuscaCodigoProxItemData(CInt(Rec2!VEN_CODIGO), CboVend)
        
        txtObservaciones.Text = IIf(IsNull(Rec2!REC_OBSER), "", Rec2!REC_OBSER)
        'DETALLE_DEL RECIBO CHEQUES
        Set rec = New ADODB.Recordset
        sql = "SELECT *"
        sql = sql & " FROM DETALLE_RECIBO_CLIENTE"
        sql = sql & " WHERE"
        sql = sql & " TCO_CODIGO=" & XN(TipoRec)
        sql = sql & " AND REC_NUMERO=" & XN(NroRec)
        sql = sql & " AND REC_SUCURSAL=" & XN(NroSuc)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            Do While rec.EOF = False
                If Not IsNull(rec!BAN_CODINT) Then 'BANCO
                    Call BuscarCheque(rec!BAN_CODINT, rec!CHE_NUMERO)
                ElseIf Not IsNull(rec!MON_CODIGO) Then 'MONEDA
                    grillaValores.AddItem "EFT" & Chr(9) & Valido_Importe(rec!DRE_MONIMP) _
                                    & Chr(9) & "" & Chr(9) & BuscarMoneda(rec!MON_CODIGO) _
                                    & Chr(9) & "" & Chr(9) & rec!MON_CODIGO
                                    
                ElseIf Not IsNull(rec!DRE_TCO_CODIGO) Then 'COMPROBANTE
                    Dim QueEs As String
                    If rec!DRE_TCO_CODIGO >= 1 And rec!DRE_TCO_CODIGO <= 13 Then
                        'QueEs = "A-CTA"
                        QueEs = "FAC"
                    Else
                        QueEs = "COMP"
                    End If
                    grillaValores.AddItem QueEs & Chr(9) & Valido_Importe(IIf(IsNull(rec!DRE_COMIMP), "", rec!DRE_COMIMP)) _
                                    & Chr(9) & IIf(IsNull(rec!DRE_COMFECHA), "", rec!DRE_COMFECHA) & Chr(9) & IIf(QueEs = "FAC", "FACTURA A", TipoCBT(rec!DRE_TCO_CODIGO)) _
                                    & Chr(9) & Format(IIf(IsNull(rec!DRE_COMSUCURSAL), "", rec!DRE_COMSUCURSAL), "0000") & "-" & Format(rec!DRE_COMNUMERO, "00000000") _
                                    & Chr(9) & rec!DRE_TCO_CODIGO
                                    ' EN "" ESTABA BuscarTipoDocAbre(rec!DRE_TCO_CODIGO)
                End If
                If Not IsNull(rec!DRE_VUELTO) Then
                    grillaValores.AddItem "VUE" & Chr(9) & Valido_Importe(rec!DRE_VUELTO) _
                                    & Chr(9) & "" & Chr(9) & "VUELTO EN PESOS"
                End If
                rec.MoveNext
            Loop
            
            grillaValores.HighLight = flexHighlightAlways
            txtTotalValores.Text = SumaGrilla(grillaValores, 1)
        End If
        rec.Close
                   
        'DETALLE_DEL RECIBO FACTURA
        sql = "SELECT *"
        sql = sql & " FROM FACTURAS_RECIBO_CLIENTE"
        sql = sql & " WHERE"
        sql = sql & " TCO_CODIGO=" & XN(TipoRec)
        sql = sql & " AND REC_NUMERO=" & XN(NroRec)
        sql = sql & " AND REC_SUCURSAL=" & XN(NroSuc)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrillaAplicar1.AddItem BuscarTipoDocAbre(rec!FCL_TCO_CODIGO) & Chr(9) & Valido_Importe(rec!REC_IMPORTE) & Chr(9) & _
                            Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") & Chr(9) & rec!FCL_FECHA _
                            & Chr(9) & "" & Chr(9) & rec!FCL_TCO_CODIGO
                            
                rec.MoveNext
            Loop
            GrillaAplicar1.HighLight = flexHighlightAlways
            txtTotalAplicar.Text = SumaGrilla(GrillaAplicar1, 1)
        End If
        FrameRecibo.Enabled = False
        FrameRemito.Enabled = False
        rec.Close
        CmdNuevo.SetFocus
        CmdGrabar.Enabled = False
    End If
    Rec2.Close
End Sub

Private Function BuscarCheque(Codigo As String, NroChe As String) As String
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT B.BAN_DESCRI,C.CHE_IMPORT,C.CHE_FECVTO"
    sql = sql & " FROM BANCO B, CHEQUE C"
    sql = sql & " WHERE C.BAN_CODINT=" & XN(Codigo)
    sql = sql & " AND C.CHE_NUMERO=" & XS(NroChe)
    sql = sql & " AND C.BAN_CODINT=B.BAN_CODINT"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        grillaValores.AddItem "CHE" & Chr(9) & Valido_Importe(Rec1!che_import) & Chr(9) & Rec1!CHE_FECVTO _
                           & Chr(9) & Rec1!BAN_DESCRI & Chr(9) & NroChe & Chr(9) & Codigo
    End If
    Rec1.Close
End Function

Private Function BuscarMoneda(Codigo As String) As String
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT MON_DESCRI"
    sql = sql & " FROM MONEDA"
    sql = sql & " WHERE MON_CODIGO=" & XN(Codigo)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscarMoneda = Rec1!MON_DESCRI
    Else
        BuscarMoneda = ""
    End If
    Rec1.Close
End Function

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

Private Sub txtSucursal_GotFocus()
    SelecTexto TxtSucursal
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtBanco_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
      KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    Dim MtrObjetos As Variant
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
        
    'ChequeRegistrado = False
    
    If Len(txtcodigo.Text) < 6 Then txtcodigo.Text = CompletarConCeros(txtcodigo.Text, 6)
     
    If Trim(Me.TxtCheNumero.Text) <> "" And _
       Trim(Me.TxtBanco.Text) <> "" And _
       Trim(Me.txtlocalidad.Text) <> "" And _
       Trim(Me.TxtSucursal.Text) <> "" And _
       Trim(Me.txtcodigo.Text) <> "" Then
       
       'BUSCO EL CODIGO INTERNO
       sql = "SELECT BAN_CODINT, BAN_DESCRI"
       sql = sql & " FROM BANCO"
       sql = sql & " WHERE BAN_BANCO = " & XS(TxtBanco.Text)
       sql = sql & " AND BAN_LOCALIDAD = " & XS(Me.txtlocalidad.Text)
       sql = sql & " AND BAN_SUCURSAL = " & XS(Me.TxtSucursal.Text)
       sql = sql & " AND BAN_CODIGO = " & XS(txtcodigo.Text)
       rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
       If rec.RecordCount > 0 Then 'EXITE
          TxtCodInt.Text = rec!BAN_CODINT
          TxtBanDescri.Text = rec!BAN_DESCRI
          rec.Close
       Else
          If Me.ActiveControl.Name <> "CmdSalir" And Me.ActiveControl.Name <> "CmdNuevo" Then
            MsgBox "Banco NO Registrado.", 16, TIT_MSGBOX
            TxtBanco.SetFocus
            Me.CmdBanco.SetFocus
          End If
          rec.Close
          Exit Sub
       End If
       
       'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE " & _
              " WHERE CHE_NUMERO = " & XS(TxtCheNumero.Text) & _
                " AND BAN_CODINT = " & XN(TxtCodInt.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then 'EXITE
            Me.TxtCheFecVto.Value = rec!CHE_FECVTO
            Me.TxtCheImport.Text = Valido_Importe(rec!che_import)
            
            'BUSCO LOS ATRIBUTOS DE BANCO
            sql = "SELECT BAN_BANCO,BAN_LOCALIDAD,BAN_SUCURSAL,BAN_CODIGO FROM BANCO " & _
                   "WHERE BAN_CODINT = " & XN(Me.TxtCodInt.Text)
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.RecordCount > 0 Then 'EXITE
                Me.TxtBanco.Text = Rec1!BAN_BANCO
                Me.txtlocalidad.Text = Rec1!BAN_LOCALIDAD
                Me.TxtSucursal.Text = Rec1!BAN_SUCURSAL
                Me.txtcodigo.Text = Rec1!BAN_CODIGO
            End If
            Rec1.Close
        Else
           MsgBox "El Cheque no fue registrado, el mismo debe ser registrado con anterioridad", vbInformation, TIT_MSGBOX
           rec.Close
           cmdNuevoCheque_Click
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtSucursal_LostFocus()
    If Len(TxtSucursal.Text) < 3 Then TxtSucursal.Text = CompletarConCeros(TxtSucursal.Text, 3)
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
    If chkFecha.Value = Unchecked And chkTipoRecibo.Value = Unchecked _
    And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub

