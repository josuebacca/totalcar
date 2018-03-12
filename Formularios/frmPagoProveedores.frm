VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPagoProveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Pago"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   9135
      TabIndex        =   135
      Top             =   6345
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   10020
      TabIndex        =   49
      Top             =   6345
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   8250
      TabIndex        =   48
      Top             =   6345
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10890
      TabIndex        =   50
      Top             =   6345
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6255
      Left            =   60
      TabIndex        =   67
      Top             =   45
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   11033
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
      TabPicture(0)   =   "frmPagoProveedores.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameRecibo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tabValores"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tabComprobantes"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameProveedor"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmPagoProveedores.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame FrameProveedor 
         Caption         =   "Proveedor..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   2925
         TabIndex        =   121
         Top             =   375
         Width           =   8655
         Begin VB.TextBox txtcodcli 
            Height          =   285
            Left            =   120
            TabIndex        =   149
            Top             =   720
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscarProveedor1 
            Height          =   300
            Left            =   2295
            MaskColor       =   &H000000FF&
            Picture         =   "frmPagoProveedores.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "Buscar Proveedor"
            Top             =   660
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboTipoProveedor 
            Height          =   315
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   285
            Width           =   3375
         End
         Begin VB.TextBox txtCodProveedor 
            Height          =   300
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   3
            Top             =   645
            Width           =   975
         End
         Begin VB.TextBox txtProvRazSoc 
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
            Left            =   2730
            MaxLength       =   50
            TabIndex        =   4
            Tag             =   "Descripción"
            Top             =   645
            Width           =   5295
         End
         Begin VB.TextBox txtCliLocalidad 
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
            Height          =   285
            Left            =   1275
            TabIndex        =   123
            Top             =   990
            Width           =   4860
         End
         Begin VB.TextBox txtDomici 
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
            Height          =   285
            Left            =   1275
            MaxLength       =   50
            TabIndex        =   122
            Top             =   1305
            Width           =   4860
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Prov.:"
            Height          =   195
            Left            =   405
            TabIndex        =   127
            Top             =   315
            Width           =   780
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
            Left            =   645
            TabIndex        =   126
            Top             =   660
            Width           =   540
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Loc.:"
            Height          =   180
            Left            =   825
            TabIndex        =   125
            Top             =   1035
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Dom.:"
            Height          =   195
            Left            =   765
            TabIndex        =   124
            Top             =   1335
            Width           =   420
         End
      End
      Begin TabDlg.SSTab tabComprobantes 
         Height          =   3975
         Left            =   120
         TabIndex        =   82
         Top             =   2175
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
         TabPicture(0)   =   "frmPagoProveedores.frx":0342
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame1"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "C&omprobantes Pendientes"
         TabPicture(1)   =   "frmPagoProveedores.frx":035E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame5"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame5 
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
            TabIndex        =   108
            Top             =   480
            Width           =   5535
            Begin VB.CommandButton cmdAceptarComprobantes 
               Caption         =   "A&ceptar"
               Height          =   360
               Left            =   4485
               TabIndex        =   9
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
               TabIndex        =   7
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
               TabIndex        =   6
               Top             =   2625
               Width           =   1185
            End
            Begin VB.CommandButton cmdAgregarFacturas 
               Caption         =   "A&gregar"
               Height          =   360
               Left            =   3570
               TabIndex        =   8
               Top             =   2985
               Width           =   900
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaAplicar 
               Height          =   2205
               Left            =   60
               TabIndex        =   5
               Top             =   255
               Width           =   5400
               _ExtentX        =   9525
               _ExtentY        =   3889
               _Version        =   393216
               Cols            =   8
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
               TabIndex        =   110
               Top             =   3045
               Width           =   1155
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Saldo:"
               Height          =   195
               Left            =   900
               TabIndex        =   109
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
            TabIndex        =   104
            Top             =   480
            Width           =   5565
            Begin VB.CommandButton cmdAceptarFacturas 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   4515
               TabIndex        =   12
               Top             =   525
               Width           =   945
            End
            Begin VB.CommandButton cmdAgregarFactura 
               Caption         =   "Agregar Com"
               Height          =   360
               Left            =   2400
               TabIndex        =   10
               Top             =   525
               Width           =   1140
            End
            Begin VB.CommandButton cmdQuitarComprobantes 
               Caption         =   "Quitar"
               Height          =   360
               Left            =   3555
               TabIndex        =   11
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
               TabIndex        =   105
               Top             =   2880
               Width           =   1170
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaAplicar1 
               Height          =   1860
               Left            =   75
               TabIndex        =   106
               Top             =   915
               Width           =   5430
               _ExtentX        =   9578
               _ExtentY        =   3281
               _Version        =   393216
               Cols            =   8
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
               TabIndex        =   111
               Top             =   2925
               Width           =   405
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Total Valores Recibidos:"
               Height          =   195
               Left            =   360
               TabIndex        =   107
               Top             =   3420
               Width           =   1725
            End
         End
      End
      Begin TabDlg.SSTab tabValores 
         Height          =   3975
         Left            =   5865
         TabIndex        =   51
         Top             =   2175
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   7011
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         Tab             =   4
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
         TabPicture(0)   =   "frmPagoProveedores.frx":037A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Cheques"
         TabPicture(1)   =   "frmPagoProveedores.frx":0396
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&Moneda"
         TabPicture(2)   =   "frmPagoProveedores.frx":03B2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame4"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "&Comprobantes"
         TabPicture(3)   =   "frmPagoProveedores.frx":03CE
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame7"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Va&lores a Cuenta"
         TabPicture(4)   =   "frmPagoProveedores.frx":03EA
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "Frame6"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         Begin VB.Frame Frame6 
            Caption         =   "Valores a Cuenta"
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
            TabIndex        =   130
            Top             =   480
            Width           =   5535
            Begin VB.CommandButton cmaAceptarACta 
               Caption         =   "A&ceptar"
               Height          =   360
               Left            =   4500
               TabIndex        =   47
               Top             =   2970
               Width           =   900
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
               TabIndex        =   45
               Top             =   2985
               Width           =   1185
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
               TabIndex        =   44
               Top             =   2610
               Width           =   1185
            End
            Begin VB.CommandButton cmdAgregarACta 
               Caption         =   "A&gregar"
               Height          =   360
               Left            =   3585
               TabIndex        =   46
               Top             =   2970
               Width           =   900
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaAFavor 
               Height          =   2205
               Left            =   90
               TabIndex        =   43
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
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Left            =   795
               TabIndex        =   132
               Top             =   3030
               Width           =   570
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Saldo:"
               Height          =   195
               Left            =   915
               TabIndex        =   131
               Top             =   2655
               Width           =   450
            End
         End
         Begin VB.Frame Frame7 
            Height          =   3405
            Left            =   -74940
            TabIndex        =   113
            Top             =   480
            Width           =   5535
            Begin VB.CommandButton cmdCancelarComprobante 
               Caption         =   "Cancelar"
               Height          =   360
               Left            =   1305
               TabIndex        =   42
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
            Begin VB.TextBox txtImporteComprobante 
               Height          =   315
               Left            =   3165
               MaxLength       =   8
               TabIndex        =   39
               Top             =   990
               Width           =   1140
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
               Left            =   1335
               MaxLength       =   8
               TabIndex        =   38
               Top             =   990
               Width           =   1140
            End
            Begin VB.ComboBox cboComprobantes 
               Height          =   315
               Left            =   1335
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   300
               Width           =   2970
            End
            Begin VB.CommandButton cmdAgregarComprobante 
               Caption         =   "Agregar"
               Height          =   345
               Left            =   4485
               TabIndex        =   40
               Top             =   990
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
               TabIndex        =   114
               Top             =   2940
               Width           =   1035
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaComp 
               Height          =   1455
               Left            =   315
               TabIndex        =   115
               Top             =   1410
               Width           =   4950
               _ExtentX        =   8731
               _ExtentY        =   2566
               _Version        =   393216
               Cols            =   5
               FixedCols       =   0
               BackColorSel    =   8388736
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
            End
            Begin MSComCtl2.DTPicker fechaComprobantes 
               Height          =   315
               Left            =   1320
               TabIndex        =   37
               Top             =   645
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   57278465
               CurrentDate     =   41098
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Left            =   2550
               TabIndex        =   120
               Top             =   1050
               Width           =   570
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Comprobante:"
               Height          =   195
               Left            =   300
               TabIndex        =   119
               Top             =   1050
               Width           =   990
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   930
               TabIndex        =   118
               Top             =   330
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fecha:"
               Height          =   195
               Index           =   3
               Left            =   795
               TabIndex        =   117
               Top             =   705
               Width           =   495
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   3525
               TabIndex        =   116
               Top             =   3000
               Width           =   405
            End
         End
         Begin VB.Frame Frame4 
            Height          =   3390
            Left            =   -74940
            TabIndex        =   98
            Top             =   480
            Width           =   5535
            Begin VB.CommandButton cmdCancelarMoneda 
               Caption         =   "Cancelar"
               Height          =   360
               Left            =   2115
               TabIndex        =   35
               Top             =   2940
               Width           =   960
            End
            Begin VB.CommandButton cmdAceptarMoneda 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   1140
               TabIndex        =   34
               Top             =   2940
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
               TabIndex        =   102
               Top             =   2505
               Width           =   1035
            End
            Begin VB.ComboBox cboMoneda 
               Height          =   315
               Left            =   1125
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   495
               Width           =   1950
            End
            Begin VB.TextBox txtEftImporte 
               Height          =   315
               Left            =   1125
               TabIndex        =   32
               Top             =   930
               Width           =   1005
            End
            Begin VB.CommandButton cmdAgregarEfectivo 
               Caption         =   "Agregar"
               Height          =   345
               Left            =   2370
               TabIndex        =   33
               Top             =   915
               Width           =   720
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaEfectivo 
               Height          =   1095
               Left            =   1095
               TabIndex        =   99
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
               TabIndex        =   103
               Top             =   2565
               Width           =   405
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Moneda:"
               Height          =   195
               Left            =   480
               TabIndex        =   101
               Top             =   525
               Width           =   630
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Index           =   2
               Left            =   540
               TabIndex        =   100
               Top             =   975
               Width           =   570
            End
         End
         Begin VB.Frame Frame3 
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
            TabIndex        =   83
            Top             =   465
            Width           =   5535
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
               Left            =   120
               TabIndex        =   88
               Top             =   720
               Width           =   4635
               Begin VB.TextBox TxtSUCURSAL 
                  Height          =   285
                  Left            =   2280
                  MaxLength       =   3
                  TabIndex        =   24
                  Top             =   255
                  Width           =   450
               End
               Begin VB.TextBox TxtBANCO 
                  Height          =   285
                  Left            =   525
                  MaxLength       =   3
                  TabIndex        =   22
                  Top             =   240
                  Width           =   450
               End
               Begin VB.TextBox TxtLOCALIDAD 
                  Height          =   285
                  Left            =   1410
                  MaxLength       =   3
                  TabIndex        =   23
                  Top             =   240
                  Width           =   450
               End
               Begin VB.TextBox TxtCODIGO 
                  Height          =   285
                  Left            =   3360
                  MaxLength       =   6
                  TabIndex        =   25
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
                  TabIndex        =   91
                  Top             =   615
                  Width           =   4500
               End
               Begin VB.TextBox TxtCodInt 
                  BackColor       =   &H80000018&
                  Height          =   300
                  Left            =   2745
                  TabIndex        =   90
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   420
               End
               Begin VB.CommandButton CmdBanco 
                  DisabledPicture =   "frmPagoProveedores.frx":0406
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
                  Picture         =   "frmPagoProveedores.frx":0710
                  Style           =   1  'Graphical
                  TabIndex        =   89
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
                  TabIndex        =   95
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
                  TabIndex        =   94
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
                  TabIndex        =   93
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
                  TabIndex        =   92
                  Top             =   285
                  Width           =   540
               End
            End
            Begin VB.ComboBox cboBanco 
               Height          =   315
               Left            =   1110
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   825
               Width           =   4110
            End
            Begin VB.ComboBox cboCtaBancaria 
               Height          =   315
               Left            =   1110
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   1185
               Width           =   1845
            End
            Begin VB.OptionButton optChequePropio 
               Caption         =   "Cheques Propios"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2820
               TabIndex        =   20
               Top             =   45
               Width           =   1755
            End
            Begin VB.OptionButton optChequeTercero 
               Caption         =   "Cheques de Terceros"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   405
               TabIndex        =   19
               Top             =   30
               Width           =   2145
            End
            Begin VB.CommandButton cmdBuscarCheques 
               Height          =   315
               Left            =   2535
               MaskColor       =   &H000000FF&
               Picture         =   "frmPagoProveedores.frx":085A
               Style           =   1  'Graphical
               TabIndex        =   137
               ToolTipText     =   "Buscar Cheques en Cartera"
               Top             =   330
               UseMaskColor    =   -1  'True
               Width           =   405
            End
            Begin VB.CommandButton cmdCancelarCheques 
               Caption         =   "Cancelar"
               Height          =   360
               Left            =   1575
               TabIndex        =   30
               Top             =   3015
               Width           =   960
            End
            Begin VB.CommandButton cmdAceptarCheques 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   585
               TabIndex        =   29
               Top             =   3015
               Width           =   960
            End
            Begin VB.TextBox TxtCheNumero 
               Height          =   315
               Left            =   1110
               MaxLength       =   8
               TabIndex        =   21
               Top             =   330
               Width           =   1380
            End
            Begin VB.CommandButton cmdNuevoCheque 
               Height          =   315
               Left            =   2970
               MaskColor       =   &H000000FF&
               Picture         =   "frmPagoProveedores.frx":0B64
               Style           =   1  'Graphical
               TabIndex        =   87
               ToolTipText     =   "Cargar Cheques"
               Top             =   330
               UseMaskColor    =   -1  'True
               Width           =   405
            End
            Begin VB.CommandButton cmdAgregarCheque 
               Caption         =   "Agregar"
               Height          =   345
               Left            =   4755
               TabIndex        =   28
               Top             =   1425
               Width           =   705
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
               TabIndex        =   86
               Top             =   3045
               Width           =   1035
            End
            Begin VB.TextBox TxtCheImport 
               Height          =   330
               Left            =   3540
               TabIndex        =   84
               Top             =   315
               Width           =   900
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaCheques 
               Height          =   1170
               Left            =   75
               TabIndex        =   85
               Top             =   1815
               Width           =   5385
               _ExtentX        =   9499
               _ExtentY        =   2064
               _Version        =   393216
               Cols            =   10
               FixedCols       =   0
               BackColorSel    =   8388736
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
            End
            Begin MSComCtl2.DTPicker TxtCheFecVto 
               Height          =   315
               Left            =   3840
               TabIndex        =   148
               Top             =   360
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   57278465
               CurrentDate     =   41098
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Banco:"
               Height          =   195
               Index           =   1
               Left            =   510
               TabIndex        =   139
               Top             =   870
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nro Cuenta:"
               Height          =   195
               Index           =   4
               Left            =   165
               TabIndex        =   138
               Top             =   1215
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nro Cheque:"
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   97
               Top             =   375
               Width           =   900
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   3840
               TabIndex        =   96
               Top             =   3105
               Width           =   405
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Valores Entregados..."
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
            TabIndex        =   78
            Top             =   480
            Width           =   5565
            Begin VB.CommandButton cmdAgregarCHE 
               Caption         =   "Agregar Che"
               Height          =   360
               Left            =   150
               TabIndex        =   13
               Top             =   540
               Width           =   1065
            End
            Begin VB.CommandButton cmdQuitarVal 
               Caption         =   "&Quitar"
               Height          =   360
               Left            =   4770
               TabIndex        =   18
               Top             =   555
               Width           =   705
            End
            Begin VB.CommandButton cmdAgregarCOMP 
               Caption         =   "Agregar Com"
               Height          =   360
               Left            =   2310
               TabIndex        =   15
               Top             =   540
               Width           =   1065
            End
            Begin VB.CommandButton cmdAgregarEFT 
               Caption         =   "Agregar Eft"
               Height          =   360
               Left            =   1230
               TabIndex        =   14
               Top             =   540
               Width           =   1065
            End
            Begin VB.CommandButton cmdAceptarValores 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   4770
               TabIndex        =   17
               Top             =   195
               Width           =   705
            End
            Begin VB.CommandButton cmdAgregarVALCTA 
               Caption         =   "Agregar Val"
               Height          =   360
               Left            =   3390
               TabIndex        =   16
               Top             =   540
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
               TabIndex        =   79
               Top             =   2895
               Width           =   1170
            End
            Begin MSFlexGridLib.MSFlexGrid grillaValores 
               Height          =   1860
               Left            =   75
               TabIndex        =   80
               Top             =   915
               Width           =   5415
               _ExtentX        =   9551
               _ExtentY        =   3281
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               RowHeightMin    =   250
               BackColorSel    =   8388736
               FocusRect       =   0
               HighLight       =   0
               SelectionMode   =   1
               AllowUserResizing=   3
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   360
               TabIndex        =   112
               Top             =   2940
               Width           =   405
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Total Valores Recibidos:"
               Height          =   195
               Left            =   360
               TabIndex        =   81
               Top             =   3420
               Width           =   1725
            End
         End
      End
      Begin VB.Frame FrameRecibo 
         Caption         =   "Recibo de Prveedor..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   105
         TabIndex        =   71
         Top             =   375
         Width           =   2805
         Begin VB.ComboBox cboOrdPag 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   133
            Top             =   300
            Width           =   1680
         End
         Begin VB.TextBox txtNroOrdenPago 
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
            Left            =   840
            MaxLength       =   8
            TabIndex        =   0
            Top             =   660
            Width           =   1140
         End
         Begin MSComCtl2.DTPicker FechaOrdenPago 
            Height          =   315
            Left            =   840
            TabIndex        =   1
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   57278465
            CurrentDate     =   41098
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   420
            TabIndex        =   134
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lblEstadoRecibo 
            AutoSize        =   -1  'True
            Caption         =   "EST. ORD PAGO"
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
            Left            =   840
            TabIndex        =   75
            Top             =   1410
            Width           =   1470
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   240
            TabIndex        =   74
            Top             =   1395
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   180
            TabIndex        =   73
            Top             =   705
            Width           =   600
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   285
            TabIndex        =   72
            Top             =   1050
            Width           =   495
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
         Height          =   2340
         Left            =   -74835
         TabIndex        =   68
         Top             =   420
         Width           =   11355
         Begin VB.TextBox txtImpChe2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4080
            TabIndex        =   62
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox txtimpChe1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4080
            TabIndex        =   60
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   9960
            TabIndex        =   145
            Top             =   1200
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox txtBBanco1 
            BackColor       =   &H8000000F&
            Height          =   330
            Left            =   5640
            TabIndex        =   144
            Top             =   1200
            Width           =   4860
         End
         Begin VB.TextBox txtchepro 
            Height          =   315
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   59
            Top             =   1215
            Width           =   1380
         End
         Begin VB.CommandButton cmdBuscaPCheque 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3600
            MaskColor       =   &H000000FF&
            Picture         =   "frmPagoProveedores.frx":0EEE
            Style           =   1  'Graphical
            TabIndex        =   143
            ToolTipText     =   "Buscar Cheques"
            Top             =   1215
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   9960
            TabIndex        =   142
            Top             =   1560
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox txtBBanco2 
            BackColor       =   &H8000000F&
            Height          =   330
            Left            =   5640
            TabIndex        =   141
            Top             =   1560
            Width           =   4860
         End
         Begin VB.TextBox txtcheqTer 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   61
            Top             =   1575
            Width           =   1380
         End
         Begin VB.CommandButton cmdBuscaCheque 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3600
            MaskColor       =   &H000000FF&
            Picture         =   "frmPagoProveedores.frx":11F8
            Style           =   1  'Graphical
            TabIndex        =   140
            ToolTipText     =   "Buscar Cheques"
            Top             =   1575
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CheckBox chkTercero 
            Caption         =   "Cheque Terceros"
            Height          =   195
            Left            =   390
            TabIndex        =   55
            Top             =   1581
            Width           =   1575
         End
         Begin VB.CheckBox chkPropio 
            Caption         =   "Cheque Propio"
            Height          =   195
            Left            =   390
            TabIndex        =   54
            Top             =   1229
            Width           =   1455
         End
         Begin VB.CommandButton cmdBuscarProveedor 
            Height          =   300
            Left            =   3165
            MaskColor       =   &H000000FF&
            Picture         =   "frmPagoProveedores.frx":1502
            Style           =   1  'Graphical
            TabIndex        =   129
            ToolTipText     =   "Buscar Proveedor"
            Top             =   780
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtProveedor 
            Height          =   300
            Left            =   2160
            MaxLength       =   40
            TabIndex        =   58
            Top             =   780
            Width           =   975
         End
         Begin VB.TextBox txtDesProv 
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
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   128
            Tag             =   "Descripción"
            Top             =   780
            Width           =   6360
         End
         Begin VB.CheckBox chkProveedor 
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   390
            TabIndex        =   53
            Top             =   877
            Width           =   1125
         End
         Begin VB.CheckBox chkTipoProveedor 
            Caption         =   "Tipo Prov"
            Height          =   195
            Left            =   390
            TabIndex        =   52
            Top             =   525
            Width           =   1050
         End
         Begin VB.ComboBox cboBuscaTipoProveedor 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   405
            Width           =   3900
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   390
            TabIndex        =   56
            Top             =   1935
            Width           =   810
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   975
            Left            =   10680
            MaskColor       =   &H000000FF&
            Picture         =   "frmPagoProveedores.frx":180C
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Buscar "
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   2760
            TabIndex        =   64
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   57278465
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   5490
            TabIndex        =   65
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   57278465
            CurrentDate     =   41098
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bco:"
            Height          =   195
            Left            =   5280
            TabIndex        =   147
            Top             =   1560
            Width           =   330
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Bco:"
            Height          =   195
            Left            =   5280
            TabIndex        =   146
            Top             =   1200
            Width           =   330
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1650
            TabIndex        =   70
            Top             =   1980
            Width           =   1005
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4440
            TabIndex        =   69
            Top             =   1980
            Width           =   960
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   3405
         Left            =   -74850
         TabIndex        =   66
         Top             =   2760
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   6006
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388736
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
         TabIndex        =   76
         Top             =   570
         Width           =   1065
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   7440
      Top             =   6330
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      TabIndex        =   77
      Top             =   6420
      Width           =   750
   End
End
Attribute VB_Name = "frmPagoProveedores"
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
        Suma = Suma + CDbl(IIf(Grilla.TextMatrix(I, COLUMNA) = "", 0, Grilla.TextMatrix(I, COLUMNA)))
    Next
    SumaGrilla = Valido_Importe(CStr(Suma))
End Function

Private Sub CboBanco_LostFocus()
    If cboBanco.ListIndex <> -1 Then
        Call CargoCtaBancaria(CStr(cboBanco.ItemData(cboBanco.ListIndex)))
        'cboCtaBancaria.ListIndex = 1
    End If
End Sub

Private Sub cboCtaBancaria_LostFocus()
    If cboCtaBancaria.ListIndex <> -1 Then
        'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE_PROPIO "
        sql = sql & " WHERE CHEP_NUMERO LIKE '" & TxtCheNumero.Text & "' "
        sql = sql & " AND BAN_CODINT = " & XN(cboBanco.ItemData(cboBanco.ListIndex))
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then 'EXISTE
            Me.TxtCheFecVto.Value = rec!CHEP_FECVTO
            Me.TxtCheImport.Text = Valido_Importe(rec!CHEP_IMPORT)
        Else
           MsgBox "El Cheque no fue registrado, el mismo debe ser registrado con anterioridad", vbInformation, TIT_MSGBOX
           'rec.Close
           cmdNuevoCheque_Click
        End If
        rec.Close
    End If
End Sub

Private Sub Check1_Click()

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

Private Sub chkPropio_Click()
    If chkPropio.Value = Checked Then
        cmdBuscaPCheque.Enabled = True
    Else
        cmdBuscaPCheque.Enabled = False
        txtchepro.Text = ""
        txtimpChe1.Text = ""
        txtBBanco1.Text = ""
    End If
End Sub

Private Sub chkProveedor_Click()
    If chkProveedor.Value = Checked Then
        txtProveedor.Enabled = True
        cmdBuscarProveedor.Enabled = True
    Else
        txtProveedor.Text = ""
        txtProveedor.Enabled = False
        cmdBuscarProveedor.Enabled = False
    End If
End Sub

Private Sub chkTercero_Click()
    If chkTercero.Value = Checked Then
        cmdBuscaCheque.Enabled = True
    Else
        cmdBuscaCheque.Enabled = False
        txtcheqTer.Text = ""
        txtImpChe2.Text = ""
        txtBBanco2.Text = ""
    End If
End Sub

Private Sub chkTipoProveedor_Click()
    If chkTipoProveedor.Value = Checked Then
        cboBuscaTipoProveedor.Enabled = True
        cboBuscaTipoProveedor.ListIndex = 0
    Else
        cboBuscaTipoProveedor.Enabled = False
        cboBuscaTipoProveedor.ListIndex = -1
    End If
End Sub

Private Sub cmaAceptarACta_Click()
    txtSaldoACta.Text = ""
    txtImporteACta.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdAceptarCheques_Click()
    
    If GrillaCheques.Rows > 1 Then
        'CARGO EN GRILLA VALORES
        For I = 1 To GrillaCheques.Rows - 1
            If GrillaCheques.TextMatrix(I, 9) = "P" Then
                grillaValores.AddItem "CHE" & "-" & GrillaCheques.TextMatrix(I, 9) & Chr(9) & GrillaCheques.TextMatrix(I, 6) & Chr(9) & _
                                      GrillaCheques.TextMatrix(I, 5) & Chr(9) & GrillaCheques.TextMatrix(I, 8) _
                                      & Chr(9) & GrillaCheques.TextMatrix(I, 4) & Chr(9) & GrillaCheques.TextMatrix(I, 7) & Chr(9) & GrillaCheques.TextMatrix(I, 3)
            Else
                grillaValores.AddItem "CHE" & "-" & GrillaCheques.TextMatrix(I, 9) & Chr(9) & _
                                       GrillaCheques.TextMatrix(I, 6) & Chr(9) & _
                                       GrillaCheques.TextMatrix(I, 5) & Chr(9) & _
                                       GrillaCheques.TextMatrix(I, 8) & Chr(9) & _
                                       GrillaCheques.TextMatrix(I, 4) & Chr(9) & _
                                       GrillaCheques.TextMatrix(I, 7)

            End If
        Next
        txtTotalValores.Text = Valido_Importe(CStr(SumaGrilla(grillaValores, 1)))
        grillaValores.HighLight = flexHighlightAlways
        GrillaCheques.Rows = 1
        txtTotalCheques.Text = ""
        tabValores.Tab = 0
    End If
End Sub

Private Sub cmdAceptarComprobante_Click()
    If GrillaComp.Rows > 1 Then
        'CARGO EN GRILLA VALORES
        For I = 1 To GrillaComp.Rows - 1
            grillaValores.AddItem "COMP" & Chr(9) & GrillaComp.TextMatrix(I, 3) & Chr(9) & GrillaComp.TextMatrix(I, 2) _
                                   & Chr(9) & GrillaComp.TextMatrix(I, 0) & Chr(9) & GrillaComp.TextMatrix(I, 1) & Chr(9) & _
                                   GrillaComp.TextMatrix(I, 4)
        Next
        txtTotalValores.Text = Valido_Importe(CStr(SumaGrilla(grillaValores, 1)))
        grillaValores.HighLight = flexHighlightAlways
        GrillaComp.Rows = 1
        txtTotalComprobante.Text = ""
        tabValores.Tab = 0
    End If
End Sub

Private Sub cmdAceptarComprobantes_Click()
    txtSaldo.Text = ""
    txtImporteApagar.Text = ""
    tabComprobantes.Tab = 0
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
    If cmdGrabar.Enabled = True Then
        cmdGrabar.SetFocus
    Else
        cmdNuevo.SetFocus
    End If
End Sub

Private Sub cmdAgregarACta_Click()
    If GrillaAFavor.Rows > 1 Then
        If grillaValores.Rows > 1 Then
            For I = 1 To grillaValores.Rows - 1
                If GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 5) = grillaValores.TextMatrix(I, 5) _
                    And GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 1) = grillaValores.TextMatrix(I, 4) _
                    And CDate(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 2)) = CDate(grillaValores.TextMatrix(I, 2)) Then
                   MsgBox "El Valor ya fue ingresado", vbInformation, TIT_MSGBOX
                   txtSaldoACta.Text = ""
                   txtImporteACta.Text = ""
                   GrillaAFavor.SetFocus
                   Exit Sub
                End If
            Next
        End If
                
        'CARGO EN GRILLA VALORES
        grillaValores.AddItem "A-CTA" & Chr(9) & Valido_Importe(txtImporteACta) & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 2) _
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
    
    If TxtCheNumero.Text = "" Then
        MsgBox "Debe ingresar el número del cheque", vbExclamation, TIT_MSGBOX
        TxtCheNumero.SetFocus
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
    If optChequePropio.Value = True Then
        If cboBanco.ListIndex = -1 Then
            MsgBox "Debe ingresar un BANCO", vbExclamation, TIT_MSGBOX
            cboBanco.SetFocus
            Exit Sub
        End If
        If cboCtaBancaria.ListIndex = -1 Then
            MsgBox "Debe ingresar la CTA-BANCARIA", vbExclamation, TIT_MSGBOX
            cboCtaBancaria.SetFocus
            Exit Sub
        End If
        GrillaCheques.AddItem "" & Chr(9) & "" & Chr(9) & _
                      "" & Chr(9) & cboCtaBancaria.Text & Chr(9) & _
                      TxtCheNumero.Text & Chr(9) & TxtCheFecVto.Value & Chr(9) & _
                      TxtCheImport.Text & Chr(9) & cboBanco.ItemData(cboBanco.ListIndex) & Chr(9) & _
                      cboBanco.List(cboBanco.ListIndex) & Chr(9) & "P"
    Else
        
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
        GrillaCheques.AddItem TxtBanco.Text & Chr(9) & txtlocalidad.Text & Chr(9) & _
                          TxtSucursal.Text & Chr(9) & txtcodigo.Text & Chr(9) & _
                          TxtCheNumero.Text & Chr(9) & TxtCheFecVto.Value & Chr(9) & _
                          TxtCheImport.Text & Chr(9) & TxtCodInt.Text & Chr(9) & _
                          TxtBanDescri.Text & Chr(9) & "T"
    End If

    GrillaCheques.HighLight = flexHighlightAlways
    txtTotalCheques.Text = Valido_Importe(CStr(SumaGrilla(GrillaCheques, 6)))
    LimpiarCheques
    cmdAgregarCheque.Enabled = False
    TxtCheNumero.SetFocus
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
    cboCtaBancaria.Clear
    cboBanco.Enabled = False
    cboCtaBancaria.Enabled = False
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
    
    'VALIDO QUE EL CHEQUE NO SE HAYA CARGADO
    If GrillaAplicar.Rows > 1 Then
        
        If ValidoIngFactura(cboComprobantes, GrillaComp, fechaComprobantes, txtNroComprobantes) = False Then
            MsgBox "El Documento ya fue ingresado", vbCritical, TIT_MSGBOX
            txtNroComprobantes.Text = ""
            cboComprobantes.SetFocus
            Exit Sub
        End If
    End If
    
    'CARGO GRILLA
    GrillaComp.AddItem BuscarTipoDocAbre(CStr(cboComprobantes.ItemData(cboComprobantes.ListIndex))) & Chr(9) & txtNroComprobantes & Chr(9) & _
                       fechaComprobantes & Chr(9) & txtImporteComprobante.Text & Chr(9) & _
                       cboComprobantes.ItemData(cboComprobantes.ListIndex)

                           
    GrillaComp.HighLight = flexHighlightAlways
    txtTotalComprobante.Text = Valido_Importe(CStr(SumaGrilla(GrillaAplicar, 3)))
    txtNroComprobantes.Text = ""
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
                And (GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 1)) = (GrillaAplicar1.TextMatrix(I, 2)) _
                And CDate(GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 2)) = CDate(GrillaAplicar1.TextMatrix(I, 3)) Then
               MsgBox "La Factura ya fue elegida", vbInformation, TIT_MSGBOX
               txtSaldo.Text = ""
               txtImporteApagar.Text = ""
               GrillaAplicar.SetFocus
               Exit Sub
            End If
        Next
    End If
    If GrillaAplicar.CellForeColor = vbBlack Then
        Call CambiaColorAFilaDeGrilla(GrillaAplicar, GrillaAplicar.RowSel, vbRed)
    Else
        Call CambiaColorAFilaDeGrilla(GrillaAplicar, GrillaAplicar.RowSel, vbBlack)
    End If
    GrillaAplicar1.AddItem GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 0) & Chr(9) & _
                           Valido_Importe(txtImporteApagar.Text) & Chr(9) & _
                           GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 1) & Chr(9) & _
                           GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 2) & Chr(9) & _
                           Valido_Importe(CStr(CDbl(txtSaldo.Text) - CDbl(txtImporteApagar.Text))) & Chr(9) & _
                           GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 5) & Chr(9) & _
                           GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 6) & Chr(9) & _
                           GrillaAplicar.TextMatrix(GrillaAplicar.RowSel, 7)

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
    GrillaEfectivo.AddItem cboMoneda.Text & Chr(9) & txtEftImporte.Text _
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

Private Function ValidoIngFactura(combo As ComboBox, Grilla As MSFlexGrid, Fecha As String, NROFAC As String) As Boolean
    For I = 1 To Grilla.Rows - 1
        If combo.ItemData(combo.ListIndex) = Grilla.TextMatrix(I, 4) And _
           Fecha = Grilla.TextMatrix(I, 2) And _
           NROFAC = Grilla.TextMatrix(I, 1) Then
           
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
    frmBuscar.TipoBusqueda = 8
    frmBuscar.Show vbModal
    'TxtCheNumero.Text = frmBuscar.grdBuscar.Col
    txtcheqTer.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
    txtBBanco2.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 10)
    txtImpChe2.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
End Sub

Private Sub cmdBuscaPCheque_Click()
    Dim codint As Integer
    frmBuscar.TipoBusqueda = 9
    frmBuscar.Show vbModal
    'TxtCheNumero.Text = frmBuscar.grdBuscar.Col
    txtchepro.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
    txtBBanco1.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 10)
    txtimpChe1.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT DISTINCT OP.OPG_NUMERO, OP.OPG_FECHA,TP.TPR_DESCRI, P.PROV_RAZSOC"
    sql = sql & " FROM ORDEN_PAGO OP, TIPO_PROVEEDOR TP, PROVEEDOR P,DETALLE_ORDEN_PAGO DO"
    sql = sql & " WHERE"
    sql = sql & " OP.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND OP.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND P.TPR_CODIGO=TP.TPR_CODIGO"
    sql = sql & " AND DO.OPG_NUMERO = OP.OPG_NUMERO"
    sql = sql & " AND DO.OPG_FECHA = OP.OPG_FECHA"
    sql = sql & " AND DO.TCO_CODIGO = OP.TCO_CODIGO"
    If chkTipoProveedor.Value = Checked Then sql = sql & " AND OP.TPR_CODIGO=" & cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.ListIndex)
    If txtProveedor.Text <> "" Then sql = sql & " AND OP.PROV_CODIGO=" & XN(txtProveedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND OP.OPG_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND OP.OPG_FECHA<=" & XDQ(FechaHasta)
    If txtcheqTer.Text <> "" Then sql = sql & " AND DO.CHE_NUMERO LIKE '" & Trim(txtcheqTer) & "'"
    If txtchepro.Text <> "" Then sql = sql & " AND DO.CHE_NUMERO LIKE '" & Trim(txtchepro) & "'"
    sql = sql & " ORDER BY OP.OPG_FECHA, OP.OPG_NUMERO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            GrdModulos.AddItem Rec1!OPG_NUMERO & Chr(9) & Rec1!OPG_FECHA & Chr(9) & _
                               Rec1!TPR_DESCRI & Chr(9) & Rec1!PROV_RAZSOC
                               
            Rec1.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        chkTipoProveedor.SetFocus
    End If
    Rec1.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdBuscarCheques_Click()
    If optChequeTercero.Value = True Then
        frmBuscar.TipoBusqueda = 6
        frmBuscar.TxtDescriB = ""
        frmBuscar.Show vbModal
        If frmBuscar.grdBuscar.Text <> "" Then
            frmBuscar.grdBuscar.Col = 1
            TxtCheNumero.Text = frmBuscar.grdBuscar.Text
            frmBuscar.grdBuscar.Col = 2
            TxtCheFecVto.Value = frmBuscar.grdBuscar.Text
            frmBuscar.grdBuscar.Col = 3
            TxtCheImport.Text = frmBuscar.grdBuscar.Text
            frmBuscar.grdBuscar.Col = 4
            TxtCodInt.Text = frmBuscar.grdBuscar.Text
            frmBuscar.grdBuscar.Col = 5
            TxtBanco.Text = frmBuscar.grdBuscar.Text
            frmBuscar.grdBuscar.Col = 6
            txtlocalidad.Text = frmBuscar.grdBuscar.Text
            frmBuscar.grdBuscar.Col = 7
            TxtSucursal.Text = frmBuscar.grdBuscar.Text
            frmBuscar.grdBuscar.Col = 8
            txtcodigo.Text = frmBuscar.grdBuscar.Text
            cmdAgregarCheque_Click
        Else
            TxtCheNumero.SetFocus
        End If
    End If
    If optChequePropio.Value = True Then
         Dim codint As Integer
        frmBuscar.TipoBusqueda = 9
        frmBuscar.Show vbModal
        frmBuscar.grdBuscar.Col = 1
        TxtCheNumero.Text = frmBuscar.grdBuscar.Text
        
        cboCtaBancaria_LostFocus
        cboBanco.ListIndex = 0
        'txtchepro.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
        'txtBBanco1.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 10)
        'txtimpChe1.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
    End If
End Sub

Private Sub cmdBuscarProveedor_Click()
    frmBuscar.TipoBusqueda = 5
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 1
        txtProveedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 2
        txtDesProv.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 3
        Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboBuscaTipoProveedor)
        txtProveedor.SetFocus
    Else
        txtProveedor.SetFocus
    End If
End Sub

Private Sub cmdBuscarProveedor1_Click()
    frmBuscar.TipoBusqueda = 5
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 1
        txtCodProveedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 2
        txtProvRazSoc.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 3
        Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboTipoProveedor)
        'txtProvRazSoc.SetFocus
        txtCodProveedor_LostFocus
    Else
        txtCodProveedor.SetFocus
    End If
End Sub

Private Sub cmdCancelarCheques_Click()
    GrillaCheques.Rows = 1
    txtTotalCheques.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdCancelarComprobante_Click()
    GrillaComp.Rows = 1
    txtTotalComprobante.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdCancelarMoneda_Click()
    GrillaEfectivo.Rows = 1
    txtTotalEfectivo.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdGrabar_Click()
    If ValidarOrdenPago = False Then Exit Sub
    If MsgBox("¿Confirma Orden de Pago?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayError
    DBConn.BeginTrans
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    sql = "SELECT EST_CODIGO"
    sql = sql & " FROM ORDEN_PAGO"
    sql = sql & " WHERE"
    sql = sql & " OPG_NUMERO=" & XN(txtNroOrdenPago.Text)
    sql = sql & " AND OPG_FECHA=" & XDQ(FechaOrdenPago.Value)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = True Then
        'CABEZA DE LA ORDEN DE PAGO
        sql = "INSERT INTO ORDEN_PAGO ("
        sql = sql & " OPG_NUMERO,OPG_FECHA,TCO_CODIGO,EST_CODIGO,TPR_CODIGO,PROV_CODIGO,OPG_TOTAL,OPG_NROSUC,"
        sql = sql & "OPG_NROSUCTXT,OPG_NUMEROTXT)"
        sql = sql & " VALUES ("
        sql = sql & XN(txtNroOrdenPago.Text) & ","
        sql = sql & XDQ(FechaOrdenPago.Value) & ","
        sql = sql & cboOrdPag.ItemData(cboOrdPag.ListIndex) & ","
        sql = sql & "3," 'ESTADO DEFINITIVO
        sql = sql & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex) & ","
        sql = sql & XN(txtCodProveedor.Text) & ","
        sql = sql & XN(txtTotalAplicar) & ","
        sql = sql & XN("1") & ","
        sql = sql & XS(Sucursal) & ","
        sql = sql & XS(Format(txtNroOrdenPago, "00000000")) & ")"
        DBConn.Execute sql
        
        'DETALLE DE LA ORDEN DE PAGO
        For I = 1 To grillaValores.Rows - 1
            sql = "INSERT INTO DETALLE_ORDEN_PAGO"
            sql = sql & " (OPG_NUMERO,OPG_FECHA,DOP_NROITEM,TCO_CODIGO,BAN_CODINT,CHE_NUMERO,CTA_NROCTA,"
            sql = sql & " MON_CODIGO,DOP_MONIMP,DOP_TCO_CODIGO,DOP_COMFECHA,DOP_COMNUMERO,DOP_COMIMP)"
            sql = sql & " VALUES ("
            sql = sql & XN(txtNroOrdenPago.Text) & ","
            sql = sql & XDQ(FechaOrdenPago.Value) & ","
            sql = sql & XN(CStr(I)) & ","
            sql = sql & cboOrdPag.ItemData(cboOrdPag.ListIndex) & "," 'TIPO ORDEN DE PAGO
            
            If grillaValores.TextMatrix(I, 0) = "CHE-P" Then
                sql = sql & XN(grillaValores.TextMatrix(I, 5)) & "," 'CODIGO BANCO
                sql = sql & XS(grillaValores.TextMatrix(I, 4)) & "," 'NUMERO CHEQUE
                sql = sql & XS(grillaValores.TextMatrix(I, 6)) & "," 'CTA_NROCTA
            ElseIf grillaValores.TextMatrix(I, 0) = "CHE-T" Then
                '-------------------------------
                sql = sql & XN(grillaValores.TextMatrix(I, 5)) & "," 'CODIGO BANCO
                sql = sql & XS(grillaValores.TextMatrix(I, 4)) & "," 'NUMERO CHEQUE
                sql = sql & "NULL," 'CTA_NROCTA
            Else
                sql = sql & "NULL,NULL,NULL,"
            End If
            If grillaValores.TextMatrix(I, 0) = "EFT" Then
                sql = sql & XN(grillaValores.TextMatrix(I, 5)) & "," 'CODIGO MONEDA
                sql = sql & XN(grillaValores.TextMatrix(I, 1)) & "," 'IMPORTE
            Else
                sql = sql & "NULL,NULL,"
            End If
            If grillaValores.TextMatrix(I, 0) = "COMP" Or grillaValores.TextMatrix(I, 0) = "A-CTA" Then
                sql = sql & XN(grillaValores.TextMatrix(I, 5)) & ","  'CODIGO COMPROBANTE
                sql = sql & XDQ(grillaValores.TextMatrix(I, 2)) & "," 'FECHA COMPROBANTE
                sql = sql & Right(grillaValores.TextMatrix(I, 4), 8) & "," 'NUMERO COMPROBANTE
                sql = sql & XN(grillaValores.TextMatrix(I, 1)) & ")"  'IMPORTE COMPROBANTE
            Else
                sql = sql & "NULL,NULL,NULL,NULL)"
            End If
            DBConn.Execute sql
        Next
        'FACTURAS CANCELADAS EN LA ORDEN DE PAGO
        For I = 1 To GrillaAplicar1.Rows - 1
            sql = "INSERT INTO FACTURAS_ORDEN_PAGO"
            sql = sql & " (TPR_CODIGO,PROV_CODIGO,FPR_TCO_CODIGO,FPR_NROSUC,FPR_NUMERO,FPR_FECHA,"
            sql = sql & "OPG_NUMERO,OPG_FECHA,TCO_CODIGO,OPG_IMPORTE)"
            sql = sql & " VALUES ("
            sql = sql & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex) & ","
            sql = sql & XN(txtCodProveedor) & ","
            sql = sql & XN(GrillaAplicar1.TextMatrix(I, 5)) & "," 'CODIGO COMPROBANTE
            sql = sql & XN(GrillaAplicar1.TextMatrix(I, 6)) & "," 'NRO SUCURSAL
            sql = sql & XN(GrillaAplicar1.TextMatrix(I, 7)) & "," 'NRO COMPROBANTE
            sql = sql & XDQ(GrillaAplicar1.TextMatrix(I, 3)) & "," 'FECHA COMPROBANTE
            sql = sql & XN(txtNroOrdenPago) & ","
            sql = sql & XDQ(FechaOrdenPago) & ","
            sql = sql & cboOrdPag.ItemData(cboOrdPag.ListIndex) & "," 'TIPO ORDEN DE PAGO
            sql = sql & XN(GrillaAplicar1.TextMatrix(I, 1)) & ")"     'IMPORTE PAGADO DEL COMPROBANTE EN LA ORDEN DE PAGO
            DBConn.Execute sql
        Next
        
        'ACTUALIZO EL SALDO DE LAS FACTURAS ELEGIDAS
        For I = 1 To GrillaAplicar1.Rows - 1
            'FALTA HACER ESTO PARA LOS GASTOS
            'ESTO LO COMENTO PQ HACE LO MISMO PARA CUALQUIER TIPO DE CBTE
            
            'If XN(GrillaAplicar1.TextMatrix(I, 5)) = "" Or XN(GrillaAplicar1.TextMatrix(I, 5)) = 1 Or XN(GrillaAplicar1.TextMatrix(I, 5)) = 12 Then
                sql = "UPDATE FACTURA_PROVEEDOR"
                sql = sql & " SET FPR_SALDO=" & XN(GrillaAplicar1.TextMatrix(I, 4))
                sql = sql & " WHERE"
                sql = sql & " TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
                sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
                sql = sql & " AND TCO_CODIGO=" & XN(GrillaAplicar1.TextMatrix(I, 5))
                sql = sql & " AND FPR_NROSUC=" & XN(GrillaAplicar1.TextMatrix(I, 6))
                sql = sql & " AND FPR_NUMERO=" & XN(GrillaAplicar1.TextMatrix(I, 7))
                DBConn.Execute sql
                
                sql = "UPDATE GASTOS_GENERALES"
                sql = sql & " SET GGR_SALDO=" & XN(GrillaAplicar1.TextMatrix(I, 4))
                sql = sql & " WHERE"
                sql = sql & " TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
                sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
                sql = sql & " AND TCO_CODIGO=" & XN(GrillaAplicar1.TextMatrix(I, 5))
                sql = sql & " AND GGR_NROSUC=" & XN(GrillaAplicar1.TextMatrix(I, 6))
                sql = sql & " AND GGR_NROCOMP=" & XN(GrillaAplicar1.TextMatrix(I, 7))
                DBConn.Execute sql
            'Else
                sql = "UPDATE NOTA_DEBITO_PROVEEDOR"
                sql = sql & " SET DPR_SALDO=" & XN(GrillaAplicar1.TextMatrix(I, 4))
                sql = sql & " WHERE"
                sql = sql & " TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
                sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
                sql = sql & " AND TCO_CODIGO=" & XN(GrillaAplicar1.TextMatrix(I, 5))
                sql = sql & " AND DPR_NROSUC=" & XN(GrillaAplicar1.TextMatrix(I, 6))
                sql = sql & " AND DPR_NUMERO=" & XN(GrillaAplicar1.TextMatrix(I, 7))
                DBConn.Execute sql
            'End If
            
        Next
        
        'ACTUALIZO EL DINERO A CUENTA (RECIBO_CLIENTE_SALDO)
        For I = 1 To GrillaAFavor.Rows - 1
            sql = "UPDATE ORDEN_PAGO_SALDO"
            sql = sql & " SET OPG_SALDO=" & XN(GrillaAFavor.TextMatrix(I, 4))
            sql = sql & " WHERE"
            sql = sql & " OPG_NUMERO=" & XN(GrillaAFavor.TextMatrix(I, 1))
            sql = sql & " AND OPG_FECHA=" & XDQ(GrillaAFavor.TextMatrix(I, 2))
            DBConn.Execute sql
            'ACTUALIZO SALDO EN FAC
            sql = "UPDATE FACTURA_CLIENTE"
            sql = sql & " SET FCL_SALDO=" & XN(GrillaAFavor.TextMatrix(I, 4))
            sql = sql & " WHERE"
            sql = sql & " FCL_NUMERO=" & Right(GrillaAFavor.TextMatrix(I, 1), 8)
            sql = sql & " AND FCL_FECHA=" & XDQ(GrillaAFavor.TextMatrix(I, 2))
            DBConn.Execute sql
        Next
        'VERIFICO SI HAY DINERO A CUENTA
        If CDbl(txtTotalAplicar.Text) < CDbl(txtTotalValores.Text) Then
            sql = "INSERT INTO ORDEN_PAGO_SALDO("
            sql = sql & "OPG_NUMERO,OPG_FECHA,TCO_CODIGO,OPG_TOTSALDO,OPG_SALDO)"
            sql = sql & " VALUES ("
            sql = sql & XN(txtNroOrdenPago) & ","
            sql = sql & XDQ(FechaOrdenPago) & ","
            sql = sql & cboOrdPag.ItemData(cboOrdPag.ListIndex) & "," 'TIPO ORDEN DE PAGO
            sql = sql & XN(CStr(CDbl(txtTotalValores.Text) - CDbl(txtTotalAplicar.Text))) & ","
            sql = sql & XN(CStr(CDbl(txtTotalValores.Text) - CDbl(txtTotalAplicar.Text))) & ")"
            DBConn.Execute sql
        End If
        
        'CAMBIO EL ESTADO A LOS CHEQUES UTILIZADOS
        For I = 1 To grillaValores.Rows - 1
            If grillaValores.TextMatrix(I, 0) = "CHE-T" Then 'CHEQUES DE TERCEROS
                'Cambio en Cheque_Estados 7 ES CHEQUES ENTREGADO
                sql = "INSERT INTO CHEQUE_ESTADOS"
                sql = sql & "(ECH_CODIGO,BAN_CODINT,CHE_NUMERO,CES_FECHA,CES_DESCRI) "
                sql = sql & " VALUES ( 7,"
                sql = sql & XN(grillaValores.TextMatrix(I, 5)) & ","
                sql = sql & XS(grillaValores.TextMatrix(I, 4)) & ","
                sql = sql & XDQ(Date) & ","
                sql = sql & XS(txtProvRazSoc.Text) & ")" ' PONGO A QUE PROVEEDOR SE LO ENTREGUE
                DBConn.Execute sql
            End If
        Next
        'ACTUALIZO EL SALDO DE LA CUENTA BANCARIA
'        For I = 1 To grillaValores.Rows - 1
'            If grillaValores.TextMatrix(I, 0) = "CHE-P" Then 'CHEQUES PROPIOS
'                sql = "UPDATE CTA_BANCARIA"
'                sql = sql & " SET CTA_SALACT = CTA_SALACT - " & XN(grillaValores.TextMatrix(I, 1)) 'IMPORTE
'                sql = sql & " WHERE BAN_CODINT = " & XN(grillaValores.TextMatrix(I, 5)) 'COD INT BANCO
'                sql = sql & " AND CTA_NROCTA = " & XS(grillaValores.TextMatrix(I, 6)) 'NRO CTA BANCARIA
'                DBConn.Execute sql
'            End If
'        Next
        
         'ACTUALIZO CUNETA CORRIENTE DEL PROVEEDOR
          DBConn.Execute AgregoCtaCteProveedores(CStr(cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)), txtCodProveedor, CStr(cboOrdPag.ItemData(cboOrdPag.ListIndex)) _
                                         , Sucursal, Format(txtNroOrdenPago, "00000000"), FechaOrdenPago, txtTotalValores.Text, "H", CStr(Date))
    Else
        'SI EXISTE
        MsgBox "La Orden de Pago ya Existe", vbCritical, TIT_MSGBOX
    End If
    DBConn.CommitTrans
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    rec.Close
    CmdNuevo_Click
    Exit Sub
    
HayError:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Function Cambio_Estado_Cheque()

End Function
Private Function ValidarOrdenPago() As Boolean
    
    If txtNroOrdenPago.Text = "" Then
        MsgBox "Debe ingresar el número de Orden de Pago", vbCritical, TIT_MSGBOX
        txtNroOrdenPago.SetFocus
        ValidarOrdenPago = False
        Exit Function
    End If
    If IsNull(FechaOrdenPago.Value) Then
        MsgBox "Debe ingresar la fecha de la Orden de Pago", vbCritical, TIT_MSGBOX
        FechaOrdenPago.SetFocus
        ValidarOrdenPago = False
        Exit Function
    End If
    If txtCodProveedor.Text = "" Then
        MsgBox "Debe ingresar un Proveedor", vbCritical, TIT_MSGBOX
        txtCodProveedor.SetFocus
        ValidarOrdenPago = False
        Exit Function
    End If
    If grillaValores.Rows = 1 Then
        MsgBox "Debe ingresar Valores Recibidos", vbCritical, TIT_MSGBOX
        cmdAgregarCHE.SetFocus
        ValidarOrdenPago = False
        Exit Function
    End If
    If GrillaAplicar.Rows = 1 Then
        MsgBox "Debe ingresar una Factura", vbCritical, TIT_MSGBOX
        cmdAgregarFactura.SetFocus
        ValidarOrdenPago = False
        Exit Function
    End If
    If CDbl(IIf(txtTotalAplicar.Text = "", 0, txtTotalAplicar.Text)) > CDbl(txtTotalValores.Text) Then
        MsgBox "El Total de Facturas supera al Total de Valores Recibidos", vbCritical, TIT_MSGBOX
        cmdAgregarCHE.SetFocus
        ValidarOrdenPago = False
        Exit Function
    End If
    If CDbl(IIf(txtTotalAplicar.Text = "", 0, txtTotalAplicar.Text)) < CDbl(txtTotalValores.Text) Then
        If MsgBox("El Total de Valores Recibidos supera al Total de Facturas," & Chr(13) & _
                "deja el importe (" & Format(CStr(CDbl(txtTotalValores.Text) - CDbl(txtTotalAplicar.Text)), "#,##0.00") & _
                ") como dinero a cuenta", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then

            cmdAgregarFactura.SetFocus
            ValidarOrdenPago = False
            Exit Function
        End If
    End If
    ValidarOrdenPago = True
End Function

Private Sub cmdImprimir_Click()
    If txtCodProveedor.Text = "" Or GrillaAplicar1.Rows = 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Buscando Orden de Pago..."
    
    sql = "DELETE FROM TMP_ORDEN_PAGO"
    DBConn.Execute sql

    Call OrdenPagoNotasDebito(txtNroOrdenPago, FechaOrdenPago)
    Call OrdenPagoFacturas(txtNroOrdenPago, FechaOrdenPago)
    Call OrdenPagoComprobante(txtNroOrdenPago, FechaOrdenPago)
    Call OrdenPagoCheques(txtNroOrdenPago, FechaOrdenPago)
    Call OrdenPagoMoneda(txtNroOrdenPago, FechaOrdenPago)

    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SITOTALCAR"
    Rep.SelectionFormula = ""

    Rep.WindowTitle = "Orden de Pago"
    Rep.ReportFileName = DRIVE & DirReport & "rptordenpago.rpt"
    
    Rep.Destination = crptToWindow
    Rep.Action = 1
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    Rep.SelectionFormula = ""
End Sub

Private Sub OrdenPagoComprobante(OrdPago As String, Fecha As String)
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT P.PROV_RAZSOC, P.PROV_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, P.PROV_CUIT, P.PROV_INGBRU"
    sql = sql & ",C.IVA_DESCRI, TC.TCO_ABREVIA, DP.DOP_COMFECHA, DP.DOP_COMNUMERO, DP.DOP_COMIMP, OP.OPG_TOTAL"
    sql = sql & " FROM PROVEEDOR P,DETALLE_ORDEN_PAGO DP, ORDEN_PAGO OP ,CONDICION_IVA C"
    sql = sql & " ,LOCALIDAD L, PROVINCIA PR, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE OP.OPG_NUMERO=" & XN(OrdPago)
    sql = sql & " AND OP.OPG_FECHA=" & XDQ(Fecha)
    sql = sql & " AND OP.OPG_NUMERO=DP.OPG_NUMERO"
    sql = sql & " AND OP.OPG_FECHA=DP.OPG_FECHA"
    sql = sql & " AND OP.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND OP.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND DP.DOP_TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND P.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND P.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND P.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND P.IVA_CODIGO=C.IVA_CODIGO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_ORDEN_PAGO ("
            sql = sql & "OPG_NUMERO,OPG_FECHA,PROV_RAZSOC,PROV_DOMICI,PROV_CUIT,PROV_INGBRU,"
            sql = sql & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            sql = sql & "OPG_TOTAL,FAC_ABREVIA,FAC_NUMERO,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL) VALUES ("
            sql = sql & XS(Format(txtNroOrdenPago, "00000000")) & ","
            sql = sql & XDQ(FechaOrdenPago) & ","
            sql = sql & XS(Rec1!PROV_RAZSOC) & ","
            sql = sql & XS(Rec1!PROV_DOMICI) & ","
            sql = sql & XS(Format(Rec1!PROV_CUIT, "##-########-#")) & ","
            sql = sql & XS(Format(Rec1!PROV_INGBRU, "###-#####-##")) & ","
            sql = sql & XS(Rec1!LOC_DESCRI) & ","
            sql = sql & XS(Rec1!PRO_DESCRI) & ","
            sql = sql & XS(Rec1!IVA_DESCRI) & ","
            sql = sql & XS(Rec1!TCO_ABREVIA) & ","
            sql = sql & XDQ(Rec1!DOP_COMFECHA) & ","
            sql = sql & XS(Format(Rec1!DOP_COMNUMERO, "00000000")) & ","
            sql = sql & XN(Rec1!DOP_COMIMP) & ","
            sql = sql & XN(Rec1!OPG_TOTAL) & ","
            sql = sql & " NULL,NULL,NULL,NULL,NULL)"
            DBConn.Execute sql
            
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub OrdenPagoCheques(OrdPago As String, Fecha As String)
    Set Rec1 = New ADODB.Recordset
    'PARA CHEQUES DE TERCEROS
    sql = "SELECT P.PROV_RAZSOC, P.PROV_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, P.PROV_CUIT, P.PROV_INGBRU"
    sql = sql & ",C.IVA_DESCRI, B.BAN_NOMCOR, CH.CHE_FECVTO ,DP.CHE_NUMERO, CH.CHE_IMPORT, OP.OPG_TOTAL"
    sql = sql & " FROM PROVEEDOR P,DETALLE_ORDEN_PAGO DP, ORDEN_PAGO OP ,CONDICION_IVA C"
    sql = sql & " ,LOCALIDAD L, PROVINCIA PR, CHEQUE CH, BANCO B"
    sql = sql & " WHERE OP.OPG_NUMERO=" & XN(OrdPago)
    sql = sql & " AND OP.OPG_FECHA=" & XDQ(Fecha)
    sql = sql & " AND OP.OPG_NUMERO=DP.OPG_NUMERO"
    sql = sql & " AND OP.OPG_FECHA=DP.OPG_FECHA"
    sql = sql & " AND OP.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND OP.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND P.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND P.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND P.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND P.IVA_CODIGO=C.IVA_CODIGO"
    sql = sql & " AND DP.BAN_CODINT=CH.BAN_CODINT"
    sql = sql & " AND DP.CHE_NUMERO=CH.CHE_NUMERO"
    sql = sql & " AND CH.BAN_CODINT=B.BAN_CODINT"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_ORDEN_PAGO ("
            sql = sql & "OPG_NUMERO,OPG_FECHA,PROV_RAZSOC,PROV_DOMICI,PROV_CUIT,PROV_INGBRU,"
            sql = sql & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            sql = sql & "OPG_TOTAL,FAC_ABREVIA,FAC_NUMERO,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL) VALUES ("
            sql = sql & XS(Format(txtNroOrdenPago, "00000000")) & ","
            sql = sql & XDQ(FechaOrdenPago) & ","
            sql = sql & XS(Rec1!PROV_RAZSOC) & ","
            sql = sql & XS(Rec1!PROV_DOMICI) & ","
            sql = sql & XS(Format(Rec1!PROV_CUIT, "##-########-#")) & ","
            sql = sql & XS(Format(Rec1!PROV_INGBRU, "###-#####-##")) & ","
            sql = sql & XS(Rec1!LOC_DESCRI) & ","
            sql = sql & XS(Rec1!PRO_DESCRI) & ","
            sql = sql & XS(Rec1!IVA_DESCRI) & ","
            sql = sql & XS(Rec1!BAN_NOMCOR) & ","
            sql = sql & XDQ(Rec1!CHE_FECVTO) & ","
            sql = sql & XS(Rec1!CHE_NUMERO) & ","
            sql = sql & XN(Rec1!che_import) & ","
            sql = sql & XN(Rec1!OPG_TOTAL) & ","
            sql = sql & " NULL,NULL,NULL,NULL,NULL)"
            DBConn.Execute sql
            
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
    'PARA CHEQUES PROPIOS
    sql = "SELECT P.PROV_RAZSOC, P.PROV_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, P.PROV_CUIT, P.PROV_INGBRU"
    sql = sql & ",C.IVA_DESCRI, B.BAN_NOMCOR, CH.CHEP_FECVTO ,DP.CHE_NUMERO, CH.CHEP_IMPORT, OP.OPG_TOTAL"
    sql = sql & " FROM PROVEEDOR P,DETALLE_ORDEN_PAGO DP, ORDEN_PAGO OP ,CONDICION_IVA C"
    sql = sql & " ,LOCALIDAD L, PROVINCIA PR, CHEQUE_PROPIO CH, BANCO B"
    sql = sql & " WHERE OP.OPG_NUMERO=" & XN(OrdPago)
    sql = sql & " AND OP.OPG_FECHA=" & XDQ(Fecha)
    sql = sql & " AND OP.OPG_NUMERO=DP.OPG_NUMERO"
    sql = sql & " AND OP.OPG_FECHA=DP.OPG_FECHA"
    sql = sql & " AND OP.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND OP.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND P.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND P.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND P.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND P.IVA_CODIGO=C.IVA_CODIGO"
    sql = sql & " AND DP.BAN_CODINT=CH.BAN_CODINT"
    sql = sql & " AND DP.CHE_NUMERO=CH.CHEP_NUMERO"
    sql = sql & " AND CH.BAN_CODINT=B.BAN_CODINT"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_ORDEN_PAGO ("
            sql = sql & "OPG_NUMERO,OPG_FECHA,PROV_RAZSOC,PROV_DOMICI,PROV_CUIT,PROV_INGBRU,"
            sql = sql & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            sql = sql & "OPG_TOTAL,FAC_ABREVIA,FAC_NUMERO,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL) VALUES ("
            sql = sql & XS(Format(txtNroOrdenPago, "00000000")) & ","
            sql = sql & XDQ(FechaOrdenPago) & ","
            sql = sql & XS(Rec1!PROV_RAZSOC) & ","
            sql = sql & XS(Rec1!PROV_DOMICI) & ","
            sql = sql & XS(Format(Rec1!PROV_CUIT, "##-########-#")) & ","
            sql = sql & XS(Format(Rec1!PROV_INGBRU, "###-#####-##")) & ","
            sql = sql & XS(Rec1!LOC_DESCRI) & ","
            sql = sql & XS(Rec1!PRO_DESCRI) & ","
            sql = sql & XS(Rec1!IVA_DESCRI) & ","
            sql = sql & XS(Rec1!BAN_NOMCOR) & ","
            sql = sql & XDQ(Rec1!CHEP_FECVTO) & ","
            sql = sql & XS(Rec1!CHE_NUMERO) & ","
            sql = sql & XN(Rec1!CHEP_IMPORT) & ","
            sql = sql & XN(Rec1!OPG_TOTAL) & ","
            sql = sql & " NULL,NULL,NULL,NULL,NULL)"
            DBConn.Execute sql
            
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub OrdenPagoMoneda(OrdPago As String, Fecha As String)
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT P.PROV_RAZSOC, P.PROV_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, P.PROV_CUIT, P.PROV_INGBRU"
    sql = sql & ",C.IVA_DESCRI, M.MON_DESCRI, DP.DOP_MONIMP, OP.OPG_TOTAL"
    sql = sql & " FROM PROVEEDOR P,DETALLE_ORDEN_PAGO DP, ORDEN_PAGO OP ,CONDICION_IVA C"
    sql = sql & " ,LOCALIDAD L, PROVINCIA PR, MONEDA M"
    sql = sql & " WHERE OP.OPG_NUMERO=" & XN(OrdPago)
    sql = sql & " AND OP.OPG_FECHA=" & XDQ(Fecha)
    sql = sql & " AND OP.OPG_NUMERO=DP.OPG_NUMERO"
    sql = sql & " AND OP.OPG_FECHA=DP.OPG_FECHA"
    sql = sql & " AND OP.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND OP.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND P.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND P.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND P.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND P.IVA_CODIGO=C.IVA_CODIGO"
    sql = sql & " AND DP.MON_CODIGO=M.MON_CODIGO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_ORDEN_PAGO ("
            sql = sql & "OPG_NUMERO,OPG_FECHA,PROV_RAZSOC,PROV_DOMICI,PROV_CUIT,PROV_INGBRU,"
            sql = sql & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            sql = sql & "OPG_TOTAL,FAC_ABREVIA,FAC_NUMERO,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL) VALUES ("
            sql = sql & XS(Format(txtNroOrdenPago, "00000000")) & ","
            sql = sql & XDQ(FechaOrdenPago) & ","
            sql = sql & XS(Rec1!PROV_RAZSOC) & ","
            sql = sql & XS(Rec1!PROV_DOMICI) & ","
            sql = sql & XS(Format(Rec1!PROV_CUIT, "##-########-#")) & ","
            sql = sql & XS(Format(Rec1!PROV_INGBRU, "###-#####-##")) & ","
            sql = sql & XS(Rec1!LOC_DESCRI) & ","
            sql = sql & XS(Rec1!PRO_DESCRI) & ","
            sql = sql & XS(Rec1!IVA_DESCRI) & ","
            sql = sql & XS(Rec1!MON_DESCRI) & ","
            sql = sql & "NULL,"
            sql = sql & "NULL,"
            sql = sql & XN(Rec1!DOP_MONIMP) & ","
            sql = sql & XN(Rec1!OPG_TOTAL) & ","
            sql = sql & " NULL,NULL,NULL,NULL,NULL)"
            DBConn.Execute sql
            
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub OrdenPagoFacturas(OrdPago As String, Fecha As String)
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT P.PROV_RAZSOC, P.PROV_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, P.PROV_CUIT, P.PROV_INGBRU"
    sql = sql & ",C.IVA_DESCRI, TC.TCO_ABREVIA, FP.FPR_NROSUC, FP.FPR_NUMERO, FP.FPR_FECHA ,F.FPR_TOTAL, FP.OPG_IMPORTE, OP.OPG_TOTAL"
    sql = sql & " FROM PROVEEDOR P, ORDEN_PAGO OP ,CONDICION_IVA C"
    sql = sql & " ,LOCALIDAD L, PROVINCIA PR, TIPO_COMPROBANTE TC, FACTURAS_ORDEN_PAGO FP,"
    sql = sql & " FACTURA_PROVEEDOR F"
    sql = sql & " WHERE OP.OPG_NUMERO=" & XN(OrdPago)
    sql = sql & " AND OP.OPG_FECHA=" & XDQ(Fecha)
    sql = sql & " AND OP.OPG_NUMERO=FP.OPG_NUMERO"
    sql = sql & " AND OP.OPG_FECHA=FP.OPG_FECHA"
    sql = sql & " AND OP.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND OP.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND FP.FPR_TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND P.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND P.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND P.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND P.IVA_CODIGO=C.IVA_CODIGO"
    sql = sql & " AND FP.FPR_TCO_CODIGO=F.TCO_CODIGO"
    sql = sql & " AND FP.FPR_NROSUC=F.FPR_NROSUC"
    sql = sql & " AND FP.FPR_NUMERO=F.FPR_NUMERO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_ORDEN_PAGO ("
            sql = sql & "OPG_NUMERO,OPG_FECHA,PROV_RAZSOC,PROV_DOMICI,PROV_CUIT,PROV_INGBRU,"
            sql = sql & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            sql = sql & "OPG_TOTAL,FAC_ABREVIA,FAC_NUMERO,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL) VALUES ("
            sql = sql & XS(Format(txtNroOrdenPago, "00000000")) & ","
            sql = sql & XDQ(FechaOrdenPago) & ","
            sql = sql & XS(Rec1!PROV_RAZSOC) & ","
            sql = sql & XS(Rec1!PROV_DOMICI) & ","
            sql = sql & XS(Format(Rec1!PROV_CUIT, "##-########-#")) & ","
            sql = sql & XS(Format(Rec1!PROV_INGBRU, "###-#####-##")) & ","
            sql = sql & XS(Rec1!LOC_DESCRI) & ","
            sql = sql & XS(Rec1!PRO_DESCRI) & ","
            sql = sql & XS(Rec1!IVA_DESCRI) & ","
            sql = sql & "NULL,"
            sql = sql & "NULL,"
            sql = sql & "NULL,"
            sql = sql & "NULL,"
            sql = sql & XN(Rec1!OPG_TOTAL) & ","
            sql = sql & XS(Rec1!TCO_ABREVIA) & ","
            sql = sql & XS(Format(Rec1!FPR_NROSUC, "0000") & "-" & Format(Rec1!FPR_NUMERO, "00000000")) & ","
            sql = sql & XS(Rec1!FPR_FECHA) & ","
            sql = sql & XN(Rec1!OPG_IMPORTE) & ","
            sql = sql & XN(Rec1!FPR_TOTAL) & ")"
            DBConn.Execute sql
            
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub
Private Sub OrdenPagoNotasDebito(OrdPago As String, Fecha As String)
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT P.PROV_RAZSOC, P.PROV_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, P.PROV_CUIT, P.PROV_INGBRU"
    sql = sql & ",C.IVA_DESCRI, TC.TCO_ABREVIA, F.DPR_NROSUC, F.DPR_NUMERO, F.DPR_FECHA ,F.DPR_TOTAL, FP.OPG_IMPORTE, OP.OPG_TOTAL"
    sql = sql & " FROM PROVEEDOR P, ORDEN_PAGO OP ,CONDICION_IVA C"
    sql = sql & " ,LOCALIDAD L, PROVINCIA PR, TIPO_COMPROBANTE TC, FACTURAS_ORDEN_PAGO FP,"
    sql = sql & " NOTA_DEBITO_PROVEEDOR F"
    sql = sql & " WHERE OP.OPG_NUMERO=" & XN(OrdPago)
    sql = sql & " AND OP.OPG_FECHA=" & XDQ(Fecha)
    sql = sql & " AND OP.OPG_NUMERO=FP.OPG_NUMERO"
    sql = sql & " AND OP.OPG_FECHA=FP.OPG_FECHA"
    sql = sql & " AND OP.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND OP.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND FP.FPR_TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND P.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND P.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND P.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND P.IVA_CODIGO=C.IVA_CODIGO"
    sql = sql & " AND FP.FPR_TCO_CODIGO=F.TCO_CODIGO"
    sql = sql & " AND FP.FPR_NROSUC=F.DPR_NROSUC"
    sql = sql & " AND FP.FPR_NUMERO=F.DPR_NUMERO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_ORDEN_PAGO ("
            sql = sql & "OPG_NUMERO,OPG_FECHA,PROV_RAZSOC,PROV_DOMICI,PROV_CUIT,PROV_INGBRU,"
            sql = sql & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            sql = sql & "OPG_TOTAL,FAC_ABREVIA,FAC_NUMERO,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL) VALUES ("
            sql = sql & XS(Format(txtNroOrdenPago, "00000000")) & ","
            sql = sql & XDQ(FechaOrdenPago) & ","
            sql = sql & XS(Rec1!PROV_RAZSOC) & ","
            sql = sql & XS(Rec1!PROV_DOMICI) & ","
            sql = sql & XS(Format(Rec1!PROV_CUIT, "##-########-#")) & ","
            sql = sql & XS(Format(Rec1!PROV_INGBRU, "###-#####-##")) & ","
            sql = sql & XS(Rec1!LOC_DESCRI) & ","
            sql = sql & XS(Rec1!PRO_DESCRI) & ","
            sql = sql & XS(Rec1!IVA_DESCRI) & ","
            sql = sql & "NULL,"
            sql = sql & "NULL,"
            sql = sql & "NULL,"
            sql = sql & "NULL,"
            sql = sql & XN(Rec1!OPG_TOTAL) & ","
            sql = sql & XS(Rec1!TCO_ABREVIA) & ","
            sql = sql & XS(Format(Rec1!DPR_NROSUC, "0000") & "-" & Format(Rec1!DPR_NUMERO, "00000000")) & ","
            sql = sql & XS(Rec1!DPR_FECHA) & ","
            sql = sql & XN(Rec1!OPG_IMPORTE) & ","
            sql = sql & XN(Rec1!DPR_TOTAL) & ")"
            DBConn.Execute sql
            
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub CmdNuevo_Click()
    Estado = 1
    cmdGrabar.Enabled = True
    txtNroOrdenPago.Enabled = True
    FrameRecibo.Enabled = True
    FrameProveedor.Enabled = True
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
    
    txtNroOrdenPago.Text = ""
    FechaOrdenPago.Value = Date
    txtTotalCheques.Text = ""
    txtTotalEfectivo.Text = ""
    txtTotalValores.Text = ""
    txtTotalAplicar.Text = ""
    txtTotalComprobante.Text = ""
    cboTipoProveedor.ListIndex = 0
    txtCodProveedor.Text = ""
    tabValores.Tab = 0
    tabComprobantes.Tab = 0
    
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRecibo) 'ESTADO PENDIENTE
    tabDatos.Tab = 0
    txtNroOrdenPago.SetFocus
End Sub

Private Sub cmdNuevoCheque_Click()
    If optChequeTercero.Value = True Then
        FrmCargaCheques.Show vbModal
        TxtCheNumero.SetFocus
    Else
        FrmCargaChequesPropios.Show vbModal
        TxtCheNumero.SetFocus
    End If
        
End Sub

Private Sub cmdQuitarComprobantes_Click()
    If GrillaAplicar1.Rows > 1 Then
        If MsgBox("¿Seguro que desea eliminar?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            If GrillaAplicar1.Rows > 2 Then
                GrillaAplicar1.RemoveItem GrillaAplicar1.RowSel
                txtTotalAplicar.Text = SumaGrilla(GrillaAplicar1, 1)
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
                grillaValores.RemoveItem grillaValores.RowSel
                txtTotalValores.Text = SumaGrilla(grillaValores, 1)
            Else
                grillaValores.Rows = 1
                txtTotalValores.Text = ""
                grillaValores.HighLight = flexHighlightNever
            End If
        End If
    End If
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmPagoProveedores = Nothing
        Unload Me
    End If
End Sub


Private Sub Command1_Click()

End Sub

Private Sub FechaOrdenPago_LostFocus()
    If IsNull(FechaOrdenPago.Value) Then FechaOrdenPago.Value = Date
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
    FechaOrdenPago.Value = Date
    Call Centrar_pantalla(Me)
    tabDatos.Tab = 0
    tabValores.Tab = 0
    tabComprobantes.Tab = 0
    'CONFIGURO GRILLAS
    ConfiguroGrillas
    'CARGO COMBO CON LOS TIPOS DE PROVEEDORES
    LlenarComboTipoProv
    'CARGO COMBO CON LAS MONEDAS
    LLenarComboMoneda
    'CARGO COMBO CON LAS FACTURAS
    LlenarComboFactura
    'CARGO COMBO CON ORDEN DE PAGO
    LlenarComboOrdPag
    'CARGO COMBO BANCO
    CargoBanco
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRecibo) 'ESTADO PENDIENTE
    Estado = 1
    '------------------------
    frameBanco.Enabled = False
    cmdAgregarCheque.Enabled = False
    cmdAgregarEfectivo.Enabled = False
    txtNroOrdenPago.Enabled = True
    cmdAgregarFacturas.Enabled = False
    lblEstado.Caption = ""
    
    
       ' PROCESO PARA ACTUALIZAR SALDOS DE FACTURA DE PROVEEDOR PAGADAS
    'EN ALGUNA ORDEN DE PAGO
'    sql = "SELECT FPR_NUMERO,OPG_IMPORTE"
'    sql = sql & " FROM FACTURAS_ORDEN_PAGO"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While rec.EOF = False
        'sql = "UPDATE GASTOS_GENERALES SET "
        'sql = sql & " GGR_SALDO = 0"
        'sql = sql & " WHERE PROV_CODIGO IN (55)"
        'sql = sql & " AND FPR_FECHA <  " & XDQ("01/08/2009") & ""
        'DBConn.Execute sql
         
            
            'sql = "UPDATE FACTURA_PROVEEDOR SET "
            'sql = sql & " FPR_SALDO = 0"
            'sql = sql & " WHERE " 'PROV_CODIGO IN (55)"
            'sql = sql & " FPR_FECHA <=  " & XDQ("31/12/2008") & ""
            'DBConn.Execute sql
'            sql = "UPDATE NOTA_DEBITO_PROVEEDOR SET "
'            sql = sql & " DPR_SALDO = 0"
'            sql = sql & " WHERE " 'PROV_CODIGO IN (95)"
'            sql = sql & " DPR_FECHA <=  " & XDQ("31/12/2008") & ""
'            DBConn.Execute sql
'

'            'rec.MoveNext
'        Loop
'    End If
'    rec.Close
End Sub

Private Sub CargoBanco()
    sql = "SELECT B.BAN_DESCRI, B.BAN_CODINT"
    sql = sql & " FROM BANCO B, CTA_BANCARIA CB"
    sql = sql & " WHERE B.BAN_CODINT=CB.BAN_CODINT"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboBanco.AddItem Trim(rec!BAN_DESCRI)
            cboBanco.ItemData(cboBanco.NewIndex) = Trim(rec!BAN_CODINT)
            rec.MoveNext
        Loop
        cboBanco.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub CargoCtaBancaria(Banco As String)
    Set Rec1 = New ADODB.Recordset
    cboCtaBancaria.Clear
    sql = "SELECT CTA_NROCTA FROM CTA_BANCARIA"
    sql = sql & " WHERE BAN_CODINT=" & XN(Banco)
    sql = sql & " AND CTA_FECCIE IS NULL"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
     Do While Rec1.EOF = False
         cboCtaBancaria.AddItem Trim(Rec1!CTA_NROCTA)
         'cboCtaBancaria.ItemData(cboCtaBancaria.NewIndex) = Trim(Rec1!CTA_NROCTA)
         Rec1.MoveNext
     Loop
     cboCtaBancaria.ListIndex = 0
    End If
    Rec1.Close
End Sub

Private Sub LlenarComboOrdPag()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_CODIGO = 17"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboOrdPag.AddItem rec!TCO_ABREVIA
            cboOrdPag.ItemData(cboOrdPag.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboOrdPag.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboTipoProv()
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboTipoProveedor.AddItem "TODOS"
        Do While rec.EOF = False
            cboTipoProveedor.AddItem rec!TPR_DESCRI
            cboTipoProveedor.ItemData(cboTipoProveedor.NewIndex) = rec!TPR_CODIGO
            cboBuscaTipoProveedor.AddItem rec!TPR_DESCRI
            cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.NewIndex) = rec!TPR_CODIGO
            rec.MoveNext
        Loop
        cboTipoProveedor.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboFactura()
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
    GrillaCheques.FormatString = "^Bco|^Loc|^Suc|^Cod|^Nro Cheque" _
                               & "|^Fec Vto|>Importe|COD INTERNO BANCO|DECRI BANCO|Cheques propios"
    GrillaCheques.ColWidth(0) = 500   'BCO
    GrillaCheques.ColWidth(1) = 500   'LOC
    GrillaCheques.ColWidth(2) = 500   'SUC
    GrillaCheques.ColWidth(3) = 700   'COD
    GrillaCheques.ColWidth(4) = 1100  'NRO CHEQUE
    GrillaCheques.ColWidth(5) = 1000  'FEC VTO
    GrillaCheques.ColWidth(6) = 1000  'IMPORTE
    GrillaCheques.ColWidth(7) = 0     'COD INTERNO BANCO
    GrillaCheques.ColWidth(8) = 0     'DESCRI BANCO
    GrillaCheques.ColWidth(9) = 0     'CHEQUES PROPIOS
    GrillaCheques.Rows = 1
    'GRILLA EFECTIVO
    GrillaEfectivo.FormatString = "Moneda|>Importe|codigo moneda"
    GrillaEfectivo.ColWidth(0) = 1900 'MONEDA
    GrillaEfectivo.ColWidth(1) = 1000 'IMPORTE
    GrillaEfectivo.ColWidth(2) = 0    'CODIGO MONEDA
    GrillaEfectivo.Rows = 1
    'GRILLA Aplicar A
    GrillaAplicar.FormatString = "Comp.|>Número|^Fecha|>Total|>Saldo|codigo comprobante|SUCURSAL|NUMERO"
    GrillaAplicar.ColWidth(0) = 750 'COMPROBANTE
    GrillaAplicar.ColWidth(1) = 1250 'NUMERO
    GrillaAplicar.ColWidth(2) = 1000 'FECHA
    GrillaAplicar.ColWidth(3) = 1000 'TOTAL
    GrillaAplicar.ColWidth(4) = 1000 'SALDO
    GrillaAplicar.ColWidth(5) = 0    'CODIGO COMPROBANTE
    GrillaAplicar.ColWidth(6) = 0    'SUCURSAL COMP
    GrillaAplicar.ColWidth(7) = 0    'NUMERO COMP
    GrillaAplicar.Rows = 1
    
    'GRILLA BUSQUEDA
    GrdModulos.FormatString = "Nro Ord Pago|^Fecha Ord Pago|Tipo Proveedor|Proveedor"
    GrdModulos.ColWidth(0) = 1200 'NRO ORDEN PAGO
    GrdModulos.ColWidth(1) = 1300 'FECHA ORDEN PAGO
    GrdModulos.ColWidth(2) = 3000 'TIPO PROVEEDOR
    GrdModulos.ColWidth(3) = 5000 'PROVEEDOR
    GrdModulos.Rows = 1
    'grilla valores
    grillaValores.FormatString = "|>Importe|^Fecha|Descripción|>Número|codigo|CTA_BANCARIA"
    grillaValores.ColWidth(0) = 650  'TIPO DE VALOR (CHE,EFT...)
    grillaValores.ColWidth(1) = 1000 'IMPORTE
    grillaValores.ColWidth(2) = 1000 'FECHA
    grillaValores.ColWidth(3) = 2500 'DESCRIPCION
    grillaValores.ColWidth(4) = 1000 'NUMERO
    grillaValores.ColWidth(5) = 0    'CODIGO
    grillaValores.ColWidth(6) = 0    'CTA_BANCARIA
    grillaValores.Rows = 1
    'grilla aplicra a 1
    GrillaAplicar1.FormatString = "Comp.|>Importe|>Número|^Fecha|>Saldo|codigo comprobante|SUCURSAL|NUMNERO"
    GrillaAplicar1.ColWidth(0) = 1000 'COMPROBANTE
    GrillaAplicar1.ColWidth(1) = 1000 'IMPORTE
    GrillaAplicar1.ColWidth(2) = 1250 'NUMERO
    GrillaAplicar1.ColWidth(3) = 1000 'FECHA
    GrillaAplicar1.ColWidth(4) = 1000 'SALDO
    GrillaAplicar1.ColWidth(5) = 0    'CODIGO COMPROBANTE
    GrillaAplicar1.ColWidth(6) = 0    'SUCURSAL COMP
    GrillaAplicar1.ColWidth(7) = 0    'NUMERO COMP
    GrillaAplicar1.Rows = 1
    'grilla COMPROBANTES
    GrillaComp.FormatString = "Comprobante|>Número|^Fecha|>Importe|codigo comprobante"
    GrillaComp.ColWidth(0) = 1500 'COMPROBANTE
    GrillaComp.ColWidth(1) = 1000 'NUMERO
    GrillaComp.ColWidth(2) = 1000 'FECHA
    GrillaComp.ColWidth(3) = 1000 'IMPORTE
    GrillaComp.ColWidth(4) = 0    'CODIGO COMPROBANTE
    GrillaComp.Rows = 1
    'GRILLA AFAVOR
    GrillaAFavor.FormatString = "Comprobante|>Número|^Fecha|>Total|>Saldo|codigo comprobante"
    GrillaAFavor.ColWidth(0) = 1200 'COMPROBANTE
    GrillaAFavor.ColWidth(1) = 900  'NUMERO
    GrillaAFavor.ColWidth(2) = 1000 'FECHA
    GrillaAFavor.ColWidth(3) = 1000 'TOTAL
    GrillaAFavor.ColWidth(4) = 1000 'SALDO
    GrillaAFavor.ColWidth(5) = 0    'CODIGO COMPROBANTE
    GrillaAFavor.Rows = 1
    GrillaAFavor.HighLight = flexHighlightNever
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

Private Sub GrdModulos_DblClick()
     If GrdModulos.Rows > 1 Then
        CmdNuevo_Click
        txtNroOrdenPago.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        FechaOrdenPago.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
        tabDatos.Tab = 0
        txtNroOrdenPago_LostFocus
     End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub GrillaAFavor_Click()
    If GrillaAFavor.Rows > 1 Then
        txtSaldoACta.Text = Valido_Importe(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4))
        txtImporteACta.Text = Valido_Importe(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4))
        txtImporteACta.SetFocus
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

Private Sub optChequePropio_Click()
    If optChequePropio.Value = True Then
        cboBanco.Visible = True
        cboCtaBancaria.Visible = True
        frameBanco.Visible = False
    End If
End Sub

Private Sub optChequeTercero_Click()
    If optChequeTercero.Value = True Then
        frameBanco.Visible = True
        cboBanco.Visible = False
        cboCtaBancaria.Visible = False
    End If
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
    'cboBuscaTipoProveedor.Enabled = False
    'FechaDesde.Enabled = False
    'FechaHasta.Enabled = False
    'txtProveedor.Enabled = False
    'cmdBuscarProveedor.Enabled = False
    'CmdGrabar.Enabled = False
    If Me.Visible = True Then chkTipoProveedor.SetFocus
  End If
End Sub

Private Sub LimpiarBusqueda()
    FechaDesde.Value = Null
    FechaHasta.Value = Null
    txtProveedor.Text = ""
    cboTipoProveedor.ListIndex = -1
    GrdModulos.Rows = 1
    chkTipoProveedor.Value = Unchecked
    chkFecha.Value = Unchecked
    chkProveedor.Value = Unchecked
End Sub

Private Sub tabValores_Click(PreviousTab As Integer)
    If tabValores.Tab = 0 Then
        If Me.tabValores.Visible = True Then cmdAgregarCHE.SetFocus
    End If
    If tabValores.Tab = 1 Then
        optChequeTercero.Value = True
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
        If optChequeTercero.Value = True Then
            frameBanco.Enabled = True
        Else
            cboBanco.Enabled = True
            cboCtaBancaria.Enabled = True
        End If
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
    If optChequeTercero.Value = True Then
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
                TxtBanDescri.Text = rec!BAN_NOMCOR
                TxtCodInt.Text = rec!BAN_CODINT
            End If
            rec.Close
        End If
    Else ' CHEQUE PROPIO
        If TxtCheNumero.Text <> "" Then
            If Len(TxtCheNumero.Text) < 8 Then TxtCheNumero.Text = CompletarConCeros(TxtCheNumero.Text, 8)
        'sql = "SELECT * FROM CHEQUE WHERE "
            sql = "SELECT DISTINCT CE.CHEP_NUMERO, CH.CHEP_IMPORT, CH.CHEP_FECVTO, CE.BAN_CODINT, B.BAN_BANCO, B.BAN_LOCALIDAD,"
            sql = sql & " B.BAN_SUCURSAL, B.BAN_CODIGO, B.BAN_NOMCOR,CE.CPES_DESCRI,B.BAN_DESCRI,CH.CTA_NROCTA"
            sql = sql & " FROM CHEQUE_PROPIO_ESTADO CE, CHEQUE_PROPIO CH, BANCO B,ESTADO_CHEQUE E"
            sql = sql & " Where "
            sql = sql & " CE.CHEP_NUMERO = CH.CHEP_NUMERO And "
            sql = sql & " CE.BAN_CODINT = CH.BAN_CODINT And "
            sql = sql & " CH.BAN_CODINT=B.BAN_CODINT  "
            'sql = sql & " CE.ECH_CODIGO= E.ECH_CODIGO AND" '
            'sql = sql & " E.ECH_CODIGO=7" ' 7-entregado
            sql = sql & " AND CH.CHEP_NUMERO LIKE '%" & Trim(TxtCheNumero) & "%'"  'CODIGO (1) ES CHEQUE EN CARTERA
            
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                TxtCheNumero.Text = rec!CHEP_NUMERO
                Call BuscaCodigoProxItemData(rec!BAN_CODINT, cboBanco)
                Call CargoCtaBancaria(rec!BAN_CODINT)
                
                'Call BuscaCodigoProx(rec!CTA_NROCTA, cboCtaBancaria)
                'TxtBanco.Text = rec!BAN_BANCO
                'txtlocalidad.Text = rec!BAN_LOCALIDAD
                'TxtSucursal.Text = rec!BAN_SUCURSAL
                'TxtCodigo.Text = rec!BAN_CODIGO
                TxtCheImport.Text = rec!CHEP_IMPORT
                
                'TxtCheFecVto.value = rec!CHE_FECVTO
                'TxtBanDescri.Text = rec!BAN_NOMCOR
                'TxtCodInt.Text = rec!BAN_CODINT
            End If
            rec.Close
        End If
    End If
End Sub
'Private Sub BuscoCtaBancaria(Banco As String)
'    Set Rec1 = New ADODB.Recordset
'    cboCtaBancaria.Clear
'    sql = "SELECT CTA_NROCTA FROM CTA_BANCARIA"
'    sql = sql & " WHERE BAN_CODINT=" & XN(Banco)
'    sql = sql & " AND CTA_FECCIE IS NULL"
'    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec1.EOF = False Then
'     Do While Rec1.EOF = False
'         cboCtaBancaria.AddItem Trim(Rec1!CTA_NROCTA)
'         Rec1.MoveNext
'     Loop
'     cboCtaBancaria.ListIndex = 0
'    End If
'    Rec1.Close
'End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto txtcodigo
End Sub

Private Sub txtCodProveedor_Change()
    If txtCodProveedor.Text = "" Then
        txtProvRazSoc.Text = ""
        txtCliLocalidad.Text = ""
        txtDomici.Text = ""
        txtcodcli.Text = ""
    End If
End Sub

Private Sub txtCodProveedor_GotFocus()
    SelecTexto txtCodProveedor
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodProveedor_LostFocus()
    If txtCodProveedor.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        Rec1.Open BuscoProveedor(txtCodProveedor), DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtProvRazSoc.Text = Rec1!PROV_RAZSOC
            txtCliLocalidad.Text = Rec1!LOC_DESCRI
            txtDomici.Text = Rec1!PROV_DOMICI
            'busca el cod cliente asociado al proveedor
            If txtCodProveedor.Text <> "" Then
                Codcliente XN(txtCodProveedor.Text)
            End If
            Call BuscaCodigoProxItemData(CInt(Rec1!TPR_CODIGO), cboTipoProveedor)
            If Estado = 1 Then
                If BuscarFactura(CStr(cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)), txtCodProveedor) = False Then
                    MsgBox "No hay Facturas pendiente de Pago para el Proveedor", vbExclamation, TIT_MSGBOX
                    txtCodProveedor.Text = ""
                    txtCodProveedor.SetFocus
                Else
                    tabComprobantes.Tab = 1
                    Call BuscarSaldosAFavor(CStr(cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)), txtCodProveedor)
                End If
            End If
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
            cboTipoProveedor.ListIndex = 0
        End If
        Rec1.Close

    End If
End Sub
Private Function Codcliente(Proveedor As Integer)
    sql = "SELECT CLI_CODIGO FROM CLIENTE WHERE PROV_CODIGO = " & Proveedor
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtcodcli.Text = IIf(IsNull(rec!CLI_CODIGO), "", rec!CLI_CODIGO)
    Else
        txtcodcli.Text = ""
    End If
    rec.Close
End Function
Private Sub BuscarSaldosAFavor(CodTipoProv As String, CodProv As String)
        GrillaAFavor.Rows = 1
        Set rec = New ADODB.Recordset
        sql = "SELECT OS.OPG_NUMERO, OS.OPG_FECHA, OS.OPG_TOTSALDO"
        sql = sql & " ,OS.OPG_SALDO"
        sql = sql & " FROM ORDEN_PAGO_SALDO OS, ORDEN_PAGO O"
        sql = sql & " WHERE "
        sql = sql & " OS.OPG_NUMERO=O.OPG_NUMERO"
        sql = sql & " AND OS.OPG_FECHA=O.OPG_FECHA"
        sql = sql & " AND OS.OPG_SALDO > 0"
        sql = sql & " AND O.TPR_CODIGO=" & XN(CodTipoProv)
        sql = sql & " AND O.PROV_CODIGO=" & XN(CodProv)
        sql = sql & " ORDER BY OS.OPG_NUMERO, OS.OPG_FECHA"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            GrillaAFavor.HighLight = flexHighlightAlways
            Do While rec.EOF = False
                If rec!OPG_SALDO > 0 Then
                    GrillaAFavor.AddItem "ORD-PAG" & Chr(9) & rec!OPG_NUMERO _
                                    & Chr(9) & rec!OPG_FECHA & Chr(9) & Valido_Importe(rec!OPG_TOTSALDO) _
                                    & Chr(9) & Valido_Importe(rec!OPG_SALDO) & Chr(9) & CStr(cboOrdPag.ItemData(cboOrdPag.ListIndex))
                End If
                rec.MoveNext
            Loop
        End If
        rec.Close
        'buscar facturas de ventas del proveedor (comisiones)
        If txtcodcli.Text <> "" Then
            sql = "SELECT FCL_NUMERO, FCL_SUCURSAL, FCL_FECHA, FCL_TOTAL, FCL_SALDO"
            sql = sql & " ,TCO_CODIGO, TCO_ABREVIA"
            sql = sql & " FROM SALDO_FACTURAS_CLIENTE_V"
            sql = sql & " WHERE "
            sql = sql & " CLI_CODIGO=" & XN(txtcodcli.Text)
            sql = sql & " ORDER BY FCL_FECHA DESC"
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
            If rec.EOF = False Then
                Do While rec.EOF = False
                    If rec!FCL_SALDO > 0 Then
                        GrillaAFavor.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") _
                                        & Chr(9) & rec!FCL_FECHA & Chr(9) & Valido_Importe(rec!FCL_TOTAL) _
                                        & Chr(9) & Valido_Importe(rec!FCL_SALDO) & Chr(9) & rec!TCO_CODIGO
                    End If
                    rec.MoveNext
                Loop
                'BuscarFactura = True
            Else
                'BuscarFactura = False
            End If
            rec.Close
        End If
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
    If txtImporteComprobante.Text <> "" Then txtImporteComprobante = Valido_Importe(txtImporteComprobante)
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
        fechaComprobantes.Value = Null
        txtImporteComprobante.Enabled = False
        cmdAgregarComprobante.Enabled = False
    Else
        txtImporteComprobante.Enabled = True
        cmdAgregarComprobante.Enabled = True
    End If
End Sub

Private Function BuscarFactura(CodTipoProv As String, CodProv As String) As Boolean
        GrillaAplicar.Rows = 1
        
        Set rec = New ADODB.Recordset
        sql = "SELECT FPR_NROSUC,FPR_NUMERO,FPR_FECHA, FPR_TOTAL, FPR_SALDO"
        sql = sql & " ,TCO_CODIGO, TCO_ABREVIA"
        sql = sql & " FROM SALDO_FACTURAS_PROVEEDOR_V"
        sql = sql & " WHERE "
        sql = sql & " TPR_CODIGO=" & XN(CodTipoProv)
        sql = sql & " AND PROV_CODIGO=" & XN(CodProv)
        sql = sql & " AND FPR_FECHA > 0"
        sql = sql & " ORDER BY FPR_FECHA DESC"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            Do While rec.EOF = False
                If rec!FPR_SALDO > 0 Then
                    GrillaAplicar.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FPR_NROSUC, "0000") & "-" & Format(rec!FPR_NUMERO, "00000000") _
                                    & Chr(9) & rec!FPR_FECHA & Chr(9) & Valido_Importe(rec!FPR_TOTAL) _
                                    & Chr(9) & Valido_Importe(rec!FPR_SALDO) & Chr(9) & rec!TCO_CODIGO _
                                    & Chr(9) & rec!FPR_NROSUC & Chr(9) & rec!FPR_NUMERO
                End If
                rec.MoveNext
            Loop
            GrillaAplicar.HighLight = flexHighlightAlways
            BuscarFactura = True
        Else
            If GrillaAplicar.Rows = 1 Then
                BuscarFactura = False
            End If
        End If
        rec.Close
        
        'BUSCAR NOTA DE DEBITO
        sql = "SELECT ND.DPR_NROSUC,ND.DPR_NUMERO,ND.DPR_FECHA, ND.DPR_TOTAL, ND.DPR_SALDO"
        sql = sql & " ,TC.TCO_CODIGO, TC.TCO_ABREVIA"
        sql = sql & " FROM NOTA_DEBITO_PROVEEDOR ND, TIPO_COMPROBANTE TC"
        sql = sql & " WHERE "
        sql = sql & " ND.TCO_CODIGO = TC.TCO_CODIGO"
        sql = sql & " and ND.TPR_CODIGO =" & XN(CodTipoProv)
        sql = sql & " AND ND.PROV_CODIGO =" & XN(CodProv)
        sql = sql & " AND ND.DPR_SALDO > 0 "
        sql = sql & " ORDER BY ND.DPR_FECHA DESC"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            Do While rec.EOF = False
                If rec!DPR_SALDO > 0 Then
                    GrillaAplicar.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!DPR_NROSUC, "0000") & "-" & Format(rec!DPR_NUMERO, "00000000") _
                                    & Chr(9) & rec!DPR_FECHA & Chr(9) & Valido_Importe(rec!DPR_TOTAL) _
                                    & Chr(9) & Valido_Importe(rec!DPR_SALDO) & Chr(9) & rec!TCO_CODIGO _
                                    & Chr(9) & rec!DPR_NROSUC & Chr(9) & rec!DPR_NUMERO
                End If
                rec.MoveNext
            Loop
            GrillaAplicar.HighLight = flexHighlightAlways
            BuscarFactura = True
        Else
            If GrillaAplicar.Rows = 1 Then
                BuscarFactura = False
            End If
        End If
        rec.Close
        
        'BUSCAR GASTOS GENERALES
        sql = "SELECT GG.GGR_NROSUC,GG.GGR_NROCOMP,GG.GGR_FECHACOMP, GG.GGR_TOTAL, GG.GGR_SALDO"
        sql = sql & " ,TC.TCO_CODIGO, TC.TCO_ABREVIA"
        sql = sql & " FROM GASTOS_GENERALES GG, TIPO_COMPROBANTE TC"
        sql = sql & " WHERE "
        sql = sql & " GG.TCO_CODIGO = TC.TCO_CODIGO"
        sql = sql & " and GG.TPR_CODIGO =" & XN(CodTipoProv)
        sql = sql & " AND GG.PROV_CODIGO =" & XN(CodProv)
        sql = sql & " AND GG.GGR_SALDO > 0 "
        sql = sql & " ORDER BY GG.GGR_FECHACOMP"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            Do While rec.EOF = False
                If rec!GGR_SALDO > 0 Then
                    GrillaAplicar.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!GGR_NROSUC, "0000") & "-" & Format(rec!GGR_NROCOMP, "00000000") _
                                    & Chr(9) & rec!GGR_FECHACOMP & Chr(9) & Valido_Importe(rec!GGR_TOTAL) _
                                    & Chr(9) & Valido_Importe(rec!GGR_SALDO) & Chr(9) & rec!TCO_CODIGO _
                                    & Chr(9) & rec!GGR_NROSUC & Chr(9) & rec!GGR_NROCOMP
                End If
                rec.MoveNext
            Loop
            GrillaAplicar.HighLight = flexHighlightAlways
            BuscarFactura = True
        Else
            If GrillaAplicar.Rows = 1 Then
                BuscarFactura = False
            End If
        End If
        rec.Close
        
        
    
End Function

Private Sub txtNroComprobantes_LostFocus()
    If txtNroComprobantes.Text <> "" Then
        Select Case cboComprobantes.ItemData(cboComprobantes.ListIndex)
            Case 4, 5, 6
                'Call BuscarNotaCredito
            Case Else
                If BuscoComprobanteEnRecibo = False Then
                    MsgBox "El comprobante de Retención ya fue cargado a un Recibo", vbInformation, TIT_MSGBOX
                    txtNroComprobantes.Text = ""
                    txtNroComprobantes.SetFocus
                End If
        End Select
    End If
End Sub

Private Function BuscoComprobanteEnRecibo() As Boolean
    Set Rec2 = New ADODB.Recordset
    
    sql = "SELECT OPG_NUMERO"
    sql = sql & " FROM DETALLE_ORDEN_PAGO"
    sql = sql & " WHERE"
    sql = sql & " DOP_TCO_CODIGO=" & cboComprobantes.ItemData(cboComprobantes.ListIndex)
    If fechaComprobantes.Value <> "" Then
        sql = sql & " AND DOP_COMFECHA=" & XDQ(fechaComprobantes.Value)
    End If
    sql = sql & " AND DOP_COMNUMERO=" & XN(txtNroComprobantes)
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Rec2.EOF = False Then
        BuscoComprobanteEnRecibo = False
    Else
        BuscoComprobanteEnRecibo = True
    End If
    Rec2.Close
End Function

Private Sub txtNroOrdenPago_GotFocus()
    SelecTexto txtNroOrdenPago
End Sub

Private Sub txtNroOrdenPago_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroOrdenPago_LostFocus()
       
    If ActiveControl.Name = "cmdNuevo" Or ActiveControl.Name = "CmdSalir" Then Exit Sub
    If txtNroOrdenPago.Text <> "" Then
        Set Rec2 = New ADODB.Recordset
        sql = "SELECT * FROM ORDEN_PAGO"
        sql = sql & " WHERE"
        sql = sql & " OPG_NUMERO=" & XN(txtNroOrdenPago)
        If FechaOrdenPago.Value <> "" Then
            sql = sql & " AND OPG_FECHA=" & XDQ(FechaOrdenPago)
        End If
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec2.EOF = False Then
            If Rec2.RecordCount > 2 Then
                Rec2.Close
                tabDatos.Tab = 1
                Exit Sub
            End If
            'CABEZA DE LA ORDEN DE PAGO
            FechaOrdenPago.Value = Rec2!OPG_FECHA
            'CARGO ESTADO
            Call BuscoEstado(CInt(Rec2!EST_CODIGO), lblEstadoRecibo)
            Estado = CInt(Rec2!EST_CODIGO)
            Call BuscaCodigoProxItemData(CInt(Rec2!TPR_CODIGO), cboTipoProveedor)
            txtCodProveedor.Text = Rec2!PROV_CODIGO
            txtCodProveedor_LostFocus
            
            'DETALLE_DE LA ORDEN DE PAGO
            Set rec = New ADODB.Recordset
            sql = "SELECT *"
            sql = sql & " FROM DETALLE_ORDEN_PAGO"
            sql = sql & " WHERE"
            sql = sql & " OPG_NUMERO=" & XN(txtNroOrdenPago)
            sql = sql & " AND OPG_FECHA=" & XDQ(FechaOrdenPago)
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic

            If rec.EOF = False Then
                Do While rec.EOF = False
                    If Not IsNull(rec!BAN_CODINT) Then 'BANCO
                        Call BuscarCheque(rec!BAN_CODINT, rec!CHE_NUMERO, ChkNull(rec!CTA_NROCTA))
                    ElseIf Not IsNull(rec!MON_CODIGO) Then 'MONEDA
                        grillaValores.AddItem "EFT" & Chr(9) & Valido_Importe(rec!DOP_MONIMP) _
                                        & Chr(9) & "" & Chr(9) & BuscarMoneda(rec!MON_CODIGO) _
                                        & Chr(9) & "" & Chr(9) & rec!MON_CODIGO

                    ElseIf Not IsNull(rec!DOP_TCO_CODIGO) Then 'COMPROBANTE
                        grillaValores.AddItem "COMP" & Chr(9) & Valido_Importe(rec!DOP_COMIMP) _
                                        & Chr(9) & rec!DOP_COMFECHA & Chr(9) & BuscarTipoDocAbre(rec!DOP_TCO_CODIGO) _
                                        & Chr(9) & rec!DOP_COMNUMERO & Chr(9) & rec!DOP_TCO_CODIGO
                    End If
                    rec.MoveNext
                Loop

                grillaValores.HighLight = flexHighlightAlways
                txtTotalValores.Text = SumaGrilla(grillaValores, 1)
            End If
            rec.Close

            'DETALLE DE FACTURAS_ORDEN_PAGO
            sql = "SELECT *"
            sql = sql & " FROM FACTURAS_ORDEN_PAGO"
            sql = sql & " WHERE"
            sql = sql & " OPG_NUMERO=" & XN(txtNroOrdenPago)
            sql = sql & " AND OPG_FECHA=" & XDQ(FechaOrdenPago)
            
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic

            If rec.EOF = False Then
                Do While rec.EOF = False
                    GrillaAplicar1.AddItem BuscarTipoDocAbre(rec!FPR_TCO_CODIGO) & Chr(9) & Valido_Importe(rec!OPG_IMPORTE) & Chr(9) & _
                                Format(rec!FPR_NROSUC, "0000") & "-" & Format(rec!FPR_NUMERO, "00000000") & Chr(9) & rec!FPR_FECHA & Chr(9) & "" & Chr(9) & rec!FPR_TCO_CODIGO _
                                & Chr(9) & rec!FPR_NROSUC & Chr(9) & rec!FPR_NUMERO

                    rec.MoveNext
                Loop
                GrillaAplicar1.HighLight = flexHighlightAlways
                txtTotalAplicar.Text = SumaGrilla(GrillaAplicar1, 1)
            End If
            FrameRecibo.Enabled = False
            FrameProveedor.Enabled = False
            rec.Close
            cmdNuevo.SetFocus
            cmdGrabar.Enabled = False
        End If
        Rec2.Close
    Else  'SI NO INGRESO UN NUMERO BUSCO EL MAYOR
        sql = "SELECT MAX(OPG_NUMERO)+1 AS NUMERO FROM ORDEN_PAGO"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If IsNull(Rec2!Numero) Then
            txtNroOrdenPago.Text = 1
        Else
            txtNroOrdenPago.Text = Rec2!Numero
        End If
        Rec2.Close
    End If
End Sub

Private Function BuscarCheque(Codigo As String, NroChe As String, CtaBan As String) As String
    
    Set Rec1 = New ADODB.Recordset
    If CtaBan = "" Then
        sql = "SELECT B.BAN_DESCRI,C.CHE_IMPORT,C.CHE_FECVTO"
        sql = sql & " FROM BANCO B, CHEQUE C"
        sql = sql & " WHERE C.BAN_CODINT=" & XN(Codigo)
        sql = sql & " AND C.CHE_NUMERO=" & XS(NroChe)
        sql = sql & " AND C.BAN_CODINT=B.BAN_CODINT"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            grillaValores.AddItem "CHE-T" & Chr(9) & Valido_Importe(Rec1!che_import) & Chr(9) & Rec1!CHE_FECVTO _
                               & Chr(9) & Rec1!BAN_DESCRI & Chr(9) & NroChe & Chr(9) & Codigo
        End If
    Else
        sql = "SELECT B.BAN_DESCRI,C.CHEP_IMPORT,C.CHEP_FECVTO"
        sql = sql & " FROM BANCO B, CHEQUE_PROPIO C"
        sql = sql & " WHERE C.BAN_CODINT=" & XN(Codigo)
        sql = sql & " AND C.CHEP_NUMERO=" & XS(NroChe)
        sql = sql & " AND C.CTA_NROCTA=" & XS(CtaBan)
        sql = sql & " AND C.BAN_CODINT=B.BAN_CODINT"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            grillaValores.AddItem "CHE-P" & Chr(9) & Valido_Importe(Rec1!CHEP_IMPORT) & Chr(9) & Rec1!CHEP_FECVTO _
                               & Chr(9) & Rec1!BAN_DESCRI & Chr(9) & NroChe & Chr(9) & Codigo & Chr(9) & CtaBan
        End If
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

Private Sub txtProveedor_Change()
    If txtProveedor.Text = "" Then
        txtDesProv.Text = ""
    End If
End Sub

Private Sub txtProveedor_GotFocus()
    SelecTexto txtProveedor
End Sub

Private Sub txtProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtProveedor_LostFocus()
    Set Rec1 = New ADODB.Recordset
    If txtProveedor.Text <> "" Then
        sql = "SELECT TPR_CODIGO,PROV_CODIGO,PROV_RAZSOC"
        sql = sql & " FROM PROVEEDOR"
        sql = sql & " WHERE"
        sql = sql & " PROV_CODIGO=" & XN(txtProveedor)
        If chkTipoProveedor.Value = Checked Then
            sql = sql & " AND TPR_CODIGO=" & XN(cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.ListIndex))
        End If
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtDesProv.Text = Rec1!PROV_RAZSOC
            Call BuscaCodigoProxItemData(CInt(Rec1!TPR_CODIGO), cboBuscaTipoProveedor)
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtProveedor.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtProvRazSoc_Change()
    If txtProvRazSoc.Text = "" Then
        txtCodProveedor.Text = ""
        txtCliLocalidad.Text = ""
        txtDomici.Text = ""
    End If
End Sub

Private Sub txtProvRazSoc_GotFocus()
    SelecTexto txtProvRazSoc
End Sub

Private Sub txtProvRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtProvRazSoc_LostFocus()
    If ActiveControl.Name = "txtCodProveedor" Then Exit Sub
    If txtCodProveedor.Text = "" And txtProvRazSoc.Text <> "" Then
        Rec2.Open BuscoProveedor(txtProvRazSoc), DBConn, adOpenStatic, adLockOptimistic
        If Rec2.EOF = False Then
            If Rec2.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 5
                frmBuscar.TxtDescriB.Text = txtProvRazSoc.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 1
                    txtCodProveedor.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 2
                    txtProvRazSoc.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 3
                    Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboTipoProveedor)
                    txtCodProveedor_LostFocus
                    txtProvRazSoc.SetFocus
                Else
                    txtCodProveedor.SetFocus
                End If
            Else
                txtCodProveedor.Text = Rec2!PROV_CODIGO
                txtProvRazSoc.Text = Rec2!PROV_RAZSOC
                txtCodProveedor_LostFocus
            End If
        Else
            MsgBox "No se encontro el Proveedor", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        Rec2.Close
    ElseIf txtCodProveedor.Text = "" And txtProvRazSoc.Text = "" Then
        MsgBox "Debe elegir un Proveedor", vbExclamation, TIT_MSGBOX
        txtCodProveedor.SetFocus
    End If
End Sub

Private Function BuscoProveedor(Pro As String) As String
    sql = "SELECT DISTINCT PRO.TPR_CODIGO,PRO.PROV_CODIGO, PRO.PROV_RAZSOC,"
    sql = sql & " PRO.PROV_DOMICI, L.LOC_DESCRI"
    sql = sql & " FROM PROVEEDOR PRO,LOCALIDAD L, TIPO_PROVEEDOR TP"
    sql = sql & " WHERE"
    sql = sql & " PRO.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND TP.TPR_CODIGO=PRO.TPR_CODIGO"
    
    
    If txtCodProveedor.Text <> "" Then
        sql = sql & " AND PRO.PROV_CODIGO=" & XN(Pro)
    Else
        sql = sql & " AND PRO.PROV_RAZSOC LIKE '%" & Trim(Pro) & "%'"
    End If
    If cboTipoProveedor.List(cboTipoProveedor.ListIndex) <> "TODOS" Then
        sql = sql & " AND PRO.TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    End If
    

    BuscoProveedor = sql
End Function

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
        If rec.EOF = False Then 'EXITE
            'ME FIJO SI ESTA EN CARTERA
            Set Rec1 = New ADODB.Recordset
            sql = "SELECT ECH_CODIGO, ECH_DESCRI"
            sql = sql & " FROM ChequeEstadoVigente"
            sql = sql & " Where CHE_NUMERO = " & XS(TxtCheNumero.Text)
            sql = sql & " AND BAN_CODINT = " & XN(TxtCodInt.Text)
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                If Rec1!ECH_CODIGO <> 1 Then
                    MsgBox "El Cheque no puede ser utilizado por su estado: " & Rec1!ECH_DESCRI, vbCritical, TIT_MSGBOX
                    Rec1.Close
                    rec.Close
                    TxtCheNumero.Text = ""
                    TxtCheNumero.SetFocus
                    Exit Sub
                End If
            End If
            Rec1.Close
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

