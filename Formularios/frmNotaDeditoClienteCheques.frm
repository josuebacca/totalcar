VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNotaDeditoClienteCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de D�dito Clientes por Cheques..."
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   8535
      TabIndex        =   13
      Top             =   7185
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10305
      TabIndex        =   15
      Top             =   7185
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   7650
      TabIndex        =   12
      Top             =   7185
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   9420
      TabIndex        =   14
      Top             =   7185
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7125
      Left            =   60
      TabIndex        =   27
      Top             =   15
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   12568
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
      TabPicture(0)   =   "frmNotaDeditoClienteCheques.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameNotaDebito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FramePara"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmNotaDeditoClienteCheques.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame FramePara 
         Caption         =   "Para..."
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
         Left            =   4455
         TabIndex        =   43
         Top             =   375
         Width           =   6585
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
            Left            =   840
            TabIndex        =   71
            Top             =   1570
            Width           =   1455
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
            Left            =   2325
            TabIndex        =   70
            Top             =   1570
            Width           =   2895
         End
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
            Left            =   5250
            TabIndex        =   69
            Top             =   1570
            Width           =   1215
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   1860
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoClienteCheques.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Buscar Cliente"
            Top             =   300
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
            Left            =   840
            MaxLength       =   50
            TabIndex        =   63
            Top             =   645
            Width           =   4400
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
            Left            =   840
            TabIndex        =   62
            Top             =   1260
            Width           =   4400
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
            Left            =   840
            TabIndex        =   61
            Top             =   952
            Width           =   4400
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
            Left            =   2295
            MaxLength       =   50
            TabIndex        =   5
            Tag             =   "Descripci�n"
            Top             =   300
            Width           =   4170
         End
         Begin VB.TextBox txtCodCliente 
            Height          =   285
            Left            =   840
            MaxLength       =   40
            TabIndex        =   4
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   180
            TabIndex        =   73
            Top             =   1575
            Width           =   600
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos"
            Height          =   195
            Left            =   5430
            TabIndex        =   72
            Top             =   1320
            Width           =   765
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   45
            TabIndex        =   66
            Top             =   980
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   180
            Left            =   105
            TabIndex        =   65
            Top             =   670
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   75
            TabIndex        =   64
            Top             =   1305
            Width           =   705
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   255
            TabIndex        =   60
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.Frame FrameNotaDebito 
         Caption         =   "Nota de D�bito..."
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
         Left            =   105
         TabIndex        =   29
         Top             =   375
         Width           =   4350
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
            Left            =   1185
            MaxLength       =   4
            TabIndex        =   1
            Top             =   825
            Width           =   555
         End
         Begin VB.TextBox txtNroNotaDebito 
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
            Left            =   1770
            MaxLength       =   8
            TabIndex        =   2
            Top             =   825
            Width           =   1065
         End
         Begin VB.ComboBox cboNotaDebito 
            Height          =   315
            Left            =   1185
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   390
            Width           =   2400
         End
         Begin MSComCtl2.DTPicker FechaNotaDebito 
            Height          =   315
            Left            =   1200
            TabIndex        =   3
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   16842753
            CurrentDate     =   41098
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   765
            TabIndex        =   46
            Top             =   420
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   630
            TabIndex        =   44
            Top             =   1250
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "N�mero:"
            Height          =   195
            Left            =   525
            TabIndex        =   42
            Top             =   835
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   585
            TabIndex        =   41
            Top             =   1665
            Width           =   540
         End
         Begin VB.Label lblEstadoNotaDebito 
            AutoSize        =   -1  'True
            Caption         =   "EST. NOTA DEBITO"
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
            Left            =   1185
            TabIndex        =   40
            Top             =   1680
            Width           =   1755
         End
      End
      Begin VB.Frame frameBuscar 
         Caption         =   "Buscar Nota de D�dito por..."
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
         Left            =   -74610
         TabIndex        =   32
         Top             =   540
         Width           =   10410
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
            TabIndex        =   74
            Top             =   775
            Width           =   4620
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3360
            TabIndex        =   21
            Top             =   775
            Width           =   990
         End
         Begin VB.CommandButton cmdBuscarVen 
            Height          =   300
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoClienteCheques.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Buscar Vendedor"
            Top             =   775
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboNotaDebito1 
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1605
            Width           =   2400
         End
         Begin VB.CheckBox chkTipoFactura 
            Caption         =   "Tipo"
            Height          =   195
            Left            =   300
            TabIndex        =   19
            Top             =   1545
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4410
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoClienteCheques.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Buscar Cliente"
            Top             =   375
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   300
            TabIndex        =   17
            Top             =   845
            Width           =   1020
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1695
            Left            =   9660
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoClienteCheques.frx":0956
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Buscar "
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
            Left            =   4845
            MaxLength       =   50
            TabIndex        =   33
            Tag             =   "Descripci�n"
            Top             =   375
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3360
            MaxLength       =   40
            TabIndex        =   20
            Top             =   375
            Width           =   975
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   300
            TabIndex        =   18
            Top             =   1195
            Width           =   810
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   300
            TabIndex        =   16
            Top             =   495
            Width           =   855
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   3360
            TabIndex        =   22
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   16842753
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   5970
            TabIndex        =   23
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   16842753
            CurrentDate     =   41098
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   2910
            TabIndex        =   59
            Top             =   1650
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Index           =   0
            Left            =   2535
            TabIndex        =   37
            Top             =   830
            Width           =   735
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4935
            TabIndex        =   36
            Top             =   1245
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2265
            TabIndex        =   35
            Top             =   1240
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
            TabIndex        =   34
            Top             =   420
            Width           =   525
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4170
         Left            =   -74625
         TabIndex        =   26
         Top             =   2745
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7355
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame3 
         Height          =   4800
         Left            =   105
         TabIndex        =   30
         Top             =   2235
         Width           =   10950
         Begin VB.CommandButton cmdBuscarCheque 
            Height          =   330
            Left            =   10500
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaDeditoClienteCheques.frx":30F8
            Style           =   1  'Graphical
            TabIndex        =   78
            TabStop         =   0   'False
            ToolTipText     =   "Buscar Cheques"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdAgregarCheque 
            Height          =   330
            Left            =   10500
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaDeditoClienteCheques.frx":3402
            Style           =   1  'Graphical
            TabIndex        =   77
            TabStop         =   0   'False
            ToolTipText     =   "Agregar Cheque"
            Top             =   705
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdQuitarCheque 
            Height          =   330
            Left            =   10500
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaDeditoClienteCheques.frx":370C
            Style           =   1  'Graphical
            TabIndex        =   76
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Cheque"
            Top             =   1050
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CheckBox chkBonificaEnPesos 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en $"
            Height          =   285
            Left            =   390
            TabIndex        =   8
            Top             =   4035
            Width           =   1290
         End
         Begin VB.CheckBox chkBonificaEnPorsentaje 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en % "
            Height          =   285
            Left            =   390
            TabIndex        =   7
            Top             =   3735
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
            TabIndex        =   57
            Top             =   4065
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
            TabIndex        =   54
            Top             =   4065
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeIva 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6900
            TabIndex        =   10
            Top             =   3735
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
            TabIndex        =   51
            Top             =   4065
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeBoni 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2850
            TabIndex        =   9
            Top             =   3735
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
            TabIndex        =   48
            Top             =   4065
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
            TabIndex        =   47
            Top             =   3735
            Width           =   1350
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   1455
            MaxLength       =   60
            TabIndex        =   11
            Top             =   4410
            Width           =   8865
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   945
            TabIndex        =   31
            Top             =   480
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3555
            Left            =   75
            TabIndex        =   6
            Top             =   135
            Width           =   10365
            _ExtentX        =   18283
            _ExtentY        =   6271
            _Version        =   393216
            Rows            =   3
            Cols            =   12
            FixedCols       =   0
            BackColorSel    =   12648447
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   1
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   4110
            TabIndex        =   58
            Top             =   4125
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   6270
            TabIndex        =   56
            Top             =   4110
            Width           =   570
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "% I.V.A.:"
            Height          =   195
            Left            =   6240
            TabIndex        =   55
            Top             =   3765
            Width           =   600
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   2235
            TabIndex        =   53
            Top             =   4110
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Bonificaci�n:"
            Height          =   195
            Left            =   1890
            TabIndex        =   52
            Top             =   3765
            Width           =   915
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   8505
            TabIndex        =   50
            Top             =   4110
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   8175
            TabIndex        =   49
            Top             =   3765
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   210
            TabIndex        =   45
            Top             =   4455
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Nota de D�bito"
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
      Index           =   1
      Left            =   3960
      TabIndex        =   75
      Top             =   7320
      Width           =   2925
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
      TabIndex        =   39
      Top             =   7260
      Width           =   750
   End
End
Attribute VB_Name = "frmNotaDeditoClienteCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim w As Integer
Dim VBonificacion As Double
Dim VTotal As Double
Dim VEstadoNotaDebito As Integer
Dim GTotAdm As Double ' variable gloabl en el modulo q uso para guardar total de gastos administrativos



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


Private Sub chkTipoFactura_Click()
    If chkTipoFactura.Value = Checked Then
        cboNotaDebito1.Enabled = True
        cboNotaDebito1.ListIndex = 0
    Else
        cboNotaDebito1.Enabled = False
        cboNotaDebito1.ListIndex = -1
    End If
End Sub

Private Sub chkTipoFactura_LostFocus()
    If chkTipoFactura.Value = Checked And chkCliente.Value = Unchecked _
        And chkVendedor.Value = Unchecked _
        And chkFecha.Value = Unchecked Then cboNotaDebito1.SetFocus
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

Private Sub cmdAgregarCheque_Click()
    FrmCargaCheques.Show vbModal
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
            
    sql = "SELECT ND.*, C.CLI_RAZSOC , TC.TCO_ABREVIA"
    sql = sql & " FROM NOTA_DEBITO_CLIENTE ND,"
    sql = sql & " TIPO_COMPROBANTE TC , CLIENTE C"
    sql = sql & " WHERE ND.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND ND.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND ND.NDC_SERVICHEQUE='C'" 'PARA QUE BUSQUE CHEQUES
    If txtCliente.Text <> "" Then sql = sql & " AND ND.CLI_CODIGO=" & XN(txtCliente)
    If Not IsNull(FechaDesde) Then sql = sql & " AND ND.NDC_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND ND.NDC_FECHA<=" & XDQ(FechaHasta)
    If chkTipoFactura.Value = Checked Then sql = sql & " AND ND.TCO_CODIGO=" & XN(cboNotaDebito1.ItemData(cboNotaDebito1.ListIndex))
    sql = sql & " ORDER BY ND.NDC_SUCURSAL,ND.NDC_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!NDC_SUCURSAL, "0000") & "-" & Format(rec!NDC_NUMERO, "00000000") _
                            & Chr(9) & rec!NDC_FECHA & Chr(9) & Format(rec!NDC_TOTAL, "#0.00") & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!EST_CODIGO _
                            & Chr(9) & rec!NDC_BONIFICA & Chr(9) & rec!NDC_IVA & Chr(9) & rec!NDC_OBSERVACION _
                            & Chr(9) & rec!TCO_CODIGO & Chr(9) & rec!NDC_BONIPESOS _
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

Private Sub cmdBuscarCheque_Click()
    Dim codint As Integer
    frmBuscar.TipoBusqueda = 6
    frmBuscar.Show vbModal
    'TxtCheNumero.Text = frmBuscar.grdBuscar.Col
    frmBuscar.TxtDescriB.Text = ""
    grdGrilla.Col = 0
    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
    txtEdit.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
    TxtEdit_KeyDown 13, 0
    
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

Private Sub cmdGrabar_Click()
    Dim VStockFisico As String
    
    If ValidarNotaBebito = False Then Exit Sub
    If MsgBox("�Confirma Nota de D�bito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayErrorFactura
    
    DBConn.BeginTrans
    sql = "SELECT * FROM NOTA_DEBITO_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & XN(cboNotaDebito.ItemData(cboNotaDebito.ListIndex))
    sql = sql & " AND NDC_NUMERO= " & XN(txtNroNotaDebito)
    sql = sql & " AND NDC_SUCURSAL=" & XN(txtNroSucursal)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = True Then
        sql = "INSERT INTO NOTA_DEBITO_CLIENTE"
        sql = sql & " (TCO_CODIGO, NDC_NUMERO, NDC_SUCURSAL, NDC_FECHA,"
        sql = sql & " CLI_CODIGO, NDC_BONIFICA, NDC_IVA, NDC_SERVICHEQUE,"
        sql = sql & " NDC_OBSERVACION,NDC_NUMEROTXT,NDC_SUBTOTAL,NDC_TOTAL,NDC_BONIPESOS, EST_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(cboNotaDebito.ItemData(cboNotaDebito.ListIndex)) & ","
        sql = sql & XN(txtNroNotaDebito) & ","
        sql = sql & XN(txtNroSucursal) & ","
        sql = sql & XDQ(FechaNotaDebito) & ","
        sql = sql & XN(txtCodCliente) & ","
        sql = sql & XN(txtPorcentajeBoni) & ","
        sql = sql & XN(txtPorcentajeIva) & ","
        sql = sql & "'C'" & "," 'SE TRATA CHEQUES
        sql = sql & XS(txtObservaciones) & ","
        sql = sql & XS(Format(txtNroNotaDebito.Text, "00000000")) & ","
        If txtSubTotalBoni.Text <> "" Then 'SUBTOTAL
            sql = sql & XN(txtSubTotalBoni) & ","
        Else
            sql = sql & XN(txtSubtotal) & ","
        End If
        sql = sql & XN(txtTotal) & ","
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
                sql = "INSERT INTO DETALLE_NOTA_DEBITO_CLIENTE"
                sql = sql & " (TCO_CODIGO, NDC_NUMERO, NDC_SUCURSAL, NDC_FECHA,"
                sql = sql & " DND_NROITEM, BAN_CODINT,"
                sql = sql & " CHE_NUMERO, DND_PRECIO, DND_BONIFICA)"
                sql = sql & " VALUES ("
                sql = sql & XN(cboNotaDebito.ItemData(cboNotaDebito.ListIndex)) & ","
                sql = sql & XN(txtNroNotaDebito) & ","
                sql = sql & XN(txtNroSucursal) & ","
                sql = sql & XDQ(FechaNotaDebito) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 11)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 10)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(I, 0)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 6)) & "," 'IMPORTE CHEQUE
                sql = sql & XN(grdGrilla.TextMatrix(I, 7)) & ")"
                DBConn.Execute sql
            End If
        Next
        
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO A LA NOTA DE DEBITO QUE CORRESPONDA
        sql = "SELECT * FROM PARAMETROS"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
                Select Case cboNotaDebito.ItemData(cboNotaDebito.ListIndex)
                    Case 7
                        sql = "UPDATE PARAMETROS SET NOTA_DEBITO_A=" & XN(txtNroNotaDebito)
                    Case 8
                        sql = "UPDATE PARAMETROS SET NOTA_DEBITO_B=" & XN(txtNroNotaDebito)
                End Select
                    DBConn.Execute sql
            
        End If
        Rec1.Close
        
        'ACTUALIZO LA CUENTA CORRIENTE DEL CLIENTE
        DBConn.Execute AgregoCtaCteCliente(txtCodCliente, CStr(cboNotaDebito.ItemData(cboNotaDebito.ListIndex)) _
                                            , txtNroNotaDebito, txtNroSucursal, _
                                            FechaNotaDebito, txtTotal, "D", CStr(Date))
        
        DBConn.CommitTrans
    Else
        MsgBox "La Nota de D�bito ya fue Registrada", vbCritical, TIT_MSGBOX
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
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarNotaBebito() As Boolean
    If IsNull(FechaNotaDebito.Value) Then
        MsgBox "La Fecha de la Nota de D�bito es requerida", vbExclamation, TIT_MSGBOX
        FechaNotaDebito.SetFocus
        ValidarNotaBebito = False
        Exit Function
    End If
    If txtCodCliente.Text = "" Then
        MsgBox "El Cliente es requerido", vbExclamation, TIT_MSGBOX
        txtCodCliente.SetFocus
        ValidarNotaBebito = False
        Exit Function
    End If
    If chkBonificaEnPesos.Value = Checked Or chkBonificaEnPorsentaje.Value = Checked Then
        If txtPorcentajeBoni.Text = "" Then
            MsgBox "Debe ingresar la Bonificaci�n", vbExclamation, TIT_MSGBOX
            txtPorcentajeBoni.SetFocus
            ValidarNotaBebito = False
            Exit Function
        End If
    End If
    ValidarNotaBebito = True
End Function

Private Sub cmdImprimir_Click()
    If MsgBox("�Confirma Impresi�n Nota de D�bito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
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
    ImprimirNotaDebito
End Sub

Public Sub ImprimirNotaDebito()
    Dim Renglon As Double
    Dim canttxt As Integer
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Imprimiendo..."
    
    For w = 1 To 2 'SE IMPRIME POR DUPLICADO
      '-----IMPRESION DEL ENCABEZADO------------------
        ImprimirEncabezado
        
      '---- IMPRESION DE LA NOTA DE DEBITO ------------------
        Renglon = 8.8
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                Imprimir 0.9, Renglon, False, grdGrilla.TextMatrix(I, 0) 'NRO CHEQUE
                If Len(grdGrilla.TextMatrix(I, 1)) <= 36 Then
                    Imprimir 3.8, Renglon, False, grdGrilla.TextMatrix(I, 5) 'DESCRIPCION CHEQUE
                Else
                    'CortarCadena 3.8, Renglon, grdGrilla.TextMatrix(I, 5)
                    justifica_printer 3.8, 14.5, Renglon, grdGrilla.TextMatrix(I, 5)
                    canttxt = Len(grdGrilla.TextMatrix(I, 5))
                    canttxt = canttxt / 36 'es para sacar la cantidad de renglones
                    canttxt = Int(canttxt)
                End If
                Imprimir 15.3, Renglon, False, grdGrilla.TextMatrix(I, 6) 'IMPORTE CHEQUE
                Imprimir 15.7, Renglon, False, IIf(grdGrilla.TextMatrix(I, 7) = "", "0,00", grdGrilla.TextMatrix(I, 4)) 'bonificacion
                Imprimir 17.6, Renglon, False, grdGrilla.TextMatrix(I, 9) 'importe
                Renglon = Renglon + (canttxt * 0.5) + 0.5
            End If
        Next I
        
        
        '-----OBSERVACIONES---------------------
        If txtObservaciones.Text <> "" Then
            Imprimir 1, Renglon + 9.5, True, "Observ.: "
            'CortarCadena 3.2, Renglon + 9.5, Trim(txtObservaciones)
            justifica_printer 3.2, 14.5, Renglon + 9.5, Trim(txtObservaciones)
        End If
        'Imprimir 0, 16.5, True, "texto de bajo del detalle"
        '-------------IMPRIMO TOTALES--------------------
        
        If txtPorcentajeBoni.Text <> "" Then
            If chkBonificaEnPesos.Value = Checked Then
                Imprimir 3.5, 15, True, "$" & txtPorcentajeBoni.Text
                Imprimir 4.4, 15.5, True, txtImporteBoni.Text
            Else
                Imprimir 3.5, 15, True, "%" & txtPorcentajeBoni.Text
                Imprimir 4.4, 15.5, True, txtImporteBoni.Text
            End If
            Imprimir 7.1, 15.5, True, txtSubTotalBoni.Text
        End If
        Imprimir 17.2, 21.9, True, txtSubtotal.Text
        Imprimir 17.2, 23.8, True, txtSubtotal.Text
        Imprimir 14.8, 24.7, True, "    " & txtPorcentajeIva.Text
        Imprimir 17.2, 24.7, True, txtImporteIva.Text
        Imprimir 16.7, 26.4, True, txtTotal.Text
        
        Printer.EndDoc
    Next w
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Public Sub ImprimirEncabezado()
 '-----------IMPRIME EL ENCABEZADO DE LA NOTA DE DEBITO-------------------
    Dim a�o As String
    
    Imprimir 15.8, 1, False, "NOTA DE D�BITO"
    'a�o = String(4, Year(FechaNotaCredito))
    a�o = Year(FechaNotaDebito)
    Imprimir 14.5, 3, False, Format(Day(FechaNotaDebito), "00")
    Imprimir 16.1, 3, False, Format(Month(FechaNotaDebito), "00")
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
        grdGrilla.TextMatrix(I, 9) = ""
        grdGrilla.TextMatrix(I, 10) = ""
        grdGrilla.TextMatrix(I, 11) = I
   Next
   txtCodCliente.Text = ""
   txtNroNotaDebito.Text = ""
   txtNroSucursal.Text = ""
   FechaNotaDebito.Value = Date
   lblEstadoNotaDebito.Caption = ""
   txtSubtotal.Text = ""
   txtTotal.Text = ""
   txtPorcentajeBoni.Text = ""
   txtPorcentajeIva.Text = ""
   txtImporteBoni.Text = ""
   txtSubTotalBoni.Text = ""
   txtImporteIva.Text = ""
   txtObservaciones.Text = ""
   lblEstado.Caption = ""
   cmdGrabar.Enabled = True
   'BUSCO IVA
   BuscoIva
   'CARGO ESTADO
   Call BuscoEstado(1, lblEstadoNotaDebito) 'ESTADO PENDIENTE
   VEstadoNotaDebito = 1
   '--------------
   chkBonificaEnPorsentaje.Value = Unchecked
   chkBonificaEnPesos.Value = Unchecked
   FrameNotaDebito.Enabled = True
   FramePara.Enabled = True
   tabDatos.Tab = 0
   cboNotaDebito.ListIndex = 0
   cboNotaDebito.SetFocus
   GTotAdm = 0
   
End Sub

Private Sub cmdQuitarCheque_Click()
    If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
        If MsgBox("Seguro que desea quitar el Cheque: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 0) & " del " & grdGrilla.TextMatrix(grdGrilla.RowSel, 5), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
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
            grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 11) = ""
            txtSubtotal.Text = ""
            txtImporteIva.Text = ""
            txtTotal.Text = ""
        End If
     End If
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmNotaDeditoClienteCheques = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
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
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    grdGrilla.FormatString = "Nro Cheque|Bco|Loc|Suc|C�digo|Banco|Imp Che|" _
                             & "Bonif.|Pre.Bonif.|Importe|COD INT BANCO|Orden"
    grdGrilla.ColWidth(0) = 1200 'NRO CHEQUE
    grdGrilla.ColWidth(1) = 500  'BCO
    grdGrilla.ColWidth(2) = 500  'LOC
    grdGrilla.ColWidth(3) = 500  'SUC
    grdGrilla.ColWidth(4) = 800  'CODIGO
    grdGrilla.ColWidth(5) = 2900 'BANCO
    grdGrilla.ColWidth(6) = 1000 'IMPORTE CHEQUE
    grdGrilla.ColWidth(7) = 700  'BONOFICACION
    grdGrilla.ColWidth(8) = 1000 'PRE BONIFICACION
    grdGrilla.ColWidth(9) = 1000 'IMPORTE
    grdGrilla.ColWidth(10) = 0   'CODIGO INTERNO BANCO
    grdGrilla.ColWidth(11) = 0   'ORDEN
    grdGrilla.Cols = 12
    grdGrilla.Rows = 1
    For I = 2 To 14
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                             & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                             & Chr(9) & (I - 1)
    Next
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "^Tipo|^N�mero|^Fecha|Importe|Cliente|Cod_Estado|" _
                              & "PORCENTAJE BONIFICA|PORCENTAJE IVA|" _
                              & "OBSERVACIONES|COD TIPO COMPROBANTE NOTA DEBITO|" _
                              & "BONIFICA EN PESOS|COD CLIENTE|REPRESENTADA"
    GrdModulos.ColWidth(0) = 900 'TIPO NOTA DEBITO
    GrdModulos.ColWidth(1) = 1300 'NUMERO
    GrdModulos.ColWidth(2) = 1200 'FECHA
    GrdModulos.ColWidth(3) = 1200 'IMPORTE
    GrdModulos.ColWidth(4) = 5500 'CLIENTE
    GrdModulos.ColWidth(5) = 0    'COD_ESTADO
    GrdModulos.ColWidth(6) = 0    'PORCENTAJE BONIFICA
    GrdModulos.ColWidth(7) = 0    'PORCENTAJE IVA
    GrdModulos.ColWidth(8) = 0    'OBSERVACIONES
    GrdModulos.ColWidth(9) = 0    'COD TIPO COMPROBANTE NOTA DEBITO
    GrdModulos.ColWidth(10) = 0    'BONIFICA EN PESOS
    GrdModulos.ColWidth(11) = 0   'COD CLIENTE
    GrdModulos.ColWidth(12) = 0   'REPRESENTADA
    GrdModulos.Cols = 13
    GrdModulos.Rows = 1
    '------------------------------------
    lblEstado.Caption = ""
    'CARGO COMBO CON LOS TIPOS DE NOTA DE DEBITO
    LlenarComboNotaDebito
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoNotaDebito) 'ESTADO PENDIENTE
    VEstadoNotaDebito = 1
    FechaNotaDebito.Value = Date
    tabDatos.Tab = 0
    'BUSCO IVA
    GTotAdm = 0
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

Private Sub LlenarComboNotaDebito()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'NOTA DE DEB%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboNotaDebito.AddItem rec!TCO_DESCRI
            cboNotaDebito.ItemData(cboNotaDebito.NewIndex) = rec!TCO_CODIGO
            cboNotaDebito1.AddItem rec!TCO_DESCRI
            cboNotaDebito1.ItemData(cboNotaDebito1.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboNotaDebito.ListIndex = 0
        cboNotaDebito1.ListIndex = -1
    End If
    rec.Close
End Sub

Private Function BuscoUltimaNotaDebito(TipoND As Integer) As String
    'ACA BUSCA EL NUMERO DE NOTA DE DEBITO SIGUIENTE AL ULTIMO CARGADO
    sql = "SELECT (NOTA_DEBITO_A) + 1 AS ND_A, (NOTA_DEBITO_B) + 1 AS ND_B"
    sql = sql & " FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Select Case TipoND
            Case 7
                BuscoUltimaNotaDebito = IIf(IsNull(rec!ND_A), 1, rec!ND_A)
            Case 8
                BuscoUltimaNotaDebito = IIf(IsNull(rec!ND_B), 1, rec!ND_B)
            Case 9
                MsgBox "No hay Notas de D�bito del tipo C", vbExclamation, TIT_MSGBOX
                cboNotaDebito.SetFocus
        End Select
    End If
    rec.Close
End Function

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1, 2, 3, 4
            LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
            grdGrilla.TextMatrix(grdGrilla.RowSel, 11) = grdGrilla.RowSel
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            txtTotal.Text = txtSubtotal.Text
            grdGrilla.Col = 0
        Case 7
            VBonificacion = 0
            grdGrilla.Text = ""
            grdGrilla.Col = 8
            grdGrilla.Text = ""
            VBonificacion = CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6))
            grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Valido_Importe(CStr(VBonificacion))
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            txtTotal.Text = txtSubtotal.Text
            grdGrilla.Col = 0
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
        Case 4
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = "" _
               And grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = "" _
               And grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = "" Then
                chkBonificaEnPorsentaje.SetFocus
            End If
        End Select
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
       (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Or _
       (grdGrilla.Col = 4) Or (grdGrilla.Col = 6) Or (grdGrilla.Col = 7) Then
        If KeyAscii = vbKeyReturn Then
            If grdGrilla.Col = 7 Then
                If grdGrilla.row < grdGrilla.Rows - 1 Then
                    grdGrilla.row = grdGrilla.row + 1
                    grdGrilla.Col = 0
                Else
                    SendKeys "{TAB}"
                End If
            ElseIf grdGrilla.Col = 4 Then
                grdGrilla.Col = 6
            Else
                grdGrilla.Col = grdGrilla.Col + 1
            End If
        Else
            If (grdGrilla.Col <> 1) Then
                If KeyAscii > 47 And KeyAscii < 58 Then
                    EDITAR grdGrilla, txtEdit, KeyAscii
                End If
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
        If txtEdit.Visible = False Then Exit Sub
        grdGrilla = txtEdit.Text
        txtEdit.Visible = False
    End If
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        Set Rec1 = New ADODB.Recordset
        lblEstado.Caption = "Buscando..."
        Screen.MousePointer = vbHourglass
        'CABEZA NOTA DEBITO
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 9)), cboNotaDebito)
        txtNroNotaDebito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        FechaNotaDebito.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
        Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 5)), lblEstadoNotaDebito)
        VEstadoNotaDebito = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 5))
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 8) <> "" Then
            txtObservaciones.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 8))
        End If
        'CLIENTE
        txtCodCliente.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 11)
        txtCodCliente_LostFocus
        
        
        '----BUSCO DETALLE DE LA NOTA DE DEBITO------------------
        sql = "SELECT NDC.* "
        sql = sql & " FROM DETALLE_NOTA_DEBITO_CLIENTE NDC" ', BANCO B"
        sql = sql & " WHERE NDC.NDC_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8))
        sql = sql & " AND NDC.NDC_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
        sql = sql & " AND NDC.NDC_FECHA=" & XDQ(GrdModulos.TextMatrix(GrdModulos.RowSel, 2))
        sql = sql & " AND NDC.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 9))
        'sql = sql & " AND NDC.BAN_CODINT=B.BAN_CODINT "
        sql = sql & " ORDER BY NDC.DND_NROITEM"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            I = 1
            Do While Rec1.EOF = False
            
                ' TENGO QUE BUSCAR EL GASTO ADM EN LA TABLA DETALLE_NOTADEBITO_CLIENTE
                ' O VER SI ME CONVIENE GUARDARLO EN OTRO LADO O EN LA TABLA CHEQUE CON UN
                ' BANCO GENERICO
                grdGrilla.TextMatrix(I, 0) = Rec1!CHE_NUMERO
                If Rec1!CHE_NUMERO <> 1 Then
                    sql = "SELECT BAN_BANCO, BAN_LOCALIDAD, BAN_SUCURSAL, BAN_CODIGO, BAN_DESCRI"
                    sql = sql & " FROM BANCO WHERE BAN_CODINT = " & Rec1!BAN_CODINT
                    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                    If rec.EOF = False Then
                        grdGrilla.TextMatrix(I, 1) = rec!BAN_BANCO
                        grdGrilla.TextMatrix(I, 2) = rec!BAN_LOCALIDAD
                        grdGrilla.TextMatrix(I, 3) = rec!BAN_SUCURSAL
                        grdGrilla.TextMatrix(I, 4) = rec!BAN_CODIGO
                        grdGrilla.TextMatrix(I, 5) = rec!BAN_DESCRI
                    Else
                        MsgBox "Banco inexistente", vbExclamation, TIT_MSGBOX
                        
                    End If
                    rec.Close
                Else
                    grdGrilla.TextMatrix(I, 5) = "GASTOS ADMINISTRATIVOS"
                End If
                If Rec1!CHE_NUMERO = 1 Then
                    GTotAdm = Valido_Importe(Rec1!DND_PRECIO)
                End If
                grdGrilla.TextMatrix(I, 6) = Valido_Importe(Rec1!DND_PRECIO)  'IMPORTE CHEQUE
                If IsNull(Rec1!DND_BONIFICA) Then
                    grdGrilla.TextMatrix(I, 7) = ""
                Else
                    grdGrilla.TextMatrix(I, 7) = Valido_Importe(Rec1!DND_BONIFICA)
                End If
                VBonificacion = 0
                If Not IsNull(Rec1!DND_BONIFICA) Then
                    VBonificacion = ((CDbl(Rec1!DND_PRECIO) * CDbl(Rec1!DND_BONIFICA)) / 100)
                    VBonificacion = (CDbl(Rec1!DND_PRECIO) - VBonificacion)
                    grdGrilla.TextMatrix(I, 8) = Valido_Importe(CStr(VBonificacion))
                    grdGrilla.TextMatrix(I, 9) = Valido_Importe(CStr(VBonificacion))
                Else
                    VBonificacion = (CDbl(Rec1!DND_PRECIO))
                    grdGrilla.TextMatrix(I, 8) = ""
                    grdGrilla.TextMatrix(I, 9) = Valido_Importe(CStr(VBonificacion))
                End If
                grdGrilla.TextMatrix(I, 10) = IIf(IsNull(Rec1!BAN_CODINT), "", Rec1!BAN_CODINT)
                grdGrilla.TextMatrix(I, 11) = Rec1!DND_NROITEM
                I = I + 1
                Rec1.MoveNext
            Loop
            VBonificacion = 0
        End If
        Rec1.Close
        '--CARGO LOS TOTALES----
        txtSubtotal.Text = Valido_Importe(SumaBonificacion)
        txtTotal.Text = txtSubtotal.Text
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 10) = "S" Then 'SI BONOFICA EN PESOS
            chkBonificaEnPesos.Value = Checked
        ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 10) = "N" Then 'SI BONIFICA EN PORCENTAJE
            chkBonificaEnPorsentaje.Value = Checked
        Else
            chkBonificaEnPesos.Value = Unchecked
            chkBonificaEnPorsentaje.Value = Unchecked
        End If
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) <> "" Then 'PORCENTAJE DE BONIFICACION
            txtPorcentajeBoni.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
            txtPorcentajeBoni_LostFocus
        End If
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 7) <> "" Then 'PORCENTAJE IVA
            txtPorcentajeIva = GrdModulos.TextMatrix(GrdModulos.RowSel, 7)
            txtPorcentajeIva_LostFocus
        End If
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        '--------------
        FrameNotaDebito.Enabled = False
        FramePara.Enabled = False
        '--------------
        tabDatos.Tab = 0
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    GrdModulos.Rows = 2
    txtCliente.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    txtVendedor.Enabled = False
    cboNotaDebito1.Enabled = False
    cmdBuscarCli.Enabled = False
    cmdBuscarVen.Enabled = False
    cmdGrabar.Enabled = False
    LimpiarBusqueda
    If Me.Visible = True Then chkCliente.SetFocus
  Else
    If VEstadoNotaDebito = 1 Then
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
    txtVendedor.Text = ""
    txtDesVen.Text = ""
    GrdModulos.Rows = 1
    cboNotaDebito1.ListIndex = -1
    chkCliente.Value = Unchecked
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
    If chkFecha.Value = Unchecked _
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
            txtCUIT.Text = Rec1!CLI_CUIT
            txtCondicionIVA.Text = Rec1!IVA_DESCRI
            txtIngBrutos.Text = IIf(IsNull(Rec1!CLI_INGBRU), "", Rec1!CLI_INGBRU)
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
        Rec1.Close
    End If
End Sub

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

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    If grdGrilla.Col = 0 Then
        KeyAscii = CarNumeroEntero(KeyAscii)
        txtEdit.MaxLength = 10
    End If
    If grdGrilla.Col = 1 Then
        KeyAscii = CarNumeroEntero(KeyAscii)
        txtEdit.MaxLength = 3
    End If
    If grdGrilla.Col = 2 Then
        KeyAscii = CarNumeroEntero(KeyAscii)
        txtEdit.MaxLength = 3
    End If
    If grdGrilla.Col = 3 Then
        KeyAscii = CarNumeroEntero(KeyAscii)
        txtEdit.MaxLength = 3
    End If
    If grdGrilla.Col = 4 Then
        KeyAscii = CarNumeroEntero(KeyAscii)
        txtEdit.MaxLength = 6
    End If
    If grdGrilla.Col = 6 Or grdGrilla.Col = 7 Then
        KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
        txtEdit.MaxLength = 15
    End If
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
            
            Case 0 'NUMERO CHEQUE
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    If Trim(txtEdit.Text) = 1 Then
                        grdGrilla.TextMatrix(grdGrilla.row, 5) = "GASTOS ADMINISTRATIVOS"
                        grdGrilla.Col = 6
                        grdGrilla.SetFocus
                        Exit Sub
                    Else
                        txtEdit.Text = Format(txtEdit.Text, "00000000")
                        grdGrilla.Col = 1
                    End If
                End If
                'BUSCO EL CHEQUE y BANCO-------------------------------------
                If grdGrilla.TextMatrix(grdGrilla.row, 0) <> "" And _
                    txtEdit.Text <> "" Then
                    
                    'BUSCO EL CODIGO INTERNO
                    sql = "SELECT B.BAN_CODINT, B.BAN_DESCRI,C.CHE_IMPORT"
                    sql = sql & ",B.BAN_BANCO, B.BAN_LOCALIDAD,B.BAN_SUCURSAL, B.BAN_CODIGO"
                    sql = sql & " FROM CHEQUE C, BANCO B"
                    sql = sql & " WHERE C.BAN_CODINT= B.BAN_CODINT AND "
                    sql = sql & "CHE_NUMERO = " & XS(grdGrilla.TextMatrix(grdGrilla.row, 0))
                    
                    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                    If rec.EOF = False Then 'EXITE
                       grdGrilla.TextMatrix(grdGrilla.row, 10) = rec!BAN_CODINT
                       grdGrilla.TextMatrix(grdGrilla.row, 5) = rec!BAN_DESCRI
                       grdGrilla.TextMatrix(grdGrilla.row, 6) = Valido_Importe(rec!che_import)
                       grdGrilla.TextMatrix(grdGrilla.row, 9) = Valido_Importe(rec!che_import)
                       grdGrilla.TextMatrix(grdGrilla.row, 1) = rec!BAN_BANCO
                       grdGrilla.TextMatrix(grdGrilla.row, 2) = rec!BAN_LOCALIDAD
                       grdGrilla.TextMatrix(grdGrilla.row, 3) = rec!BAN_SUCURSAL
                       grdGrilla.TextMatrix(grdGrilla.row, 4) = rec!BAN_CODIGO
                       grdGrilla.Col = 6
                       grdGrilla.SetFocus
                       
                       txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                       txtTotal.Text = txtSubtotal.Text
                       txtPorcentajeIva_LostFocus
                       
                       rec.Close
                    Else
                       If Me.ActiveControl.Name <> "CmdSalir" And Me.ActiveControl.Name <> "CmdNuevo" Then
                         MsgBox "Cheque NO Registrado.", 16, TIT_MSGBOX
                         grdGrilla.Col = 1
                         grdGrilla.SetFocus
                       End If
                       rec.Close
                       Exit Sub
                    End If
                
                
                
                End If
            Case 1 'BCO
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    txtEdit.Text = Format(txtEdit.Text, "000")
                    grdGrilla.Col = 2
                End If
                
            Case 2 'LOC
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    txtEdit.Text = Format(txtEdit.Text, "000")
                    grdGrilla.Col = 3
                End If
            
            Case 3 'SUC
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    txtEdit.Text = Format(txtEdit.Text, "000")
                    grdGrilla.Col = 4
                End If
            
            Case 4 'CODIGO
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    txtEdit.Text = Format(txtEdit.Text, "0000")
                    grdGrilla.Col = 6
                End If
                'BUSCO EL BANCO-------------------------------------
                If grdGrilla.TextMatrix(grdGrilla.row, 0) <> "" And _
                    grdGrilla.TextMatrix(grdGrilla.row, 1) <> "" And _
                    grdGrilla.TextMatrix(grdGrilla.row, 2) <> "" And _
                    grdGrilla.TextMatrix(grdGrilla.row, 3) <> "" And _
                    txtEdit.Text <> "" Then
                    
                    'BUSCO EL CODIGO INTERNO
                    sql = "SELECT BAN_CODINT, BAN_DESCRI"
                    sql = sql & " FROM BANCO"
                    sql = sql & " WHERE BAN_BANCO = " & XS(grdGrilla.TextMatrix(grdGrilla.row, 1))
                    sql = sql & " AND BAN_LOCALIDAD = " & XS(grdGrilla.TextMatrix(grdGrilla.row, 2))
                    sql = sql & " AND BAN_SUCURSAL = " & XS(grdGrilla.TextMatrix(grdGrilla.row, 3))
                    sql = sql & " AND BAN_CODIGO = " & XS(txtEdit.Text)
                    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                    If rec.EOF = False Then 'EXITE
                       grdGrilla.TextMatrix(grdGrilla.row, 10) = rec!BAN_CODINT
                       grdGrilla.TextMatrix(grdGrilla.row, 5) = rec!BAN_DESCRI
                       rec.Close
                    Else
                       If Me.ActiveControl.Name <> "CmdSalir" And Me.ActiveControl.Name <> "CmdNuevo" Then
                         MsgBox "Banco NO Registrado.", 16, TIT_MSGBOX
                         grdGrilla.Col = 1
                         grdGrilla.SetFocus
                       End If
                       rec.Close
                       Exit Sub
                    End If
                Else
                    MsgBox "Faltan Datos", vbCritical, TIT_MSGBOX
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                End If
                
            Case 6 'IMPORTE CHEQUE
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    txtEdit.Text = Valido_Importe(txtEdit.Text)
                    grdGrilla.Col = 7
                End If
                
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> 1 Then
                    If grdGrilla.TextMatrix(grdGrilla.row, 0) <> "" And _
                        grdGrilla.TextMatrix(grdGrilla.row, 1) <> "" And _
                        grdGrilla.TextMatrix(grdGrilla.row, 2) <> "" And _
                        grdGrilla.TextMatrix(grdGrilla.row, 3) <> "" And _
                        txtEdit.Text <> "" Then
                    
                        VBonificacion = CDbl(txtEdit.Text)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Valido_Importe(CStr(VBonificacion))
                        If grdGrilla.TextMatrix(grdGrilla.RowSel, 7) <> "" Then
                            VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 9)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 7))) / 100)
                            VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 9)) - VBonificacion)
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = Valido_Importe(CStr(VBonificacion))
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Valido_Importe(CStr(VBonificacion))
                        End If
                        txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                        txtTotal.Text = txtSubtotal.Text
                        txtPorcentajeIva_LostFocus
                    Else
                        MsgBox "No puede ingresar el importe del cheque, Faltan datos!!", vbExclamation, TIT_MSGBOX
                        grdGrilla.TextMatrix(grdGrilla.row, 6) = ""
                        grdGrilla.Col = 0
                        grdGrilla.SetFocus
                        Exit Sub
                    End If
                 Else
                    VBonificacion = CDbl(txtEdit.Text)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Valido_Importe(CStr(VBonificacion))
                    If grdGrilla.TextMatrix(grdGrilla.RowSel, 7) <> "" Then
                        VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 9)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 7))) / 100)
                        VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 9)) - VBonificacion)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = Valido_Importe(CStr(VBonificacion))
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Valido_Importe(CStr(VBonificacion))
                    End If
                    txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                    txtTotal.Text = txtSubtotal.Text
                    GTotAdm = grdGrilla.TextMatrix(grdGrilla.RowSel, 9)
                    txtPorcentajeIva_LostFocus
                 End If
                
            Case 7 'BONIFICACION
                If Trim(txtEdit) <> "" Then
                    If txtEdit.Text = ValidarPorcentaje(txtEdit) = False Then
                        Exit Sub
                    End If
                    VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 9)) * CDbl(txtEdit.Text)) / 100)
                    VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 9)) - VBonificacion)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = Valido_Importe(CStr(VBonificacion))
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Valido_Importe(CStr(VBonificacion))
                    txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                    txtTotal.Text = txtSubtotal.Text
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
                MsgBox "El Servicio ya fue elegido anteriormente", vbExclamation, TIT_MSGBOX
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
Private Function SumaTotal() As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 9) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(I, 9))
        End If
    Next
    SumaTotal = Valido_Importe(CStr(VTotal))
End Function

Private Function SumaBonificacion() As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 9) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(I, 9))
        End If
    Next
    SumaBonificacion = Valido_Importe(CStr(VTotal))
End Function

Private Sub txtNroNotaDebito_GotFocus()
    SelecTexto txtNroNotaDebito
End Sub

Private Sub txtNroNotaDebito_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaDebito_LostFocus()
    If txtNroNotaDebito.Text = "" Then
        'BUSCO EL NUMERO DE NOTA DE DEBITO QUE CORRESPONDE
        txtNroNotaDebito.Text = Format(BuscoUltimaNotaDebito(cboNotaDebito.ItemData(cboNotaDebito.ListIndex)), "00000000")
    Else
        txtNroNotaDebito.Text = Format(txtNroNotaDebito.Text, "00000000")
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
            txtImporteIva.Text = (GTotAdm * CDbl(txtPorcentajeIva.Text)) / 100 ' solo iva de gastos administrativos
            txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
            txtTotal.Text = CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
        End If
    End If
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
        Imprimir COLUMNA, Renglon + 0.5, False, Trim(cadena2)
        Imprimir COLUMNA, Renglon + 1, False, Trim(cadena3)
        Imprimir COLUMNA, Renglon + 1.5, False, Trim(cadena4)
        Imprimir COLUMNA, Renglon + 2, False, Trim(cadena5)
        Imprimir COLUMNA, Renglon + 2.5, False, Trim(cadena6)
        Imprimir COLUMNA, Renglon + 3, False, Trim(cadena7)
        Imprimir COLUMNA, Renglon + 3.5, False, Trim(cadena8)
    'Else
    '    cadena1 = cadena
    '    MsgBox cadena1
    'End If
    
End Sub
