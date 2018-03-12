VERSION 5.00
Begin VB.Form frmconfiguraimpresion 
   Caption         =   "Configuración de Impresión de Comprobantes"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form2"
   ScaleHeight     =   6540
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Guardar"
      DisabledPicture =   "frmconfiguraimpresion.frx":0000
      Height          =   720
      Index           =   0
      Left            =   5100
      Picture         =   "frmconfiguraimpresion.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmconfiguraimpresion.frx":0614
      Height          =   720
      Index           =   2
      Left            =   5970
      Picture         =   "frmconfiguraimpresion.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Salir"
      DisabledPicture =   "frmconfiguraimpresion.frx":0C28
      Height          =   720
      Left            =   7710
      Picture         =   "frmconfiguraimpresion.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdBotonDatos 
      Caption         =   "&Eliminar"
      DisabledPicture =   "frmconfiguraimpresion.frx":123C
      Height          =   720
      Index           =   1
      Left            =   6840
      Picture         =   "frmconfiguraimpresion.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   2160
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   1320
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   6960
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   6120
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   2280
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1440
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   6960
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   6120
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   7080
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   6240
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2280
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Y"
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
         Left            =   2400
         TabIndex        =   45
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "X"
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
         Left            =   1680
         TabIndex        =   44
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Detalle"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   41
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Y"
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
         Left            =   7200
         TabIndex        =   40
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "X"
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
         Left            =   6480
         TabIndex        =   39
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Remito"
         Height          =   195
         Index           =   2
         Left            =   5160
         TabIndex        =   36
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Y"
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
         Left            =   2520
         TabIndex        =   35
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "X"
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
         Left            =   1800
         TabIndex        =   34
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Condicion Venta"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   3120
         Width           =   1170
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Y"
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
         Left            =   7200
         TabIndex        =   30
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "X"
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
         Left            =   6480
         TabIndex        =   29
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "CUIT"
         Height          =   195
         Left            =   5160
         TabIndex        =   26
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Y"
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
         Left            =   2400
         TabIndex        =   25
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "X"
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
         Left            =   1680
         TabIndex        =   24
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "IVA"
         Height          =   195
         Left            =   720
         TabIndex        =   21
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Y"
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
         Left            =   7320
         TabIndex        =   20
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "X"
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
         Left            =   6600
         TabIndex        =   19
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   5640
         TabIndex        =   16
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Y"
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
         Left            =   2520
         TabIndex        =   13
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "X"
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
         Left            =   1800
         TabIndex        =   12
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Comprobante:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.Label lblEstado 
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
      Left            =   240
      TabIndex        =   5
      Top             =   6000
      Width           =   750
   End
End
Attribute VB_Name = "frmconfiguraimpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
