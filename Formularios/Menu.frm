VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MENU 
   BackColor       =   &H80000010&
   ClientHeight    =   7830
   ClientLeft      =   435
   ClientTop       =   2580
   ClientWidth     =   11010
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   12600
      Left            =   0
      Picture         =   "Menu.frx":08CA
      ScaleHeight     =   12540
      ScaleWidth      =   10950
      TabIndex        =   1
      Top             =   0
      Width           =   11010
      Begin VB.Frame Frame1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Caption         =   "Stock"
         Height          =   1695
         Index           =   3
         Left            =   6720
         TabIndex        =   14
         Top             =   6840
         Width           =   3615
         Begin VB.CommandButton cmdListados 
            Height          =   1095
            Left            =   240
            Picture         =   "Menu.frx":E07D
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Administrador de Listados"
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            Caption         =   "ADMINISTRADOR DE LISTADOS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   2865
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Caption         =   "Stock"
         Height          =   1695
         Index           =   2
         Left            =   1680
         TabIndex        =   10
         Top             =   6600
         Width           =   2415
         Begin VB.CommandButton cmdPrecio 
            Caption         =   "&Precios"
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
            Index           =   2
            Left            =   1200
            Picture         =   "Menu.frx":FD47
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Lista de Precios"
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdPto 
            Caption         =   "&Productos"
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
            Picture         =   "Menu.frx":11A11
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "ABM de Productos"
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   360
            TabIndex        =   13
            Top             =   240
            Width           =   630
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Caption         =   "Ventas"
         Height          =   1695
         Index           =   1
         Left            =   6840
         TabIndex        =   6
         Top             =   1800
         Width           =   3615
         Begin VB.CommandButton cmdGastosGrales 
            Caption         =   "&Gastos"
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
            Index           =   0
            Left            =   1320
            Picture         =   "Menu.frx":136DB
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Ingreso Facturas de Proveedores"
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemCom 
            Caption         =   "&Remitos"
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
            Left            =   240
            Picture         =   "Menu.frx":13A65
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Ingreso Remitos de Proveedores"
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdFacCom 
            Caption         =   "&Facturas"
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
            Index           =   1
            Left            =   2400
            Picture         =   "Menu.frx":1572F
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Ingreso Facturas de Proveedores"
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            Caption         =   "COMPRAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Caption         =   "Ventas"
         Height          =   1695
         Index           =   0
         Left            =   1680
         TabIndex        =   2
         Top             =   1800
         Width           =   2415
         Begin VB.CommandButton cmdFacturas 
            Caption         =   "&Facturas"
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
            Index           =   0
            Left            =   1200
            Picture         =   "Menu.frx":173F9
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Ingreso Facturas de Clientes"
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemitos 
            Caption         =   "&Remitos"
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
            Picture         =   "Menu.frx":190C3
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Ingreso Remitos de Clientes"
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            Caption         =   "VENTAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   120
            Width           =   750
         End
      End
   End
   Begin ComctlLib.StatusBar B1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   7485
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "NÚM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1409
            MinWidth        =   1409
            TextSave        =   "MAYÚS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Alignment       =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "INS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   9701
            MinWidth        =   9701
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "10/03/2018"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1500
            MinWidth        =   1500
            TextSave        =   "6:41"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MousePointer    =   2
   End
   Begin VB.Menu mnuSistema 
      Caption         =   "&Sistema"
      Begin VB.Menu mnuConectar 
         Caption         =   "&Conectar"
      End
      Begin VB.Menu mnuDesconectar 
         Caption         =   "&Desconectar"
      End
      Begin VB.Menu mnuRaya23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuusuario 
         Caption         =   "&Usuarios   "
         HelpContextID   =   1
      End
      Begin VB.Menu mnupermisos 
         Caption         =   "&Permisos   "
         HelpContextID   =   1
      End
      Begin VB.Menu MNURAYA1 
         Caption         =   "-"
         HelpContextID   =   1
      End
      Begin VB.Menu mnuBackupBD 
         Caption         =   "&Backup BD"
      End
      Begin VB.Menu mnuRestaurarBD 
         Caption         =   "&Restaurar BD"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnurayabkp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParametros 
         Caption         =   "&Parametros"
      End
      Begin VB.Menu mnuCalendario 
         Caption         =   "C&alendario"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCalculadora 
         Caption         =   "Calc&uladora"
      End
      Begin VB.Menu MNURAYA8 
         Caption         =   "-"
      End
      Begin VB.Menu mnusalir 
         Caption         =   "Salir   "
         HelpContextID   =   1
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuComprasFacturacion 
      Caption         =   "   &Ventas"
      Begin VB.Menu mnuFacturaActualiza 
         Caption         =   "&Actualizaciones"
         Begin VB.Menu mnuABMClientes 
            Caption         =   "...ABM de &Clientes"
         End
         Begin VB.Menu mnuABMVendedores 
            Caption         =   "...ABM de &Vendedores"
         End
         Begin VB.Menu mnuABMPais 
            Caption         =   "...ABM de &País"
         End
         Begin VB.Menu mnuABMProvincias 
            Caption         =   "...ABM de &Provincias"
         End
         Begin VB.Menu mnuABMLocalidades 
            Caption         =   "...ABM de &Localidades"
         End
         Begin VB.Menu mnuFacturacionTipoComprobante 
            Caption         =   "...ABM Tipo de Comprobante"
         End
         Begin VB.Menu mnuABMInscIVA 
            Caption         =   "...ABM de Insc. &IVA"
         End
         Begin VB.Menu mnuABMCanales 
            Caption         =   "...ABM de Tipos de Clientes"
         End
         Begin VB.Menu mnuABMEstadoDocumento 
            Caption         =   "...ABM Estado de Documentos"
         End
         Begin VB.Menu mnuABMFormaPago 
            Caption         =   "...ABM de Forma de Pago"
         End
         Begin VB.Menu mnuABMConceptoNotaCredito 
            Caption         =   "...ABM Concepto Nota de Crédito"
         End
         Begin VB.Menu mnuABMServicios 
            Caption         =   "...ABM de Servicios"
         End
      End
      Begin VB.Menu mnuFacturacionPedidos 
         Caption         =   "&Presupuestos"
         Begin VB.Menu mnuIngNotaPedidoCliente 
            Caption         =   "Ingreso de Presupuestos"
         End
         Begin VB.Menu mnuListadoNotaPedidoCliente 
            Caption         =   "Listado de Presupuestos"
         End
      End
      Begin VB.Menu mnuFacturacionRemitos 
         Caption         =   "&Remitos"
         Begin VB.Menu mnuIngRemitosClientes 
            Caption         =   "Ingreso de Remitos de Clientes"
         End
         Begin VB.Menu mnuListadoRemitosClientes 
            Caption         =   "Listado de Remitos de Clientes"
         End
      End
      Begin VB.Menu mnuFacturacionFacturacion 
         Caption         =   "&Facturación"
         Begin VB.Menu mnuFacturacionFactura 
            Caption         =   "Facturación"
         End
         Begin VB.Menu mnuNotaCredito 
            Caption         =   "Notas de Crédito"
         End
         Begin VB.Menu mnuComprasNotaDebito 
            Caption         =   "Notas de Débito"
            Begin VB.Menu mnuNotaDebitoPorServicios 
               Caption         =   "...por Servicio"
            End
            Begin VB.Menu mnuNotaDebitoPorChequeDevuelto 
               Caption         =   "..por Cheque Devuelto"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu MNURAYA13 
            Caption         =   "-"
         End
         Begin VB.Menu mnuimputaNCFAC 
            Caption         =   "Imputar NC a Facturas"
         End
      End
      Begin VB.Menu mnuFacturacionRecibos 
         Caption         =   "R&ecibos"
         Begin VB.Menu mnuIngRecibosClientes 
            Caption         =   "Ingreso de Recibos de Clientes"
         End
         Begin VB.Menu mnuListadoRecibosClientes 
            Caption         =   "Listado de Recibos de Clientes"
         End
      End
      Begin VB.Menu mnuFacturacionCtaCte 
         Caption         =   "&Cuenta Corriente de Clientes"
      End
      Begin VB.Menu mnuConsultaAnulaciones 
         Caption         =   "C&onsulta - Anulaciones"
         Begin VB.Menu mnuConAnuPedidos 
            Caption         =   "... de Presupuestos"
         End
         Begin VB.Menu mnuConAnuRemitos 
            Caption         =   "... de Remitos"
         End
         Begin VB.Menu mnuConAnuFactura 
            Caption         =   "... de Factura"
         End
         Begin VB.Menu mnuConAnuRecibos 
            Caption         =   "... de Recibos"
         End
      End
      Begin VB.Menu mnuRaya19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListadoS 
         Caption         =   "&Listados"
         Begin VB.Menu mnuListadoClientes 
            Caption         =   "Clientes"
         End
         Begin VB.Menu mnuRaya22 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFacturasPendientesCliente 
            Caption         =   "Facturas pendientes por Cliente"
         End
         Begin VB.Menu mnuPagosRealizadosPorCliente 
            Caption         =   "Pagos realizados por Clientes"
         End
         Begin VB.Menu mnuRayaLis 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEstaCantidadVendida 
            Caption         =   "Listado de Cantidades Vendidas"
         End
         Begin VB.Menu mnuEstaVentaporCliente 
            Caption         =   "Ventas por Cliente"
         End
         Begin VB.Menu mnuListadoVentaPorVendedor 
            Caption         =   "Ventas por Vendedor"
         End
      End
      Begin VB.Menu mnurayal 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLibroIvaVentas 
         Caption         =   "Libro &Débito Fiscal"
      End
   End
   Begin VB.Menu mnuCompras 
      Caption         =   "   &Compras"
      Index           =   1
      Begin VB.Menu mnuComprasActualiza 
         Caption         =   "&Actualizaciones"
         Begin VB.Menu mnuABMTipoProveedores 
            Caption         =   "... ABM de Tipo de Proveedores"
         End
         Begin VB.Menu mnuABMProveedores 
            Caption         =   "... ABM de Proveedores"
         End
         Begin VB.Menu mnuABMTipoGastos 
            Caption         =   "... ABM de Tipo de Gastos"
         End
      End
      Begin VB.Menu mnuComprasOc 
         Caption         =   "&Compras"
         Begin VB.Menu mnuOrdendeCompra 
            Caption         =   "Ingreso de Ordenes de Compras"
         End
         Begin VB.Menu mnuListadoOrdCompra 
            Caption         =   "Listado de Ordenes de Compra"
         End
      End
      Begin VB.Menu mnuremito 
         Caption         =   "&Remitos"
         Begin VB.Menu mnuRemitoCompras 
            Caption         =   "Ingreso de Remitos de Proveedores"
         End
         Begin VB.Menu mnuListadoRemProv 
            Caption         =   "Listado de Remitos de Proveedores"
         End
      End
      Begin VB.Menu mnuProveedores 
         Caption         =   "&Facturación"
         Begin VB.Menu mnuProveedoresFacturas 
            Caption         =   "&Facturación"
         End
         Begin VB.Menu mnuNotaCreditoProveedor 
            Caption         =   "Nota de &Crédito"
         End
         Begin VB.Menu mnuNotaDebitoProveedor 
            Caption         =   "Nota de &Débito"
            Begin VB.Menu mnuNDProvServicio 
               Caption         =   "..por Servicio"
            End
            Begin VB.Menu mnuNDProvCheque 
               Caption         =   "..por Cheque Devuelto"
            End
         End
         Begin VB.Menu MNURAYA12 
            Caption         =   "-"
         End
         Begin VB.Menu mnuImputarNCFacturasProveedores 
            Caption         =   "Imputar NC a Facturas"
         End
      End
      Begin VB.Menu mnuGastos 
         Caption         =   "&Gastos Generales"
         Begin VB.Menu mnuGastosGeneralesRegistro 
            Caption         =   "R&egistro de Gastos Generales"
         End
         Begin VB.Menu mnuListadoGastosGrales 
            Caption         =   "Listado de Gastos Generales"
         End
      End
      Begin VB.Menu mnuPagoProveedores 
         Caption         =   "Orden de Pago"
         Begin VB.Menu mnuPagoProveedoresOrdenPago 
            Caption         =   "Orden de Pago"
         End
      End
      Begin VB.Menu mnuComprasCtaCte 
         Caption         =   "&Cuenta  Corriente de Proveedores"
      End
      Begin VB.Menu mnconsanu 
         Caption         =   "C&onsulta - Anulaciones"
         Begin VB.Menu mnuOrdenCompra 
            Caption         =   "... de Ordenes de Compra"
         End
         Begin VB.Menu mnuremitos 
            Caption         =   "... de Remitos"
         End
         Begin VB.Menu mnufacturas 
            Caption         =   "...de Facturas"
         End
         Begin VB.Menu mnuOrdenesdePago 
            Caption         =   "...de Ordenes de Pago"
         End
         Begin VB.Menu mnuAnularGastos 
            Caption         =   "...de Gastos Generales"
         End
         Begin VB.Menu mnuAnulaNC 
            Caption         =   "...de Notas de Credito"
         End
      End
      Begin VB.Menu mnuraya10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComprasListado 
         Caption         =   "&Listado"
         Begin VB.Menu mnuListadoProveedores 
            Caption         =   "Listado de P&roveedores"
         End
         Begin VB.Menu mnuraya2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFacturasPendientesProveedor 
            Caption         =   "Facturas pendientes por Proveedor"
         End
         Begin VB.Menu mnuPagosRealizadosProveedores 
            Caption         =   "Orden de Pago"
         End
         Begin VB.Menu mnuraya20 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComprasListaResumen 
            Caption         =   "Resumen de Compras"
         End
      End
      Begin VB.Menu mnuraya11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComprasCreditoFiscal 
         Caption         =   "Libro &Crédito Fiscal"
      End
   End
   Begin VB.Menu mnuGestionStock 
      Caption         =   "   &Stock"
      Begin VB.Menu mnuStockActualiza 
         Caption         =   "&Actualizaciones"
         Begin VB.Menu mnuStockABMProductos 
            Caption         =   "...ABM de Productos"
         End
         Begin VB.Menu mnuRaya16 
            Caption         =   "-"
         End
         Begin VB.Menu mnuABMLineas 
            Caption         =   "...ABM de Lineas"
         End
         Begin VB.Menu mnuABMRubros 
            Caption         =   "...ABM de Rubros"
         End
         Begin VB.Menu mnuABMPresentacion 
            Caption         =   "...ABM de Marcas"
         End
         Begin VB.Menu mnuraya 
            Caption         =   "-"
         End
         Begin VB.Menu mnuABMTransporte 
            Caption         =   "...ABM de Transporte"
         End
      End
      Begin VB.Menu mnuEntradaProductos 
         Caption         =   "&Recepción de Mercadería"
      End
      Begin VB.Menu mnuEntregaDeProductos 
         Caption         =   "&Salida de Mercadería"
      End
      Begin VB.Menu mnuRaya21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockAjuste 
         Caption         =   "Control de &Stock"
      End
      Begin VB.Menu mnuConStock 
         Caption         =   "Consulta de Stock"
      End
      Begin VB.Menu MNURAYA5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListaPrecios 
         Caption         =   "Lista de &Precios"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuConsListaPrecios 
         Caption         =   "C&onsulta de Lista de Precios"
      End
   End
   Begin VB.Menu mnuFondos 
      Caption         =   "   F&ondos"
      Begin VB.Menu mnuFondosActualizaciones 
         Caption         =   "&Actualizaciones"
         Begin VB.Menu mnuABMTipoCuentas 
            Caption         =   "...ABM Tipo de Cuentas"
         End
         Begin VB.Menu mnuFondosCuentas 
            Caption         =   "...ABM de Cuentas"
         End
         Begin VB.Menu mnuFondosBancos 
            Caption         =   "...ABM de Bancos"
         End
         Begin VB.Menu mnuFondosEstadoCheques 
            Caption         =   "...ABM de Estados de Cheques"
         End
         Begin VB.Menu mnuABMMoneda 
            Caption         =   "...ABM de Moneda"
         End
         Begin VB.Menu mnuABMTiposGastosBancarios 
            Caption         =   "...ABM Tipos de Gastos Bancarios"
         End
      End
      Begin VB.Menu mnuFondosGestonBancaria 
         Caption         =   "Gestión &Bancaria"
         Begin VB.Menu mnuFondosMovBancarios 
            Caption         =   "Movimientos Bancararios"
         End
         Begin VB.Menu mnuFondosResumenCuneta 
            Caption         =   "Resumen de Cuenta"
         End
      End
      Begin VB.Menu mnuFondosGestionCaja 
         Caption         =   "Gestión &Caja"
         Begin VB.Menu mnuFondosCargaIngresos 
            Caption         =   "Carga de Ingresos"
         End
         Begin VB.Menu mnuFondosCargaEgresos 
            Caption         =   "Carga de Egresos"
         End
         Begin VB.Menu mnuRaya15 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLiquidacionCobranza 
            Caption         =   "Liquidación de Cobranza"
         End
         Begin VB.Menu mnuFonfosCierreCaja 
            Caption         =   "Cierre de Caja"
         End
      End
      Begin VB.Menu mnuFondosValores 
         Caption         =   "&Valores"
         Begin VB.Menu mnuFondosCargaCheques 
            Caption         =   "Carga de Cheques de Terceros"
         End
         Begin VB.Menu mnuBoletaDeposito 
            Caption         =   "Boleta Déposito"
         End
         Begin VB.Menu mnuFondosCambioEstadoChques 
            Caption         =   "Cambio de Estado Cheques de Terceros"
         End
         Begin VB.Menu mnuFondosListadoCheques 
            Caption         =   "Listado de Cheques de Terceros"
         End
         Begin VB.Menu mnuraya25 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFondosCargaChequesPropios 
            Caption         =   "Carga de Cheques Propios"
         End
         Begin VB.Menu mnuCambioEstadoChequesPropios 
            Caption         =   "Cambio de Estado Cheques Propios"
         End
         Begin VB.Menu mnuListadoChequesPropios 
            Caption         =   "Listado de Cheques Propios"
         End
         Begin VB.Menu mnuraya26 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIngresoGastosBancarios 
            Caption         =   "Ingreso de Gastos Bancarios"
         End
      End
   End
   Begin VB.Menu mnuayuda 
      Caption         =   "  A&yuda"
   End
   Begin VB.Menu mnuacercade 
      Caption         =   "  &Acerca de... "
   End
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TituloPrincipal As String

Private Declare Function ShellAbout Lib "shell32.dll" Alias _
"ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, _
ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Private Sub cmdFacCom_Click(Index As Integer)
    frmfacturaproveedor.Show vbModal
    
End Sub



Private Sub cmdFacturas_Click(Index As Integer)
    frmFacturaCliente.Show vbModal
End Sub

Private Sub cmdGastosGrales_Click(Index As Integer)
    frmCargaGastosGenerales.Show vbModal
End Sub

Private Sub cmdListados_Click()
    frmAdmListados.Show vbModal
End Sub

Private Sub cmdPrecio_Click(Index As Integer)
    Consulta = 2
    FrmListadePrecios.Show vbModal
End Sub

Private Sub cmdPto_Click()
    ABMProducto.Show vbModal
End Sub

Private Sub cmdRemCom_Click()
    frmRemitoProveedor.Show vbModal
End Sub

Private Sub cmdRemitos_Click()
    frmRemitoCliente.Show vbModal
End Sub

Private Sub mnuABMCanales_Click()
    ABMCanal.Show vbModal
End Sub

Private Sub mnuABMClientes_Click()
    ABMCliente.Show vbModal
End Sub

Private Sub mnuABMConceptoNotaCredito_Click()
    ABMConceptoNotaCredito.Show vbModal
End Sub

Private Sub mnuABMEstadoDocumento_Click()
    ABMEstadoDocumento.Show vbModal
End Sub

Private Sub mnuABMFormaPago_Click()
    ABMFormaPago.Show vbModal
End Sub

Private Sub mnuABMInscIVA_Click()
    ABMInscIVA.Show vbModal
End Sub

Private Sub mnuABMLineas_Click()
    ABMLinea.Show vbModal
End Sub

Private Sub mnuABMLocalidades_Click()
    ABMLocalidad.Show vbModal
End Sub

Private Sub mnuABMMoneda_Click()
     ABMMoneda.Show vbModal
End Sub

Private Sub mnuABMPais_Click()
    ABMPais.Show vbModal
End Sub

Private Sub mnuABMPresentacion_Click()
    ABMPresentacion.Show vbModal
End Sub

Private Sub mnuABMProveedores_Click()
    ABMProveedor.Show vbModal
End Sub

Private Sub mnuABMProvincias_Click()
    ABMProvincia.Show vbModal
End Sub

Private Sub mnuABMRepresentada_Click()
    'ABMRepresentada.Show vbModal
End Sub

Private Sub mnuABMRubros_Click()
    ABMRubro.Show vbModal
End Sub

Private Sub mnuABMServicios_Click()
    ABMServicios.Show vbModal
End Sub

Private Sub mnuABMSucursales_Click()
    ABMSucursal.Show vbModal
End Sub

Private Sub mnuABMTipoCuentas_Click()
    ABMTipoCuenta.Show vbModal
End Sub

Private Sub mnuABMTipoGastos_Click()
    ABMTipoGasto.Show vbModal
End Sub

Private Sub mnuABMTipoProveedores_Click()
    ABMTipoProveedor.Show vbModal
End Sub

Private Sub mnuABMTiposGastosBancarios_Click()
    ABMTipoGastoBancario.Show vbModal
End Sub

Private Sub mnuABMTransporte_Click()
    ABMTransporte.Show vbModal
End Sub

Private Sub mnuABMVendedores_Click()
    ABMVendedor.Show vbModal
End Sub

Private Sub mnuacercade_Click()
    Call ShellAbout(Me.hwnd, "Mi Programa", "Copyright 2002, PMMF", Me.Icon)
End Sub

Private Sub mnuAnulaNC_Click()
    frmAnulaDocumentos.TipodeAnulacion = 10
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuAnularGastos_Click()
    frmAnulaDocumentos.TipodeAnulacion = 9
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuayuda_Click()
    MsgBox "La Ayuda no está disponible en estos momentos", vbExclamation, TIT_MSGBOX
'    Call WinHelp(Me.hwnd, App.Path & "\Pecari.hlp", HelpFinder, 0&)
End Sub

Private Sub mnuBackupBD_Click()
    With frmRestaurarBD
        .Caption = "Backup de Base de Datos"
        .optCopiarA.Value = True
        .Label1 = "Guardar Backup en:"
        .Show
    End With
End Sub

Private Sub mnuBoletaDeposito_Click()
    FrmBoletaDeposito.Show vbModal
End Sub

Private Sub mnuCalculadora_Click()
    On Error Resume Next
    Shell "C:\WINDOWS\system32\calc.exe", vbNormalFocus
    'Form1.Show vbModal
End Sub

Private Sub mnuCalendario_Click()
    frmCalendario.Show
End Sub

Private Sub mnuCambioEstadoChequesPropios_Click()
    ABMCambioEstadoChPropio.Show vbModal
End Sub

Private Sub mnuCobranzaClientes_Click()
    frmListadoCobranzaCliente.Show vbModal
End Sub

Private Sub mnuComisionesVenderores_Click()
    ABMComision.Show vbModal
End Sub

Private Sub mnuComprasCreditoFiscal_Click()
'    frmLibroIvaCompras.Show
    frmLibroCompras2.Show vbModal
End Sub

Private Sub mnuComprasCtaCte_Click()
    frmCtaCteProveedores.Show vbModal
End Sub

Private Sub mnuComprasListaResumen_Click()
    frmListadoComprasProveedores.Show vbModal
End Sub

Private Sub mnuConAnuFactura_Click()
    frmAnulaDocumentos.TipodeAnulacion = 3
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuConAnuPedidos_Click()
    frmAnulaDocumentos.TipodeAnulacion = 1
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuConAnuRecibos_Click()
    frmAnulaDocumentos.TipodeAnulacion = 4
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuConAnuRemitos_Click()
    frmAnulaDocumentos.TipodeAnulacion = 2
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuConectar_Click()
    FrmInicio.Show vbModal
    Me.Caption = TituloPrincipal & " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & ")"
    Me.mnuConectar.Enabled = False
End Sub

Private Sub mnuConsListaPrecios_Click()
    Consulta = 1
    FrmListadePrecios.Show vbModal
End Sub

Private Sub mnuConsultaStock_Click()

End Sub

Private Sub mnuConStock_Click()
    Stock = 1 'cosnulta stock
    
    frmControlStock.Show
End Sub

Private Sub mnuDesconectar_Click()
    If DBConn.State = adStateOpen Then
        DBConn.Close
        
        DeshabilitarMenu Me
        
        Me.mnuSistema.Enabled = True
        Me.mnuConectar.Enabled = True
        Me.mnusalir.Enabled = True
        Me.mnuDesconectar.Enabled = False
        
        Me.Caption = TituloPrincipal & " - (No conectado)"
    End If
End Sub

Private Sub MDIForm_Load()
    TituloPrincipal = "Sistema de Gestión y Administración"
    Me.Caption = TituloPrincipal
    
    Me.Show
    FrmInicio.Show vbModal
       
    Me.Caption = TituloPrincipal & " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & " - " & UCase(BASEDATO) & ")"
    MENU.mnuConectar.Enabled = False
End Sub

Private Sub Cargar(formu As Form, Optional modalmente As Integer)
    On Error GoTo CLAVOSE
    formu.Top = 0
    formu.Left = 0
    Load formu
    formu.Show modalmente
    Exit Sub
    
CLAVOSE:
    MsgBox "Ha ocurrido un  error al tratar de cargar el formulario !" & Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub mnuEntradaProductos_Click()
    frmEntradaProductos.Show vbModal
End Sub

Private Sub mnuEntregaDeProductos_Click()
    frmEntregaProductos.Show vbModal
End Sub

Private Sub mnuEstaCantidadVendida_Click()
    frmListadoCantidadesVendidas.Show vbModal
End Sub

Private Sub mnuEstaVentaporCliente_Click()
    frmListadoVentasPorCliente.Show vbModal
End Sub

Private Sub mnuFacturacionCtaCte_Click()
    frmCtaCteCliente.Show vbModal
End Sub

Private Sub mnuFacturacionFactura_Click()
    frmFacturaCliente.Show vbModal
End Sub

Private Sub mnuFacturacionTipoComprobante_Click()
    ABMTipoComprobante.Show vbModal
End Sub

Private Sub mnufacturas_Click()
    frmAnulaDocumentos.TipodeAnulacion = 7
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuFacturasPendientesCliente_Click()
    frmListadoFacturasPendientePorCliente.Show vbModal
End Sub

Private Sub mnuFacturasPendientesProveedor_Click()
    frmListadoFacturasPendientePorProveedor.Show vbModal
End Sub

Private Sub mnuFondosBancos_Click()
    ABMBanco.Show vbModal
End Sub

Private Sub mnuFondosCambioEstadoChques_Click()
    ABMCambioEstado.Show vbModal
End Sub

Private Sub mnuFondosCargaCheques_Click()
    FrmCargaCheques.Show vbModal
End Sub

Private Sub mnuFondosCargaChequesPropios_Click()
    FrmCargaChequesPropios.Show vbModal
End Sub

Private Sub mnuFondosCargaEgresos_Click()
    ABMEgresos.Show vbModal
End Sub

Private Sub mnuFondosCargaIngresos_Click()
    'ABMIngresos.Show vbModal
End Sub

Private Sub mnuFondosCuentas_Click()
    ABMCuentasBancarias.Show vbModal
End Sub

Private Sub mnuFondosEstadoCheques_Click()
    ABMEstadoCheques.Show vbModal
End Sub

Private Sub mnuFondosListadoCheques_Click()
    FrmListCheques.Show vbModal
End Sub

Private Sub mnuFondosResumenCuneta_Click()
    frmResumenCuentaBanco.Show vbModal
End Sub

Private Sub mnuFonfosCierreCaja_Click()
    frmCierreCaja.Show
End Sub

Private Sub mnuGastosGeneralesRegistro_Click()
    frmCargaGastosGenerales.Show vbModal
End Sub

Private Sub mnuImputarNCFacturas_Click()
'    frmImputarNCaFactura.Show vbModal
End Sub

Private Sub mnuimputaNCFAC_Click()
    frmImputarNCaFactura.Show vbModal
End Sub

Private Sub mnuImputarNCFacturasProveedores_Click()
    frmImputarNCaFacturaProveedores.Show vbModal
End Sub

Private Sub mnuIngNotaPedidoCliente_Click()
    frmNotaDePedido.Show vbModal
End Sub

Private Sub mnuIngRecibosClientes_Click()
    frmReciboCliente.Show vbModal
End Sub

Private Sub mnuIngRemitosClientes_Click()
    frmRemitoCliente.Show vbModal
End Sub

Private Sub mnuIngresoGastosBancarios_Click()
'    frmIngresoGastosBancarios.Show vbModal
End Sub

Private Sub mnuLibroIvaVentas_Click()
    'frmLibroIvaVentas.Show
    frmLibroVentas2.Show vbModal
End Sub

Private Sub mnuLiquidacionCobranza_Click()
    'frmLiquidacionCobranza.Show vbModal
End Sub

Private Sub mnuListadoChequesPropios_Click()
    FrmListChequesPropios.Show vbModal
End Sub

Private Sub mnuListadoClientes_Click()
    frmListadoClientes.Show vbModal
End Sub

Private Sub mnuListadoGastosGrales_Click()
    frmListadoGastosGrales.Show vbModal
End Sub

Private Sub mnuListadoNotaPedidoCliente_Click()
    frmListadoNotaDePedido.Show vbModal
End Sub

Private Sub mnuListadoOrdCompra_Click()
    frmListadoOrdenCompra.Show vbModal
End Sub

Private Sub mnuListadoProveedores_Click()
    frmListadoProvedores.Show vbModal
End Sub

Private Sub mnuListadoRecibosClientes_Click()
    frmListadoReciboCliente.Show vbModal
End Sub

Private Sub mnuListadoRemitosClientes_Click()
    frmListadoRemitoCliente.Show vbModal
End Sub

Private Sub mnuListadoSucursalesCliente_Click()
    frmListadoSucursalesCliente.Show vbModal
End Sub

Private Sub mnuListadoRemProv_Click()
    frmListadoRemitoProveedor.Show vbModal
End Sub

Private Sub mnuListadoVentaPorVendedor_Click()
    frmListadoVentasPorVendedor.Show vbModal
End Sub

Private Sub mnuListaPrecios_Click()
    Consulta = 2
    FrmListadePrecios.Show vbModal
End Sub

Private Sub mnuMovimientoStock_Click()

End Sub

Private Sub mnuNDProvCheque_Click()
    frmNotaDeditoProvCheques.Show vbModal
End Sub

Private Sub mnuNDProvServicio_Click()
    frmNotaDebitoProveedores.Show vbModal
End Sub

Private Sub mnuNotaCredito_Click()
    frmNotaCreditoCliente.Show vbModal
End Sub

Private Sub mnuNotaCreditoProveedor_Click()
    frmNotaCreditoProveedor.Show vbModal
End Sub

Private Sub mnuNotaDebitoPorChequeDevuelto_Click()
    frmNotaDeditoClienteCheques.Show vbModal
End Sub

Private Sub mnuNotaDebitoPorServicios_Click()
    frmNotaDeditoClienteServicio.Show vbModal
End Sub

Private Sub mnuOrdenCompra_Click()
    frmAnulaDocumentos.TipodeAnulacion = 5
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuOrdendeCompra_Click()
    frmOrdenesCompra.Show vbModal
End Sub

Private Sub mnuPagoProveedoresAnularPago_Click()
    frmAnulaOrdenPago.Show vbModal
End Sub

Private Sub mnuOrdenesdePago_Click()
    frmAnulaDocumentos.TipodeAnulacion = 8
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuPagoProveedoresOrdenPago_Click()
    frmPagoProveedores.Show vbModal
End Sub

Private Sub mnuPagosRealizadosPorCliente_Click()
    frmListadoPagosPorCliente.Show vbModal
End Sub

Private Sub mnuPagosRealizadosProveedores_Click()
    frmListadoPagosPorProveedor.Show vbModal
End Sub

Private Sub mnuParametros_Click()
    frmParametros.Show vbModal
End Sub

Private Sub mnupermisos_Click()
    FrmPermisos.Show vbModal
End Sub

Private Sub mnuProveedoresFacturas_Click()
    frmfacturaproveedor.Show vbModal
End Sub

Private Sub mnuRemitoCompras_Click()
    frmRemitoProveedor.Show vbModal
End Sub

Private Sub mnuremitos_Click()
    frmAnulaDocumentos.TipodeAnulacion = 6
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuRestaurarBD_Click()
    With frmRestaurarBD
        .Caption = "Restaurar Base de Datos"
        .optCopiarDesde.Value = True
        .Label1 = "Restaurar desde:"
        .Show
    End With
End Sub

Private Sub mnusalir_Click()
   Set MENU = Nothing
   End
End Sub

Private Sub mnuStockABMProductos_Click()
    Consulta = 1
    ABMProducto.CODIGOLISTA = 0
    ABMProducto.Show vbModal
End Sub

Private Sub mnuStockAjuste_Click()
    Stock = 0 'Actualiza el Stock
    frmControlStock.Show vbModal
End Sub

Private Sub mnuusuario_Click()
     Cargar FrmUsuarios, 1
End Sub

Private Sub tbrPrincipal_ButtonClick(ByVal Button As ComctlLib.Button)
        Select Case Button.Index
        Case 2: Call mnuRemitosdirecto_Click
        Case 3: Call mnuClientesdirecto_Click
        Case 4: Call mnuFacturacionFactura_Click
        Case 5: Call mnuListadoClientes_Click
        'Case 4: Separador
        Case 7: Call mnuRemitosProvdirecto_Click
        Case 8: Call mnuABMProveedores_Click
        'Case 9: Separador
        Case 10: Call mnuProductosdirecto_Click
        Case 11: Call mnuListaPrecios_Click
        Case 12: Call mnuStockFaltantes_Click
        'Case 9: Separador
        Case 14: Call cmdListados_Click
        Case 16: Call mnusalir_Click
        End Select
End Sub
Private Sub mnuRemitosProvdirecto_Click()
    frmRemitoProveedor.Show vbModal
End Sub
Private Sub mnuClientesdirecto_Click()
    ABMCliente.Show vbModal
End Sub
Private Sub mnuProductosdirecto_Click()
    ABMProducto.Show vbModal
End Sub
Private Sub mnuRemitosdirecto_Click()
    frmRemitoCliente.Show
End Sub
Private Sub mnuStockFaltantes_Click()
    frmListadoStock.Show vbModal
End Sub

