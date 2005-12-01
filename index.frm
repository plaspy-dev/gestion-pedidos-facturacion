VERSION 5.00
Begin VB.Form index 
   BackColor       =   &H8000000C&
   Caption         =   "Pescaderia Martinez Martelo & CIA"
   ClientHeight    =   7365
   ClientLeft      =   165
   ClientTop       =   570
   ClientWidth     =   13125
   LinkTopic       =   "Form2"
   ScaleHeight     =   7365
   ScaleWidth      =   13125
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   7455
      Left            =   0
      Picture         =   "index.frx":0000
      ScaleHeight     =   7395
      ScaleWidth      =   13395
      TabIndex        =   0
      Top             =   0
      Width           =   13455
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Begin VB.Menu clientes 
         Caption         =   "Facturaciòn"
         Begin VB.Menu facturacli 
            Caption         =   "Factura Clientes"
         End
         Begin VB.Menu verfactu 
            Caption         =   "Verificar Factura"
         End
      End
      Begin VB.Menu vendedor 
         Caption         =   "Ventas"
         Begin VB.Menu pedido 
            Caption         =   "Pedido"
         End
         Begin VB.Menu vped 
            Caption         =   "Verificar pedido"
         End
         Begin VB.Menu regvendedor 
            Caption         =   "Registro de Vendedor"
         End
         Begin VB.Menu regcli 
            Caption         =   "Registro De Clientes"
         End
      End
      Begin VB.Menu proveedor 
         Caption         =   "Inventario"
         Begin VB.Menu regpro 
            Caption         =   "Registro de proveedor"
         End
         Begin VB.Menu regproducto 
            Caption         =   "Registro de Producto "
         End
         Begin VB.Menu reginventario 
            Caption         =   "Registro de Inventario"
         End
      End
      Begin VB.Menu inventario 
         Caption         =   "Reportes"
         Begin VB.Menu reporven 
            Caption         =   "Reporte de Ventas  "
         End
         Begin VB.Menu repinventario 
            Caption         =   "Reporte de Inventario"
         End
      End
   End
   Begin VB.Menu ayuda 
      Caption         =   "Ayuda"
   End
   Begin VB.Menu salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub facturacli_Click()
facturas.Show
End Sub
Private Sub facturapro_Click()
facturas1.Show
End Sub
Private Sub pedido_Click()
pedidos.Show
End Sub
Private Sub regcli_Click()
cliente.Show
End Sub
Private Sub reginventario_Click()
detalleprod.Show
End Sub
Private Sub regpro_Click()
proveedores.Show
End Sub
Private Sub regproducto_Click()
productos.Show
End Sub
Private Sub regvendedor_Click()
vendedores.Show
End Sub
Private Sub salir_Click()
End
End Sub
Private Sub verfactu_Click()
vfactura.Show
End Sub

Private Sub vped_Click()
vpedido.Show
End Sub
