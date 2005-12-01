VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form facturas 
   BackColor       =   &H00808000&
   Caption         =   "factura de Clientes "
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   5280
      TabIndex        =   16
      Top             =   3600
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2640
      TabIndex        =   15
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   5040
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">>"
      Height          =   285
      Left            =   6720
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   5280
      TabIndex        =   10
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   2880
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6600
      TabIndex        =   6
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7200
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   6840
      TabIndex        =   4
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Almacenar"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   7680
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   0
      Picture         =   "factura.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   9195
      TabIndex        =   1
      Top             =   0
      Width           =   9255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "factura.frx":386F
      Left            =   2160
      List            =   "factura.frx":387C
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   240
      TabIndex        =   30
      Top             =   4800
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Producto"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Peso"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Precio unitario"
         Object.Width           =   2929
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2929
      EndProperty
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   29
      Top             =   3960
      Width           =   1620
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4080
      TabIndex        =   28
      Top             =   3600
      Width           =   1035
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo del Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   27
      Top             =   3600
      Width           =   2235
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Peso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4200
      TabIndex        =   26
      Top             =   4320
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo del Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   25
      Top             =   4320
      Width           =   2145
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   24
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5520
      TabIndex        =   23
      Top             =   3240
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   22
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4320
      TabIndex        =   21
      Top             =   2520
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3840
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   19
      Top             =   2520
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FACTURA NO."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5520
      TabIndex        =   18
      Top             =   2040
      Width           =   1530
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5400
      TabIndex        =   17
      Top             =   7080
      Width           =   1170
   End
End
Attribute VB_Name = "facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Text1_Change()
Me.Text1.SetFocus
End Sub
