VERSION 5.00
Begin VB.Form usuario 
   BackColor       =   &H00808000&
   Caption         =   "Registro de Usuario"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   8040
   LinkTopic       =   "Form6"
   ScaleHeight     =   4755
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   -960
      Picture         =   "usuario.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   9795
      TabIndex        =   8
      Top             =   0
      Width           =   9855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   1170
   End
End
Attribute VB_Name = "usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

