VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00808000&
   Caption         =   "login"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   8160
   LinkTopic       =   "Form7"
   ScaleHeight     =   3915
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   -960
      Picture         =   "login.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   9795
      TabIndex        =   6
      Top             =   0
      Width           =   9855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3480
      TabIndex        =   1
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2760
      Width           =   3735
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
      Left            =   840
      TabIndex        =   3
      Top             =   2760
      Width           =   1170
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
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   2385
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 index.Show
 Unload Me
End Sub

Private Sub Command2_Click()
 End
End Sub

Private Sub Form_Load()
 Me.Text2.Enabled = False
 Me.Command1.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Trim(Me.Text1) = "" Then
   Me.Text1.SetFocus
  Else
   Me.Text2.Enabled = True
   Me.Text2.SetFocus
  End If
 End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Me.Command1.Enabled = True
  Me.Command1.SetFocus
 End If
End Sub


