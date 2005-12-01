VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmaviso 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6135
   ClientLeft      =   5610
   ClientTop       =   4770
   ClientWidth     =   8055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   120
      Picture         =   "frmaviso.frx":0000
      ScaleHeight     =   5475
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UN MOMENTO POR FAVOR........."
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   3840
         TabIndex        =   2
         Top             =   4080
         Width           =   3885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "EL PROGRAMA ESTA CARGANDO"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   2400
         Width           =   3975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmaviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
   Static Percent As Long
   
   Percent = Percent + 35
   If Percent > 100 Then
      Timer1.Enabled = False
      Percent = 0
   
          frmaviso.Hide
    login.Show
    Unload Me
    Timer1.Interval = 0
   End If
   Me.ProgressBar1 = Percent
   

End Sub
