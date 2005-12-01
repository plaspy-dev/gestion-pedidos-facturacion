VERSION 5.00
Begin VB.Form cliente 
   BackColor       =   &H00808000&
   Caption         =   "Datos del Cliente"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   8085
   LinkTopic       =   "Form4"
   ScaleHeight     =   6285
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   3120
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   2640
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   -1320
      Picture         =   "clientes.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   9795
      TabIndex        =   3
      Top             =   0
      Width           =   9855
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   5040
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
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
      Left            =   1320
      TabIndex        =   18
      Top             =   5040
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
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
      Left            =   1320
      TabIndex        =   17
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
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
      Left            =   1320
      TabIndex        =   16
      Top             =   4080
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido"
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
      Left            =   1320
      TabIndex        =   15
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   1320
      TabIndex        =   14
      Top             =   2640
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Del Cliente"
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
      Left            =   1320
      TabIndex        =   13
      Top             =   2160
      Width           =   2250
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad"
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
      Left            =   1320
      TabIndex        =   12
      Top             =   3600
      Width           =   855
   End
End
Attribute VB_Name = "cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As ADODB.Connection
Public rs As ADODB.Recordset

Private Sub Command1_Click()
 Dim aa
 aa = MsgBox("Desea almacenar el registro?", vbYesNo)
 If aa = vbYes Then
  Set rs = cn.Execute("insert into cliente ( id_cliente, nomcliente, apecliente, ciudad, direccliente, telcliente, mailcliente) values ('" & Trim(Me.Text1) & "', '" & Trim(Me.Text2) & "','" & Trim(Me.Text3) & "',  '" & Trim(Me.Text4) & "', '" & Trim(Me.Text5) & "', '" & Trim(Me.Text6) & "', '" & Trim(Me.Text7) & "')")
  MsgBox "Registro almacenado", 64
  limpiar
  bloquear
 Else
  MsgBox "Registro no almacenado", 64
  limpiar
  bloquear
 End If
End Sub

Private Sub Command2_Click()
 Dim aa
 aa = MsgBox("Desea actualizar el registro?", vbYesNo)
 If aa = vbYes Then
  Set rs = cn.Execute("update cliente set nomcliente  = '" & Me.Text2 & "', apecliente = '" & Me.Text3 & "', ciudad = '" & Me.Text4 & "', direccliente = '" & Me.Text5 & "', telcliente = '" & Me.Text6 & "', mailcliente = '" & Me.Text7 & "' where id_cliente like '" & Me.Text1 & "' ")
  MsgBox "El registro fue actualizado", 64
  Me.Text1.SetFocus
  limpiar
  bloquear
 Else
  MsgBox "Registro no actualizado", 64
  limpiar
  bloquear
 End If
End Sub

Private Sub Command3_Click()
 Dim aa
 aa = MsgBox("¿desea borrar el registro?", vbYesNo)
 If aa = vbYes Then
  Set rs = cn.Execute("delete from cliente where id_cliente like '" & Trim(Me.Text1) & "' ")
  MsgBox "El registro fue Borrado", 64
  Me.Text1.SetFocus
  limpiar
 Else
  MsgBox "el registro no fue borrado", 64
  limpiar
  Me.Text1.SetFocus
  bloquear
 End If
End Sub

Private Sub Command4_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Set cn = New ADODB.Connection
 cn.Open "dsn=pmm"
 cn.CommandTimeout = 300
 bloquear
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Trim(Me.Text1) = "" Then
   Me.Text1.SetFocus
  Else
   Set rs = New ADODB.Recordset
   rs.LockType = adLockBatchOptimistic
   rs.CursorType = adOpenKeyset
   rs.Source = ("select * from cliente where id_cliente like '" & Trim(Me.Text1) & "'")
   rs.ActiveConnection = cn
   rs.Open
   If rs.RecordCount > 0 Then
    Me.Text2 = rs("nomcliente")
    Me.Text3 = rs("apecliente")
    Me.Text4 = rs("ciudad")
    Me.Text5 = rs("direccliente")
    Me.Text6 = rs("telcliente")
    Me.Text7 = rs("mailcliente")
    bloquear
    Me.Command2.Enabled = True
    Me.Command3.Enabled = True
    Me.Text2.Enabled = True
    Me.Text2.SetFocus
    
   Else
    MsgBox "Cliente no existente, crear una cuenta nueva", 32
    bloquear
    Me.Text2.Enabled = True
    Me.Text2.SetFocus
    Me.Text2 = ""
    Me.Text3 = ""
    Me.Text4 = ""
    Me.Text5 = ""
    Me.Text6 = ""
    Me.Text7 = ""
   End If
  End If
 End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Trim(Me.Text2) = "" Then
   Me.Text2.SetFocus
  Else
   Me.Text3.Enabled = True
   Me.Text3.SetFocus
  End If
 End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Trim(Me.Text3) = "" Then
   Me.Text3.SetFocus
  Else
   Me.Text4.Enabled = True
   Me.Text4.SetFocus
  End If
 End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Trim(Me.Text4) = "" Then
   Me.Text4.SetFocus
  Else
   Me.Text5.Enabled = True
   Me.Text5.SetFocus
  End If
 End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Trim(Me.Text5) = "" Then
   Me.Text5.SetFocus
  Else
   Me.Text6.Enabled = True
   Me.Text6.SetFocus
  End If
 End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Trim(Me.Text6) = "" Then
   Me.Text6.SetFocus
  Else
   Me.Text7.Enabled = True
   Me.Text7.SetFocus
  End If
 End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Me.Command1.Enabled = True
  Me.Command1.SetFocus
 End If
End Sub

Public Sub limpiar()
 Me.Text1 = ""
 Me.Text2 = ""
 Me.Text3 = ""
 Me.Text4 = ""
 Me.Text5 = ""
 Me.Text6 = ""
 Me.Text7 = ""
End Sub

Public Sub bloquear()
 Me.Text2.Enabled = False
 Me.Text3.Enabled = False
 Me.Text4.Enabled = False
 Me.Text5.Enabled = False
 Me.Text6.Enabled = False
 Me.Text7.Enabled = False
 Me.Command1.Enabled = False
 Me.Command2.Enabled = False
 Me.Command3.Enabled = False
End Sub





