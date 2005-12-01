VERSION 5.00
Begin VB.Form detalleprod 
   BackColor       =   &H00808000&
   Caption         =   "Detalle Producto"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3360
      TabIndex        =   12
      Top             =   4800
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   -600
      Picture         =   "detallepro.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   9795
      TabIndex        =   11
      Top             =   0
      Width           =   9855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3360
      TabIndex        =   10
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox producto 
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Top             =   2880
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Por Kilo "
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
      Left            =   720
      TabIndex        =   15
      Top             =   4320
      Width           =   2160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor "
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
      Left            =   720
      TabIndex        =   14
      Top             =   4800
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor "
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
      Left            =   720
      TabIndex        =   9
      Top             =   3840
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Por Kilo "
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
      Left            =   720
      TabIndex        =   8
      Top             =   3360
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Producto"
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
      Left            =   720
      TabIndex        =   6
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Del Producto"
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
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Width           =   2490
   End
End
Attribute VB_Name = "detalleprod"
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
  Set rs = cn.Execute("insert into producto ( Id_producto, nomproduc, cantproduc, valor ) values ('" & Trim(Me.producto) & "','" & Trim(Me.Text2) & "', '" & Trim(Me.Text3) & "', '" & Trim(Me.Text4) & "')")
  MsgBox "Registro almacenado", 64
  Me.producto = ""
  Me.Text2 = ""
  Me.Text3 = ""
  Me.Text4 = ""
  bloquear
  Else
  MsgBox "Registro no almacenado", 64
 End If
End Sub

Private Sub Command2_Click()
Dim aa
 aa = MsgBox("Desea almacenar el registro?", vbYesNo)
 If aa = vbYes Then
  Set rs = cn.Execute("update producto set nomproduc= '" & Me.Text2 & "',cantproduc = '" & Trim(Me.Text3) & "', valor = '" & Trim(Me.Text4) & "' where Id_producto like '" & Me.producto & "' ")
  MsgBox "Registro almacenado", 64
limpiar
Me.producto.SetFocus
  Else
  MsgBox "Registro no almacenado", 64
  Me.producto.SetFocus
 End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Me.producto.Enabled = True
Me.producto.SetFocus
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
 cn.Open "dsn=pmm"
 cn.CommandTimeout = 300
 bloquear
 Me.Command4.Enabled = True
End Sub
Private Sub producto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Trim(Me.producto) = "" Then
   Me.producto.SetFocus
  Else
   Set rs = New ADODB.Recordset
   rs.LockType = adLockBatchOptimistic
   rs.CursorType = adOpenKeyset
   rs.Source = ("select * from producto where Id_producto like '" & Trim(Me.producto) & "'")
   rs.ActiveConnection = cn
   rs.Open
   If rs.RecordCount > 0 Then
    Me.Text2 = rs("nomproduc")
    Me.Text3 = rs("cantproduc")
    Me.Text4 = rs("valor")
    Me.Command2.Enabled = True
    Me.Text5.Enabled = True
    Me.Text5.SetFocus
   Else
    MsgBox " el producto no existente ", 32
    
    Me.Text3.Enabled = True
    Me.Text3.SetFocus
    Set rs = New ADODB.Recordset
   rs.LockType = adLockBatchOptimistic
   rs.CursorType = adOpenKeyset
   rs.Source = ("select * from producto1 where Id_producto like '" & Trim(Me.producto) & "'")
   rs.ActiveConnection = cn
   rs.Open
   If rs.RecordCount > 0 Then
    Me.Text2 = rs("nombreproduc")
    Else
    Me.producto.SetFocus
    End If
    End If
  End If
 End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim c, b, f
If KeyAscii = 13 Then
  If Trim(Me.Text1) = "" Then
   Me.Text1.SetFocus
   Else
   c = Val(Me.Text1)
   b = Val(Me.Text4)
   If b = c Then
   a = Val(Me.Text4)
   Else
    If b < c Then
    c = Val(Me.Text1) * 0.05
    f = b + c
    Me.Text4 = f
    Else
      If b > c Then
      c = Val(Me.Text1) * 0.05
      f = b - c
      Me.Text4 = f
      End If
    End If
   End If
   Me.Command2.Enabled = True
   Me.Command2.SetFocus
  End If
 End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Trim(Me.Text2) = "" Then
   Me.Text2.SetFocus
  Else
   Me.Command1.Enabled = True
   Me.Command1.SetFocus
  End If
 End If
End Sub
Public Sub limpiar()
Me.producto = ""
Me.Text2 = ""
Me.Text3 = ""
Me.Text4 = ""
Me.Text5 = ""
Me.Text1 = ""
End Sub
Public Sub bloquear()
 Me.Text1.Enabled = False
 Me.Text2.Enabled = False
 Me.Text3.Enabled = False
 Me.Text4.Enabled = False
 Me.Text5.Enabled = False
 Me.Command1.Enabled = False
 Me.Command2.Enabled = False
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
   Me.Command1.Enabled = True
   Me.Command1.SetFocus
  End If
 End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Trim(Me.Text5) = "" Then
   Me.Text5.SetFocus
   Else
   a = Val(Me.Text3) + Val(Me.Text5)
   Me.Text3 = a
   Me.Text1.Enabled = True
   Me.Text1.SetFocus
  End If
 End If
End Sub
