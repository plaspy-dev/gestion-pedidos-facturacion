VERSION 5.00
Begin VB.Form productos 
   BackColor       =   &H00808000&
   Caption         =   "producto"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   8850
   LinkTopic       =   "Form5"
   ScaleHeight     =   4380
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   -600
      Picture         =   "producto.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   9795
      TabIndex        =   8
      Top             =   0
      Width           =   9855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   2880
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox producto 
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   2400
      Width           =   2490
   End
End
Attribute VB_Name = "productos"
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
  Set rs = cn.Execute("insert into producto1 ( Id_producto, nombreproduc) values ('" & Trim(Me.producto) & "', '" & Trim(Me.Text2) & "')")
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
 aa = MsgBox("Desea Actualizar el registro?", vbYesNo)
 If aa = vbYes Then
  Set rs = cn.Execute("update producto1 set Id_producto= '" & Me.producto & "', nombreproduc= '" & Me.Text2 & "' ")
  MsgBox "Registro actualizado", 64
  limpiar
  bloquear
 Else
  MsgBox "Registro no actualizado", 64
 End If
 Me.Command2.Enabled = False
End Sub

Private Sub Command3_Click()
 Dim aa
 aa = MsgBox("¿desea borrar el registro?", vbYesNo)
 If aa = vbYes Then
  Set rs = cn.Execute("delete from producto1 where Id_producto like '" & Trim(Me.producto) & "' ")
  MsgBox "El registro fue Borrado", 64
  Me.producto.SetFocus
  limpiar
  bloquear
 Else
  MsgBox "el registro no fue borrado", 64
  Me.Command3.Enabled = False
 End If
End Sub

Private Sub Command4_Click()
 Unload Me
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
 ' cn.Open "DSN=MS Access Database;DBQ=pmm.mdb;DefaultDir=.\;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;"
 cn.Open "dsn=pmm"
 cn.CommandTimeout = 300
 bloquear
End Sub

Private Sub producto_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Trim(Me.producto) = "" Then
   Me.producto.SetFocus
  Else
   Set rs = New ADODB.Recordset
   rs.LockType = adLockBatchOptimistic
   rs.CursorType = adOpenKeyset
   rs.Source = ("select * from producto1 where Id_producto like '" & Trim(Me.producto) & "'")
   rs.ActiveConnection = cn
   rs.Open
   If rs.RecordCount > 0 Then
    Me.Text2.Enabled = True
    Me.Text2 = rs("nombreproduc")
    Me.Command2.Enabled = True
    Me.Command3.Enabled = True
        bloquear
    Me.Command2.Enabled = True
    Me.Command3.Enabled = True
    Me.Text2.Enabled = True
    Me.Text2.SetFocus
   Else
    MsgBox "El producto no exite, registrar producto", 32
    Me.Text2 = ""
    bloquear
    Me.Text2.Enabled = True
    Me.Text2.SetFocus
   End If
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
 Me.Text2 = ""
 Me.producto = ""
End Sub
Public Sub bloquear()
 Me.Text2.Enabled = False
 Me.Command1.Enabled = False
 Me.Command2.Enabled = False
 Me.Command3.Enabled = False
End Sub



