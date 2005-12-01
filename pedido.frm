VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form pedidos 
   BackColor       =   &H00808000&
   Caption         =   "Pedido"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   8910
   LinkTopic       =   "Form8"
   ScaleHeight     =   8355
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "date"
      Height          =   255
      Left            =   4080
      TabIndex        =   34
      Top             =   6840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   7080
      TabIndex        =   18
      Top             =   2040
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "pedido.frx":0000
      Left            =   6720
      List            =   "pedido.frx":000D
      TabIndex        =   17
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   5280
      TabIndex        =   15
      Top             =   7200
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,000000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   14
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   12
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">>"
      Height          =   285
      Left            =   7440
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   6240
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5040
      TabIndex        =   7
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   6840
      TabIndex        =   4
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   7800
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   0
      Picture         =   "pedido.frx":002B
      ScaleHeight     =   1875
      ScaleWidth      =   9195
      TabIndex        =   1
      Top             =   0
      Width           =   9255
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   240
      TabIndex        =   19
      Top             =   4080
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   4500
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Peso"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Precio unitario"
         Object.Width           =   2929
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
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
      Left            =   4920
      TabIndex        =   33
      Top             =   6840
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
      TabIndex        =   32
      Top             =   7200
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
      TabIndex        =   31
      Top             =   7200
      Width           =   2235
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Despacho"
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
      TabIndex        =   30
      Top             =   6840
      Width           =   2115
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
      Left            =   3960
      TabIndex        =   29
      Top             =   3600
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
      TabIndex        =   28
      Top             =   3600
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
      TabIndex        =   27
      Top             =   2520
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
      Left            =   5160
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
      Top             =   2880
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
      TabIndex        =   23
      Top             =   2520
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
      TabIndex        =   22
      Top             =   2880
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PEDIDO NO."
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
      Left            =   5640
      TabIndex        =   21
      Top             =   2040
      Width           =   1320
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
      TabIndex        =   20
      Top             =   6360
      Width           =   1170
   End
   Begin VB.Menu Lista 
      Caption         =   "Lista"
      Visible         =   0   'False
      Begin VB.Menu Editar 
         Caption         =   "&Editar"
      End
      Begin VB.Menu Eliminar 
         Caption         =   "E&liminar"
      End
   End
End
Attribute VB_Name = "pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public IdPedido As String
Public IndexProducto As Integer

Sub Elimina(index As Integer)
    
    If MsgBox("Desea eliminar '" + ListView1.SelectedItem.SubItems(1) + "'", vbYesNo) = vbYes Then
        If ListView1.SelectedItem.Tag <> "" And IdPedido <> "" Then
            ' lo borramos de la base de datos
            Set rs = cn.Execute("DELETE FROM Detallepedido WHERE id_ordenpedido LIKE '" + IdPedido + "' AND Id_producto LIKE '" + Str(ListView1.SelectedItem.Tag) + "'")
        End If
        ListView1.ListItems.Remove (ListView1.SelectedItem.index)
        RecalculaTotal
    End If
End Sub

Function GetNumero(numero As String) As Double
    ' cambiar comas por puntos
    GetNumero = Val(Replace(numero, ",", "."))
End Function

Function verifica(codigo As String) As Boolean
    For i = 1 To ListView1.ListItems.Count
        If codigo = ListView1.ListItems.Item(i).Text Then
            verifica = True
            Exit Function
        End If
    Next i
    verifica = False
End Function


Sub RecalculaTotal()
    Dim total, subtotal As Double
    For i = 1 To ListView1.ListItems.Count
        ' el replace es debido a que los valores tienen comas en vez de puntos
        ' en caso contrario no pasa nada
        total = total + GetNumero(ListView1.ListItems.Item(i).ListSubItems(4))
    Next i
    Me.Text12.Text = Str(total)
End Sub

Function GetString(numero As String) As String
    ' cambiar puntos a comas para que lo acepte la base de datos
    GetString = Replace(numero, ".", ",")
End Function

Private Sub Combo1_Click()
    Me.Text14.Enabled = True
    Me.Text14.SetFocus
End Sub

Private Sub Command1_Click()

    If Text1.Text = "" Then
       MsgBox "El codigo de pedido no puede estar en blanco"
       Text1.SetFocus
       Exit Sub
    End If
    
    If Text5.Text = "" Then
       MsgBox "la fecha no puede estar en blanco"
       Exit Sub
    End If
    
    If Text2.Text = "" Then
       MsgBox "el cliente no puede estar en blanco"
       Text2.SetFocus
       Exit Sub
    End If
    
    If Text14.Text = "" Then
       MsgBox "El vendedor no puede estar en blanco"
       Text14.SetFocus
       Exit Sub
    End If
    
    If ListView1.ListItems.Count = 0 Then
       MsgBox "Debe ingresar por lo menos un producto"
       Text8.SetFocus
       Exit Sub
    End If
 
    If IdPedido = "" Then
        If MsgBox("Desea almacenar el registro?", vbYesNo) = vbYes Then
            Set rs = cn.Execute("insert into Ordendepedido ( id_ordenpedido, fecha, id_cliente, id_vendedor ) values ('" & Trim(Me.Text1) & "', '" & Trim(Me.Text5) & "','" & Trim(Me.Text2) & "',  '" & Trim(Me.Text14) & "')")
            With ListView1.ListItems
                For i = 1 To .Count
                    Set rs = cn.Execute("INSERT INTO Detallepedido ( id_ordenpedido, Id_producto, peso) VALUES ('" + Text1.Text + "', '" + .Item(i).Text + "', '" + GetString(.Item(i).SubItems(2)) + "')")
                Next i
            End With
            IdPedido = ""
        Else
            MsgBox "Registro no almacenado", 64
        End If
    Else
        If MsgBox("Desea actualizar el registro?", vbYesNo) = vbYes Then
            Set rs = cn.Execute("UPDATE Ordendepedido SET fecha='" + Text5.Text + "',id_cliente='" + Text2.Text + "',id_vendedor='" + Text14.Text + "' WHERE id_ordenpedido LIKE '" + Text1.Text + "'")
            With ListView1.ListItems
                For i = 1 To .Count
                    If .Item(i).Tag = 0 Then ' no existia agregar
                        Set rs = cn.Execute("INSERT INTO Detallepedido VALUES ('" + IdPedido + "', '" + .Item(i).Text + "', '" + GetString(.Item(i).SubItems(2)) + "')")
                    End If
                    If .Item(i).Tag < 0 Then ' si existia y hay que actualizar
                        '
                        Set rs = cn.Execute("UPDATE Detallepedido SET Id_producto='" + .Item(i).Text + "', peso='" + GetString(.Item(i).SubItems(2)) + "' WHERE id_ordenpedido LIKE '" + IdPedido + "' AND Id_producto LIKE '" + .Item(i).Text + "'")
                    End If
                    ' no se hace nada ya que no se modifico de la lista
                Next i
            End With
        End If
    End If
    IdPedido = ""
    Me.Command6.Enabled = True
    limpiar
    bloquear
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    Dim UserDate As Date

    UserDate = CVDate(Text13)
    If frmCalendar.GetDate(UserDate) Then
        Text13 = UserDate
    End If
End Sub

Private Sub Command4_Click()
 Unload Me
End Sub

Private Sub Command5_Click()
Me.ListView1.ListItems.Clear
End Sub

Private Sub Command6_Click()
limpiar
End Sub

Private Sub Editar_Click()

    Text8.Text = ListView1.SelectedItem.Text
    Text9.Enabled = True
    Text9.Text = ListView1.SelectedItem.SubItems(2)
    IndexProducto = ListView1.SelectedItem.index
    Text8.SetFocus
End Sub

Private Sub Eliminar_Click()
    Elimina ListView1.SelectedItem.index
End Sub

Private Sub Form_Activate()
Set rs = New ADODB.Recordset
   rs.LockType = adLockBatchOptimistic
   rs.CursorType = adOpenKeyset
   rs.Source = ("select * from Ordendepedido where id_ordenpedido like  '" & Trim(Me.Text1) & "'")
   IndexProducto = 0
   rs.ActiveConnection = cn
   rs.Open
   h = 1
   c = c + Val(Me.Text1)
   h = h + 1
   Text1.SetFocus
End Sub

Private Sub Form_Load()
 Set cn = New ADODB.Connection
 ' cn.Open "DSN=MS Access Database;DBQ=pmm.mdb;DefaultDir=.\;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;"
 cn.Open "dsn=pmm"
 cn.CommandTimeout = 300
 IdPedido = ""
 Me.Text5 = Date
 bloquear
 Text13 = Date
 Me.Text1.Enabled = True
 Dim c, h
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If
End Sub

Private Sub Text1_LostFocus()

    If Trim(Me.Text1) = "" Then
        IdPedido = ""
        limpiar
        bloquear
    Else
        If IdPedido <> Text1.Text Then
            IdPedido = Text1.Text
            limpiar
            Text1.Text = IdPedido
            Set rs = New ADODB.Recordset
            rs.LockType = adLockBatchOptimistic
            rs.CursorType = adOpenKeyset
            rs.Source = ("SELECT * FROM Ordendepedido WHERE id_ordenpedido LIKE '" + IdPedido + "'")
            rs.ActiveConnection = cn
            rs.Open
            
            If rs.RecordCount > 0 Then ' encontrado
            
                Text5.Text = rs("fecha")
                Text2.Text = rs("id_cliente")
                Text14.Text = rs("id_vendedor")
                Text2_KeyPress (13)
                Text14_KeyPress (13)
                Dim rs1 As Recordset
                Set rs1 = New Recordset
                rs1.CursorLocation = adUseClient
                rs1.ActiveConnection = cn
                rs1.Open "SELECT id_producto,peso FROM Detallepedido WHERE id_ordenpedido=" + IdPedido
                Text8.Enabled = True
                For i = 1 To rs1.RecordCount
                    Text9.Enabled = True
                    Text8.Text = rs1("id_producto")
                    Text9.Text = rs1("peso")
                    Text9_KeyPress (13)
                    ListView1.ListItems(ListView1.ListItems.Count).Tag = rs1("id_producto")
                    rs1.MoveNext
                Next
            Else
                IdPedido = ""
            End If
        End If
    End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Trim(Me.Text14) = "" Then
   Me.Text2.SetFocus
  Else
   Set rs = New ADODB.Recordset
   rs.LockType = adLockBatchOptimistic
   rs.CursorType = adOpenKeyset
   rs.Source = ("select * from Vendedor where id_vendedor like '" & Trim(Me.Text14) & "'")
   rs.ActiveConnection = cn
   rs.Open
   If rs.RecordCount > 0 Then
    Me.Text11 = rs("nomvende")
   Else
    MsgBox "el vendedor no existe"
   End If
  End If
  Me.Command1.Enabled = True
  Me.Command1.SetFocus
 End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 Me.Text2.Enabled = True
 If KeyAscii = 13 Then
  If Trim(Me.Text2) = "" Then
  Else
   Set rs = New ADODB.Recordset
   rs.LockType = adLockBatchOptimistic
   rs.CursorType = adOpenKeyset
   rs.Source = ("select * from cliente where id_cliente like '" & Trim(Me.Text2) & "'")
   rs.ActiveConnection = cn
   rs.Open
   If rs.RecordCount > 0 Then
    Me.Text3 = rs("nomcliente")
    Me.Text4 = rs("ciudad")
    Me.Text6 = rs("direccliente")
    Me.Text7 = rs("telcliente")
   Else
    MsgBox "el cliente no existe"
    Me.Text2 = ""
    Me.Text3 = ""
    Me.Text4 = ""
    Me.Text6 = ""
    Me.Text7 = ""
   End If
  End If
  Me.Text8.Enabled = True
  Me.Text8.SetFocus
 End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Trim(Me.Text8) = "" Then
   Me.Text8.SetFocus
  Else
    If verifica(Text8.Text) = False Or IndexProducto <> 0 Then
  
        Set rs = New ADODB.Recordset
        rs.LockType = adLockBatchOptimistic
        rs.CursorType = adOpenKeyset
        rs.Source = ("select * from producto where Id_producto like '" & Trim(Me.Text8) & "'")
        rs.ActiveConnection = cn
        rs.Open
        If rs.RecordCount > 0 Then
         Me.Text9.Enabled = True
         Me.Text9.SetFocus
        Else
         MsgBox "el producto no existe"
         Me.Text9.Enabled = False
        End If
    Else
        MsgBox "Producto existente por favor editelo"
    End If
  End If
 End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
  
 If KeyAscii = 13 Then
  If Trim(Me.Text9) = "" Then
   Me.Text9.SetFocus
  Else
   Set rs = New ADODB.Recordset
   rs.LockType = adLockBatchOptimistic
   rs.CursorType = adOpenKeyset
   rs.Source = ("select * from producto where Id_producto like '" & Trim(Me.Text8) & "'")
   rs.ActiveConnection = cn
   rs.Open
   ' Me.Text9.SetFocus
   If rs.RecordCount > 0 Then
    
    If IndexProducto = 0 Then
        Dim itemx1 As ListItem
        ListView1.Enabled = True
        Set itemxi = Me.ListView1.ListItems.Add
        itemxi.Text = Me.Text8
        itemxi.SubItems(1) = rs("nomproduc")
        itemxi.SubItems(2) = Me.Text9
        itemxi.SubItems(3) = Val(rs("valor"))
        itemxi.SubItems(4) = GetNumero(rs("valor")) * GetNumero(Me.Text9.Text)
        itemxi.Tag = 0
    Else
        With ListView1.ListItems(IndexProducto)
        .Text = Text8.Text
        .SubItems(1) = rs("nomproduc")
        .SubItems(2) = Text9.Text
        .SubItems(3) = rs("valor")
        .SubItems(4) = GetNumero(rs("valor")) * GetNumero(Me.Text9.Text)
        ' se cambia el indice del producto a negativo para marcarlo como modificado
        .Tag = Abs(.Tag) * -1
        End With
    End If
    Me.Text8 = ""
    Me.Text9 = ""
    
    RecalculaTotal
    
    IndexProducto = 0
    Me.Text8.SetFocus
    Me.Text9.Enabled = False
    
   End If
   End If
 End If
Me.Text14.Enabled = True
End Sub

Public Sub limpiar()
 Me.Text1 = ""
 Me.Text2 = ""
 Me.Text3 = ""
 Me.Text4 = ""
 Me.Text5 = Date
 Me.Text6 = ""
 Me.Text7 = ""
 Me.Text8 = ""
 Me.Text9 = ""
 Me.Text11 = ""
 Me.Text12 = ""
 Me.Text13 = Date
 Me.Text14 = ""
 Combo1.Text = ""
 Me.ListView1.ListItems.Clear
 IndexProducto = 0
End Sub

Public Sub bloquear()
 Me.Text3.Enabled = False
 Me.Text4.Enabled = False
 Me.Text5.Enabled = False
 Me.Text6.Enabled = False
 Me.Text7.Enabled = False
 Me.Text8.Enabled = False
 Me.Text9.Enabled = False
 Me.Text11.Enabled = False
 Me.Text12.Enabled = False
 Me.Text13.Enabled = False
 Me.Text14.Enabled = False
 Me.Command1.Enabled = False
 Me.Command6.Enabled = False
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyDelete And ListView1.ListItems.Count > 0 Then
        Elimina ListView1.SelectedItem.index, True
    End If
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton And ListView1.ListItems.Count > 0 Then
        PopupMenu Lista
    End If
End Sub

Private Sub ListView1_DblClick()
    
    If ListView1.ListItems.Count > 0 Then
        Text8.Text = ListView1.SelectedItem.Text
        Text9.Enabled = True
        Text9.Text = ListView1.SelectedItem.SubItems(2)
        Text8.SetFocus
        IndexProducto = ListView1.SelectedItem.index
    End If
End Sub
