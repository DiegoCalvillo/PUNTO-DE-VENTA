VERSION 5.00
Begin VB.Form FRMPEDIDOS 
   Caption         =   "Levantar Pedidos"
   ClientHeight    =   6945
   ClientLeft      =   5505
   ClientTop       =   2595
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   9810
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8040
      ScaleHeight     =   315
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar pedido"
      Height          =   735
      Left            =   8640
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "DATOS DEL PRODUCTO"
      Height          =   2055
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   8055
      Begin VB.TextBox txtProducto 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtCantidad 
         Height          =   375
         Left            =   6120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtIVA 
         Height          =   375
         Left            =   6120
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtCompra 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtVenta 
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "NO. Producto"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Producto"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Precio Compra"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Precio Venta"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "IVA"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblCodigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECCIONE EL PROVEEDOR"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.PictureBox DTCProveedores 
         Height          =   315
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   5475
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Label lblCodigoProveedor 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6840
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "FRMPEDIDOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRegistrar_Click()
If CODPRODUCTOS = 0 Then
    MsgBox "No se la ha seleccionado ningun producto", vbInformation, "ERROR"
    Exit Sub
Else
    CANT = InputBox("Ingrese la cantidad pedida")
    txtCantidad.Text = CANT + Val(txtCantidad.Text)
    With RsProductos
        .Requery
        .Find "ID_PRODUCTO='" & Trim(CODPRODUCTOS) & " '"
            !CANTIDAD = txtCantidad.Text
        .UpdateBatch
        .Requery
    End With
End If
End Sub

Private Sub DTCProveedores_Change()
With RsProveedores
    .Requery
    .Find "NOMBRE='" & Trim(DTCProveedores.Text) & "'"
    lblCodigoProveedor.Caption = !IDPROVEEDOR
End With
Adodc1.RecordSource = "SELECT * FROM PRODUCTOS WHERE [PROVEEDOR] LIKE '" & lblCodigoProveedor.Caption & "'"
Adodc1.Refresh
End Sub

Private Sub Form_Load()
PRODUCTOS
PROVEEDORES
Set DTCProveedores.RowSource = RsProveedores
DTCProveedores.BoundColumn = "NOMBRE"
DTCProveedores.ListField = "NOMBRE"
Adodc1.CursorLocation = adUseClient
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PUNTODEVENTA.mdb;Persist Security Info=False"
Adodc1.RecordSource = "SELECT * FROM PRODUCTOS WHERE [PROVEEDOR] LIKE '" & lblCodigoProveedor.Caption & "'"
Adodc1.Refresh
Set GrillaProductos.DataSource = Adodc1
BLOQUEAR_GRILLA
HABILITARCAJAS True

End Sub

Private Sub GrillaProductos_Click()
With RsProductos
    If .BOF Or .EOF Then Exit Sub
    CODPRODUCTOS = GrillaProductos.Columns(0).Text
    lblCodigo.Caption = GrillaProductos.Columns(0).Text
    txtProducto.Text = GrillaProductos.Columns(1).Text
    txtCompra.Text = GrillaProductos.Columns(4).Text
    txtVenta.Text = GrillaProductos.Columns(3).Text
    txtCantidad.Text = GrillaProductos.Columns(2).Text
    txtIVA.Text = GrillaProductos.Columns(5).Text
End With
With RsProveedores
    .Requery
    .Find "IDPROVEEDOR='" & Trim(GrillaProductos.Columns(6).Text) & "'"
End With
End Sub

Sub BLOQUEAR_GRILLA()
    GrillaProductos.Columns(0).Locked = True
    GrillaProductos.Columns(1).Locked = True
    GrillaProductos.Columns(2).Locked = True
    GrillaProductos.Columns(3).Locked = True
    GrillaProductos.Columns(4).Locked = True
    GrillaProductos.Columns(5).Locked = True
    GrillaProductos.Columns(6).Locked = True
End Sub

Public Sub HABILITARCAJAS(ESTADO As Boolean)
txtProducto.Locked = ESTADO
txtCantidad.Locked = ESTADO
txtCompra.Locked = ESTADO
txtVenta.Locked = ESTADO
txtIVA.Locked = ESTADO

End Sub
