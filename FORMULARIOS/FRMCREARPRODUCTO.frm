VERSION 5.00
Begin VB.Form FRMCREARPRODUCTO 
   Caption         =   "Alta de productos"
   ClientHeight    =   5580
   ClientLeft      =   4215
   ClientTop       =   2805
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   11700
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   10080
      Picture         =   "FRMCREARPRODUCTO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   735
      Left            =   10080
      Picture         =   "FRMCREARPRODUCTO.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "ALTA DE PRODUCTOS"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9015
      Begin VB.PictureBox DTCProveedores 
         Height          =   315
         Left            =   6720
         ScaleHeight     =   255
         ScaleWidth      =   1755
         TabIndex        =   7
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtCantidad 
         Height          =   375
         Left            =   6720
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtIVA 
         Height          =   375
         Left            =   6720
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtVenta 
         Height          =   375
         Left            =   4440
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtCompra 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtProducto 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Marca"
         Height          =   255
         Left            =   6000
         TabIndex        =   16
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "NO. Producto"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Producto"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
         Height          =   375
         Index           =   1
         Left            =   5880
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Precio Compra"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Precio Venta"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "IVA"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   1
         Top             =   840
         Width           =   495
      End
   End
End
Attribute VB_Name = "FRMCREARPRODUCTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
If txtCodigo = "" Then MsgBox "El Campo NO. Producto no puede estar vacio", vbCritical, "ERROR": txtCodigo.SetFocus: Exit Sub
If txtProducto = "" Then MsgBox "El Campo Producto no puede estar vacio", vbCritical, "ERROR": txtProducto.SetFocus: Exit Sub
If txtCantidad = "" Then MsgBox "El Campo Cantidad no puede estar vacio", vbCritical, "ERROR": txtCantidad.SetFocus: Exit Sub
If txtCompra = "" Then MsgBox "El Campo Precio Compra no puede estar vacio", vbCritical, "ERROR": txtCompra.SetFocus: Exit Sub
If DTCProveedores = "" Then MsgBox "El Campo Marca no puede estar vacio", vbCritical, "ERROR": DTCProveedores.SetFocus: Exit Sub
PORCENT = InputBox("Especifique el porcentaje de margen de ganancia")
CALPORCENT = PORCENT / 100
PRECIOVENTA = txtCompra.Text / (1 - CALPORCENT)
txtVenta.Text = PRECIOVENTA
IVA = PRECIOVENTA * 0.16
txtIVA.Text = IVA
CODPRODUCTOS = txtCodigo.Text
With RsProductos
    .Requery
    .Find "ID_PRODUCTO='" & Trim(CODPRODUCTOS) & "'"
    If .EOF Then
        .AddNew
            !ID_PRODUCTO = CODPRODUCTOS
            !PRODUCTO = txtProducto.Text
            !CANTIDAD = txtCantidad.Text
            !PRECIO_COMPRA = txtCompra.Text
            !PRECIO_VENTA = txtVenta.Text
            !IVA = txtIVA.Text
            Dim CODPROVEEDORES1
            With RsProveedores
                .Requery
                .Find "NOMBRE='" & Trim(DTCProveedores.Text) & "'"
                CODPROVEEDORES1 = !IDPROVEEDOR
            End With
            !PROVEEDOR = CODPROVEEDORES1
        .UpdateBatch
        .Requery
    Else
        MsgBox "El codigo ya existe", vbCritical, "AVISO"
        .Requery
        Exit Sub
    End If
End With
MsgBox "Producto agregado exitosamente a la base de datos", vbInformation, "AVISO"
CODPRODUCTOS = 0
FRMPRODUCTOS.Show
Unload Me
End Sub

Private Sub cmdCancelar_Click()
LIMPIAR
MsgBox "Proceso cancelado", vbInformation, "AVISO"
FRMPRODUCTOS.Show
Unload Me

End Sub

Private Sub Form_Load()
PROVEEDORES
PRODUCTOS
Set DTCProveedores.RowSource = RsProveedores
DTCProveedores.BoundColumn = "NOMBRE"
DTCProveedores.ListField = "NOMBRE"
Set GrillaProductos.DataSource = RsProductos
txtVenta.Enabled = False
txtIVA.Enabled = False
BLOQUEAR_GRILLA
End Sub

Sub MAYUSCULAS()
    Dim I As Integer
    txtProducto.Text = UCase(txtProducto.Text)
    I = Len(txtProducto.Text)
    txtProducto.SelStart = I
End Sub

Private Sub txtProducto_Change()
MAYUSCULAS

End Sub

Sub BLOQUEAR_GRILLA()
    GrillaProductos.Columns(0).Locked = True
    GrillaProductos.Columns(1).Locked = True
    GrillaProductos.Columns(2).Locked = True
    GrillaProductos.Columns(3).Locked = True
    GrillaProductos.Columns(4).Locked = True
    GrillaProductos.Columns(5).Locked = True
End Sub

Sub LIMPIAR()
txtCodigo.Text = ""
txtProducto.Text = ""
txtCompra.Text = ""
txtVenta.Text = ""
txtIVA.Text = ""
txtCantidad.Text = ""
DTCProveedores.Text = ""
End Sub
