VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMPRODUCTOS 
   Caption         =   "Cambios, Altas y Bajas de Productos"
   ClientHeight    =   5640
   ClientLeft      =   4215
   ClientTop       =   3030
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   12285
   Begin MSDataGridLib.DataGrid GrillaProductos 
      Height          =   2775
      Left            =   240
      TabIndex        =   21
      Top             =   2760
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4895
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   735
      Left            =   9720
      Picture         =   "FRMPRODUCTOS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "INFORMACION DEL PRODUCTO SELECCIONADO"
      Height          =   2295
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   9015
      Begin VB.PictureBox DTCProveedores 
         Height          =   315
         Left            =   6840
         ScaleHeight     =   255
         ScaleWidth      =   1755
         TabIndex        =   0
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtVenta 
         Height          =   375
         Left            =   4560
         TabIndex        =   19
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtCompra 
         Height          =   375
         Left            =   2040
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtIVA 
         Height          =   375
         Left            =   6840
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtCantidad 
         Height          =   375
         Left            =   6840
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtProducto 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Marca"
         Height          =   255
         Left            =   6120
         TabIndex        =   20
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblCodigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "IVA"
         Height          =   375
         Index           =   1
         Left            =   6240
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Precio Venta"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Precio Compra"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
         Height          =   375
         Index           =   1
         Left            =   6000
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Producto"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "NO. Producto"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   735
      Left            =   10800
      Picture         =   "FRMPRODUCTOS.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   9720
      Picture         =   "FRMPRODUCTOS.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   735
      Left            =   10800
      Picture         =   "FRMPRODUCTOS.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   735
      Left            =   10800
      Picture         =   "FRMPRODUCTOS.frx":2808
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   735
      Left            =   9720
      Picture         =   "FRMPRODUCTOS.frx":320A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
End
Attribute VB_Name = "FRMPRODUCTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
FRMCREARPRODUCTO.Show
Unload Me
End Sub

Private Sub cmdBuscar_Click()
FRMBUSCARPRODUCTO.Show
Unload Me

End Sub

Private Sub cmdCancelar_Click()
HABILITARBOTONES True, False
HABILITARCAJAS True
MsgBox "Proceso cancelado", vbInformation, "AVISO"
CODPRODUCTOS = 0
End Sub

Private Sub cmdEliminar_Click()
If CODPRODUCTOS = 0 Then
    MsgBox "Elija un registro de la tabla", vbCritical, "ERROR"
    Exit Sub
Else
    With RsProductos
        .Find "ID_PRODUCTO='" & Trim(CODPRODUCTOS) & " '"
        If .EOF Then
            MsgBox "No se encontro ningun registro", vbInformation, "AVISO"
            Exit Sub
        Else
            If MsgBox("¿Desea eliminar este registro?" & GrillaProductos.Columns(0).Text, vbInformation + vbYesNo) = vbYes Then
                .Delete
                .Requery
                CODPRODUCTOS = 0
            End If
        End If
    End With
End If
End Sub

Private Sub cmdGuardar_Click()
With RsProductos
    .Requery
    .Find "ID_PRODUCTO='" & Trim(CODPRODUCTOS) & " '"
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
End With
CODPRODUCTOS = 0
HABILITARBOTONES True, False
HABILITARCAJAS True

End Sub

Private Sub cmdModificar_Click()
If CODPRODUCTOS = 0 Then
    MsgBox "No ha seleccionado ningun registro", vbCritical, "ERROR"
    Exit Sub
Else
    MODI = True
    HABILITARCAJAS False
    HABILITARBOTONES False, True
    If MsgBox("¿Desea corregir el margen de ganacia?", vbInformation + vbYesNo) = vbYes Then
        PORCENT = InputBox("Especifique el porcentaje de margen de ganancia")
        CALPORCENT = PORCENT / 100
        PRECIOVENTA = txtCompra.Text / (1 - CALPORCENT)
        txtVenta.Text = PRECIOVENTA
        IVA = PRECIOVENTA * 0.16
        txtIVA.Text = IVA
    End If
End If
End Sub

Private Sub Form_Load()
PRODUCTOS
PROVEEDORES
Set DTCProveedores.RowSource = RsProveedores
DTCProveedores.BoundColumn = "NOMBRE"
DTCProveedores.ListField = "NOMBRE"
Set GrillaProductos.DataSource = RsProductos
BLOQUEAR_GRILLA
HABILITARCAJAS True
HABILITARBOTONES True, False
txtVenta.Enabled = False
txtIVA.Enabled = False
FORMATOGRILLA
ENCUENTRA
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

Private Sub GrillaProductos_Click()
With RsProductos
    CODPRODUCTOS = GrillaProductos.Columns(0).Text
    txtProducto.Text = GrillaProductos.Columns(1).Text
    txtCantidad.Text = GrillaProductos.Columns(2).Text
    txtCompra.Text = GrillaProductos.Columns(4).Text
    txtVenta.Text = GrillaProductos.Columns(3).Text
    txtIVA.Text = GrillaProductos.Columns(5).Text
With RsProveedores
    If .BOF Or .EOF Then Exit Sub
        .Requery
        .Find "IDPROVEEDOR='" & Trim(GrillaProductos.Columns(6)) & "'"
        DTCProveedores.Text = !NOMBRE
End With
End With
lblCodigo.Caption = CODPRODUCTOS
End Sub

Public Sub HABILITARCAJAS(ESTADO As Boolean)
txtProducto.Locked = ESTADO
txtCantidad.Locked = ESTADO
txtCompra.Locked = ESTADO
DTCProveedores.Locked = ESTADO
End Sub

Public Sub HABILITARBOTONES(ESTADO1 As Boolean, ESTADO2 As Boolean)
cmdAgregar.Enabled = ESTADO1
cmdGuardar.Enabled = ESTADO2
cmdCancelar.Enabled = ESTADO2
cmdModificar.Enabled = ESTADO1
cmdBuscar.Enabled = ESTADO1
cmdEliminar.Enabled = ESTADO1
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

Sub ENCUENTRA()
With RsProductos
        .Requery
        .Find "ID_PRODUCTO='" & Trim(CODPRODUCTOS) & "'"
        If .EOF Then
            'MsgBox "No se encontro ningun regisgtro", vbInformation, "AVISO"
            '.Requery
            BLOQUEAR_GRILLA
            Exit Sub
        Else
            lblCodigo.Caption = CODPRODUCTOS
            txtProducto.Text = !PRODUCTO
            txtCantidad.Text = !CANTIDAD
            txtVenta.Text = !PRECIO_VENTA
            txtIVA.Text = !IVA
            txtCompra.Text = !PRECIO_COMPRA
            Dim MARCA
            MARCA = !PROVEEDOR
            Dim PROVEEDOR123
            With RsProveedores
                .Requery
                .Find "IDPROVEEDOR='" & Trim(MARCA) & "'"
                PROVEEDOR123 = !NOMBRE
            End With
            DTCProveedores = PROVEEDOR123
        End If
    BLOQUEAR_GRILLA
    End With
End Sub

Sub FORMATOGRILLA()
With RsProductos
    GrillaProductos.Columns(6).Width = 0
    GrillaProductos.Columns(1).Width = 4000
    
End With
End Sub
