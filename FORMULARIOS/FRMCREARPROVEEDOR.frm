VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMCREARPROVEEDOR 
   Caption         =   "Alta de Proveedores"
   ClientHeight    =   5940
   ClientLeft      =   3780
   ClientTop       =   3030
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   13035
   Begin MSDataGridLib.DataGrid GrillaProveedores 
      Height          =   2775
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   9975
      _ExtentX        =   17595
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
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   10680
      Picture         =   "FRMCREARPROVEEDOR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   735
      Left            =   10680
      Picture         =   "FRMCREARPROVEEDOR.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "ALTA DE PROVEEDORES"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9975
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   6120
         TabIndex        =   14
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtTelefono 
         Height          =   375
         Left            =   6120
         TabIndex        =   13
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtCodigoPostal 
         Height          =   375
         Left            =   6120
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtDireccion 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   1800
         Width           =   5295
      End
      Begin VB.TextBox txtContacto 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Email"
         Height          =   375
         Left            =   4800
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Telefono"
         Height          =   255
         Left            =   4800
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Codigo Postal"
         Height          =   375
         Left            =   4800
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Direccion"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Contacto"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblCodigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FRMCREARPROVEEDOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub LIMPIAR()
txtNombre.Text = ""
txtDireccion.Text = ""
txtContacto.Text = ""
txtCodigoPostal.Text = ""
txtTelefono.Text = ""
txtEmail.Text = ""
End Sub

Private Sub cmdAgregar_Click()
If txtNombre = "" Then MsgBox "El campo Nombre no puede estar vacio", vbCritical, "ERROR": txtNombre.SetFocus: Exit Sub
If txtContacto = "" Then MsgBox "El campo Contacto no puede estar vacio", vbCritical, "ERROR": txtContacto.SetFocus: Exit Sub
If txtDireccion = "" Then MsgBox "El campo Direccion no puede estar vacio", vbCritical, "ERROR": txtDireccion.SetFocus: Exit Sub
If txtTelefono = "" Then MsgBox "El campo Telefono no puede estar vacio", vbCritical, "ERROR": txtTelefono.SetFocus: Exit Sub
With RsProveedores
    .Requery
    .AddNew
        !NOMBRE = txtNombre.Text
        !NOMBRE_CONTACTO = txtContacto.Text
        !DIRECCION = txtDireccion.Text
        !CODIGO_POSTAL = txtCodigoPostal.Text
        !TELEFONO = txtTelefono.Text
        !EMAIL = txtEmail.Text
    .Update
    .Requery
End With
MsgBox "Proveedor registrado en el sistema", vbInformation, "AVISO"
FRMPROVEEDORES.Show
Unload Me
End Sub

Private Sub cmdCancelar_Click()
LIMPIAR
MsgBox "Proceso cancelado", vbInformation, "AVISO"
FRMPROVEEDORES.Show
Unload Me
End Sub

Private Sub Form_Load()
PROVEEDORES
Set GrillaProveedores.DataSource = RsProveedores
BLOQUEAR_GRILLA
End Sub

Sub MAYUSCULAS()
    Dim I As Integer
    txtNombre.Text = UCase(txtNombre.Text)
    I = Len(txtNombre.Text)
    txtNombre.SelStart = I
    
    txtContacto.Text = UCase(txtContacto.Text)
    I = Len(txtContacto.Text)
    txtContacto.SelStart = I
    
    txtDireccion.Text = UCase(txtDireccion.Text)
    I = Len(txtDireccion.Text)
    txtDireccion.SelStart = I
    
    txtEmail.Text = UCase(txtEmail.Text)
    I = Len(txtEmail.Text)
    txtEmail.SelStart = I
End Sub

Private Sub txtContacto_Change()
MAYUSCULAS

End Sub

Private Sub txtDireccion_Change()
MAYUSCULAS

End Sub

Private Sub txtEmail_Change()
MAYUSCULAS

End Sub

Private Sub txtNombre_Change()
MAYUSCULAS

End Sub

Sub BLOQUEAR_GRILLA()
    GrillaProveedores.Columns(0).Locked = True
    GrillaProveedores.Columns(1).Locked = True
    GrillaProveedores.Columns(2).Locked = True
    GrillaProveedores.Columns(3).Locked = True
    GrillaProveedores.Columns(4).Locked = True
    GrillaProveedores.Columns(5).Locked = True
    GrillaProveedores.Columns(6).Locked = True
End Sub
