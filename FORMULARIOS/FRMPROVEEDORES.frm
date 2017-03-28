VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMPROVEEDORES 
   Caption         =   "Proveedores"
   ClientHeight    =   5790
   ClientLeft      =   3780
   ClientTop       =   3030
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   13035
   Begin MSDataGridLib.DataGrid GrillaProveedores 
      Height          =   2535
      Left            =   360
      TabIndex        =   21
      Top             =   3120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4471
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
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   735
      Left            =   11760
      Picture         =   "FRMPROVEEDORES.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   10680
      Picture         =   "FRMPROVEEDORES.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   735
      Left            =   11760
      Picture         =   "FRMPROVEEDORES.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   735
      Left            =   10680
      Picture         =   "FRMPROVEEDORES.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   735
      Left            =   11760
      Picture         =   "FRMPROVEEDORES.frx":2808
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   735
      Left            =   10680
      Picture         =   "FRMPROVEEDORES.frx":320A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATOS DEL PROVEEDOR"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9975
      Begin VB.TextBox txtCodigoPostal 
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtTelefono 
         Height          =   375
         Left            =   6240
         TabIndex        =   14
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   6240
         TabIndex        =   12
         Top             =   1320
         Width           =   2895
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
         TabIndex        =   6
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Email"
         Height          =   375
         Left            =   5040
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Telefono"
         Height          =   375
         Left            =   5040
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Codigo Postal"
         Height          =   255
         Left            =   5040
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Direccion"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Contacto"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   975
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
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "FRMPROVEEDORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
FRMCREARPROVEEDOR.Show
Unload Me

End Sub

Private Sub cmdCancelar_Click()
HABILITARBOTONES True, False
HABILITARCAJAS True
MsgBox "Proceso cancelado", vbInformation, "AVISO"
CODPROVEEDORES = 0
End Sub

Private Sub cmdEliminar_Click()
'If CODPROVEEDORES = 0 Then
'    MsgBox "Elija un registro de la tabla", vbCritical, "ERROR"
'    Exit Sub
'Else
'    With RsProveedores
'        .Find "IDPROVEEDOR='" & Trim(CODPROVEEDORES) & " '"
'        If .EOF Then
'            MsgBox "No se encuentra ningun registro", vbInformation, "AVISO"
'            Exit Sub
'        Else
'            If MsgBox("¿Desea eliminar este registro?" & GrillaProveedores.Columns(0).Text, vbInformation + vbYesNo) = vbYes Then
'                .Delete
'                .Requery
'                CODPROVEEDORES = 0
'            End If
'        End If
'    End With
'End If
End Sub

Private Sub cmdGuardar_Click()
With RsProveedores
    .Requery
    .Find "IDPROVEEDOR='" & Trim(CODPROVEEDORES) & " '"
        !NOMBRE = txtNombre.Text
        !NOMBRE_CONTACTO = txtContacto.Text
        !DIRECCION = txtDireccion.Text
        !CODIGO_POSTAL = txtCodigoPostal.Text
        !TELEFONO = txtTelefono.Text
        !EMAIL = txtEmail.Text
    .Update
    .Requery
End With
HABILITARBOTONES True, False
HABILITARCAJAS True
CODPROVEEDORES = 0
End Sub

Private Sub cmdModificar_Click()
If CODPROVEEDORES = 0 Then
    MsgBox "No ha seleccionado ningun registro", vbCritical, "ERROR"
    Exit Sub
Else
    MODI = True
    HABILITARCAJAS False
    HABILITARBOTONES False, True
End If
End Sub

Private Sub Form_Load()
PROVEEDORES
Set GrillaProveedores.DataSource = RsProveedores
HABILITARCAJAS True
HABILITARBOTONES True, False
BLOQUEAR_GRILLA
End Sub

Public Sub HABILITARCAJAS(ESTADO As Boolean)
txtNombre.Locked = ESTADO
txtContacto.Locked = ESTADO
txtDireccion.Locked = ESTADO
txtCodigoPostal.Locked = ESTADO
txtTelefono.Locked = ESTADO
txtEmail.Locked = ESTADO
End Sub

Public Sub HABILITARBOTONES(ESTADO1 As Boolean, ESTADO2 As Boolean)
cmdAgregar.Enabled = ESTADO1
cmdGuardar.Enabled = ESTADO2
cmdCancelar.Enabled = ESTADO2
cmdModificar.Enabled = ESTADO1
cmdBuscar.Enabled = ESTADO1
cmdEliminar.Enabled = ESTADO1
End Sub

Private Sub GrillaProveedores_Click()
With RsProveedores
    CODPROVEEDORES = GrillaProveedores.Columns(0).Text
    txtNombre.Text = GrillaProveedores.Columns(1).Text
    txtContacto.Text = GrillaProveedores.Columns(2).Text
    txtDireccion.Text = GrillaProveedores.Columns(3).Text
    txtCodigoPostal.Text = GrillaProveedores.Columns(4).Text
    txtTelefono.Text = GrillaProveedores.Columns(5).Text
    txtEmail.Text = GrillaProveedores.Columns(6).Text
End With
lblCodigo.Caption = CODPROVEEDORES
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

