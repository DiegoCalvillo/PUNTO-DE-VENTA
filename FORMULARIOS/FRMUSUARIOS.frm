VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMUSUARIOS 
   Caption         =   "Usuarios del Sistema"
   ClientHeight    =   5595
   ClientLeft      =   4215
   ClientTop       =   3030
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   12165
   Begin MSDataGridLib.DataGrid Grillausuarios 
      Height          =   2655
      Left            =   360
      TabIndex        =   20
      Top             =   2760
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4683
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
      Left            =   9840
      Picture         =   "FRMUSUARIOS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   735
      Left            =   10920
      Picture         =   "FRMUSUARIOS.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   735
      Left            =   10920
      Picture         =   "FRMUSUARIOS.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   735
      Left            =   9840
      Picture         =   "FRMUSUARIOS.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   735
      Left            =   10920
      Picture         =   "FRMUSUARIOS.frx":2808
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   735
      Left            =   9840
      Picture         =   "FRMUSUARIOS.frx":320A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2880
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "DATOS DEL USUARIO "
      Height          =   2175
      Left            =   5400
      TabIndex        =   8
      Top             =   360
      Width           =   5535
      Begin VB.TextBox txtVentas 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtApellido 
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label7 
         Caption         =   "Ventas"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Apellido"
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CLAVES DE ACCESO AL SISTEMA"
      Height          =   1935
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4575
      Begin VB.TextBox txtContraseña 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtId 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "ID Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblCodigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo del Usuario:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FRMUSUARIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
FRMCREARUSUARIO.Show
Unload Me
End Sub

Private Sub cmdBuscar_Click()
FRMBUSCARUSUARIO.Show
Unload Me

End Sub

Private Sub cmdCancelar_Click()
HABILITARBOTONES True, False
HABILITARCAJAS True
MsgBox "Proceso cancelado", vbInformation, "AVISO"
CODUSUARIOS = 0
End Sub

Private Sub cmdEliminar_Click()
If CODUSUARIOS = 0 Then
    MsgBox "Elija un registro de la tabla", vbCritical, "ERROR"
    Exit Sub
Else
    With RsUsuarios
        .Find "CODIGO='" & Trim(CODUSUARIOS) & " '"
        If .EOF Then
            MsgBox "No se encontro ningun registro", vbInformation, "AVISO"
            Exit Sub
        Else
            If MsgBox("¿Desea eliminar este registro?" & Grillausuarios.Columns(5).Text, vbInformation + vbYesNo) = vbYes Then
                .Delete
                .Requery
                CODUSUARIOS = 0
            End If
        End If
    End With
End If
End Sub

Private Sub cmdGuardar_Click()
With RsUsuarios
    .Requery
    .Find "CODIGO='" & Trim(CODUSUARIOS) & " '"
        !LOGIN = txtId
        !CONTRASEÑA = txtContraseña
        !NOMBRE = txtNombre
        !APELLIDOS = txtApellido
    .Update
    .Requery
End With
HABILITARBOTONES True, False
HABILITARCAJAS True
CODUSUARIOS = 0
End Sub

Private Sub cmdModificar_Click()
If CODUSUARIOS = 0 Then
    MsgBox "Debe elegir un registro de la tabla", vbCritical, "ERROR"
    Exit Sub
Else
    HABILITARCAJAS False
    HABILITARBOTONES False, True
    MODI = True
End If
End Sub

Private Sub Form_Load()
USUARIOS
Set Grillausuarios.DataSource = RsUsuarios
HABILITARCAJAS True
HABILITARBOTONES True, False
ENCUENTRAUSUARIO
End Sub

Private Sub GrillaUsuarios_Click()
With RsUsuarios
    CODUSUARIOS = Grillausuarios.Columns(5).Text
    txtId.Text = Grillausuarios.Columns(0).Text
    txtContraseña.Text = Grillausuarios.Columns(1).Text
    txtNombre.Text = Grillausuarios.Columns(2).Text
    txtApellido.Text = Grillausuarios.Columns(3).Text
    txtVentas.Text = Grillausuarios.Columns(4).Text
End With
lblCodigo.Caption = CODUSUARIOS

End Sub

Public Sub HABILITARCAJAS(ESTADO As Boolean)
txtId.Locked = ESTADO
txtContraseña.Locked = ESTADO
txtNombre.Locked = ESTADO
txtApellido.Locked = ESTADO
txtVentas.Locked = ESTADO
End Sub

Public Sub HABILITARBOTONES(ESTADO1 As Boolean, ESTADO2 As Boolean)
cmdAgregar.Enabled = ESTADO1
cmdGuardar.Enabled = ESTADO2
cmdCancelar.Enabled = ESTADO2
cmdModificar.Enabled = ESTADO1
cmdBuscar.Enabled = ESTADO1
cmdEliminar.Enabled = ESTADO1

End Sub

Sub BLOQUEAR_GRILLA()
    Grillausuarios.Columns(5).Locked = True
    Grillausuarios.Columns(0).Locked = True
    Grillausuarios.Columns(1).Locked = True
    Grillausuarios.Columns(2).Locked = True
    Grillausuarios.Columns(3).Locked = True
    Grillausuarios.Columns(4).Locked = True
End Sub

Sub MAYUSCULAS()
Dim I As Integer
    txtId.Text = UCase(txtId.Text)
    I = Len(txtId.Text)
    txtId.SelStart = I
    
    txtContraseña.Text = UCase(txtContraseña.Text)
    I = Len(txtContraseña.Text)
    txtContraseña.SelStart = I
    
    txtNombre.Text = UCase(txtNombre.Text)
    I = Len(txtNombre.Text)
    txtNombre.SelStart = I
    
    txtApellido.Text = UCase(txtApellido.Text)
    I = Len(txtApellido.Text)
    txtApellido.SelStart = I
End Sub

Private Sub txtApellido_Change()
MAYUSCULAS

End Sub

Private Sub txtContraseña_Change()
MAYUSCULAS

End Sub

Private Sub txtId_Change()
MAYUSCULAS

End Sub

Private Sub txtNombre_Change()
MAYUSCULAS

End Sub

Sub ENCUENTRAUSUARIO()
With RsUsuarios
    .Requery
    .Find "CODIGO='" & Trim(CODUSUARIOS) & "'"
    If .EOF Then
        'MsgBox "No se encontro ningun registro", vbInformation, "AVISO"
        '.Requery
        Exit Sub
    Else
        lblCodigo.Caption = CODUSUARIOS
        txtId = !LOGIN
        txtContraseña = !CONTRASEÑA
        txtNombre = !NOMBRE
        txtApellido = !APELLIDOS
        BLOQUEAR_GRILLA
    End If
End With
BLOQUEAR_GRILLA
End Sub
