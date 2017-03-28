VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMCREARUSUARIO 
   Caption         =   "Alta de Usuario"
   ClientHeight    =   5505
   ClientLeft      =   4440
   ClientTop       =   2805
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   11820
   Begin MSDataGridLib.DataGrid Grillausuarios 
      Height          =   2295
      Left            =   600
      TabIndex        =   16
      Top             =   3000
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4048
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
      Left            =   10080
      Picture         =   "FRMCREARUSUARIO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear"
      Height          =   735
      Left            =   10080
      Picture         =   "FRMCREARUSUARIO.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "ALTA DE DATOS DE USUARIO"
      Height          =   2055
      Left            =   5760
      TabIndex        =   7
      Top             =   360
      Width           =   5775
      Begin VB.TextBox txtApellido 
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   "Apellido"
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ALTA DE CLAVES DE ACCESO"
      Height          =   2415
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4935
      Begin VB.TextBox txtConfirmContraseña 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtContraseña 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtId 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Confirmar contraseña"
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
         Left            =   360
         TabIndex        =   15
         Top             =   1680
         Width           =   1575
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
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
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
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblCodigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo de Usuario"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FRMCREARUSUARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
LIMPIAR
MsgBox "Proceso cancelado", vbInformation, "AVISO"
FRMUSUARIOS.Show
Unload Me
End Sub

Private Sub cmdCrear_Click()
If txtId = "" Then MsgBox "El campo ID Usuario no puede estar vacio", vbCritical, "ERROR": txtId.SetFocus: Exit Sub
If txtContraseña = "" Then MsgBox "El campo Contraseña no puede estar vacio", vbCritical, "ERROR": txtContraseña.SetFocus: Exit Sub
If txtNombre = "" Then MsgBox "El campo Nombre no puede estar vacio", vbCritical, "ERROR": txtNombre.SetFocus: Exit Sub
If txtApellido = "" Then MsgBox "El campo Apellido no puede estar vacio", vbCritical, "ERROR": txtApellido.SetFocus: Exit Sub
If txtConfirmContraseña = "" Then MsgBox "Es necesario que se confirme la contraseña", vbCritical, "ERROR": txtConfirmContraseña.SetFocus: Exit Sub
If txtContraseña.Text = txtConfirmContraseña.Text Then
    With RsUsuarios
        .Requery
        .AddNew
            !LOGIN = txtId.Text
            !CONTRASEÑA = txtContraseña.Text
            !NOMBRE = txtNombre.Text
            !APELLIDOS = txtApellido.Text
        .Update
        .Requery
    End With
Else
    MsgBox "La confirmacion de la contraseña no coincide. Por favor verifique", vbCritical, "ERROR"
    txtConfirmContraseña.SetFocus
    Exit Sub
End If
MsgBox "El USUARIO fue agregado exitosamente a la base de datos", vbInformation, "AVISO"
FRMUSUARIOS.Show
Unload Me

End Sub

Private Sub Form_Load()
USUARIOS
Set Grillausuarios.DataSource = RsUsuarios
BLOQUEAR_GRILLA
End Sub

Sub BLOQUEAR_GRILLA()
    Grillausuarios.Columns(5).Width = 0
    Grillausuarios.Columns(0).Locked = True
    Grillausuarios.Columns(1).Locked = True
    Grillausuarios.Columns(2).Locked = True
    Grillausuarios.Columns(3).Locked = True
    Grillausuarios.Columns(4).Locked = True
End Sub

Sub LIMPIAR()
txtId.Text = ""
txtContraseña.Text = ""
txtNombre.Text = ""
txtApellido.Text = ""
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

Private Sub txtNombre_Change()
MAYUSCULAS

End Sub
