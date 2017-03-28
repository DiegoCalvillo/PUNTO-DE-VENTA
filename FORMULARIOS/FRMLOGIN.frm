VERSION 5.00
Begin VB.Form FRMLOGIN 
   Caption         =   "INICIO DE SESION - PUNTO DE VENTA"
   ClientHeight    =   3930
   ClientLeft      =   6585
   ClientTop       =   3885
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   6645
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   735
      Left            =   3720
      Picture         =   "FRMLOGIN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdAccesar 
      Caption         =   "Accesar"
      Height          =   735
      Left            =   1680
      Picture         =   "FRMLOGIN.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "INTRODUZCA CLAVES DE ACCESO"
      Height          =   2055
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   5055
      Begin VB.TextBox txtContraseña 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtId 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
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
         Left            =   840
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FRMLOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAccesar_Click()
If txtId.Text = "" Then MsgBox "No ha ingresado la ID de Usuario", vbCritical, "ERROR": txtId.SetFocus: Exit Sub
If txtContraseña.Text = "" Then MsgBox "No ha ingresado la Contreseña de Usuario", vbCritical, "ERROR": txtContraseña.SetFocus: Exit Sub
With RsUsuarios
    .Requery
    .Find "LOGIN='" & Trim(txtId.Text) & "'"
    If .EOF Then
        MsgBox "ID de Usuario Incorrecto", vbCritical, "ERROR"
        txtId.Text = ""
        Exit Sub
    Else
        If !CONTRASEÑA = Trim(txtContraseña.Text) Then
            'FRMSUGERENCIA.Show
            FRMPRINCIPAL.Show
            FRMSISTEMACOBRO.Show
            Unload Me
        Else
            MsgBox "Contraseña incorrecta", vbCritical, "ERROR"
            txtContraseña.Text = ""
            Exit Sub
        End If
    End If
End With
End Sub

Private Sub cmdSalir_Click()
If MsgBox("¿Desea salir del sistema?", vbInformation + vbYesNo) = vbYes Then
Unload Me
End If
End Sub

Private Sub Form_Load()
USUARIOS

End Sub

Private Sub txtContraseña_Change()
Dim I As Integer
txtContraseña.Text = UCase(txtContraseña.Text)
I = Len(txtContraseña.Text)
txtContraseña.SelStart = I
End Sub

Private Sub txtId_Change()
Dim I As Integer
txtId.Text = UCase(txtId.Text)
I = Len(txtId.Text)
txtId.SelStart = I
End Sub
