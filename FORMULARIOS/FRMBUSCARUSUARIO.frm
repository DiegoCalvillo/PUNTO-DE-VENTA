VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMBUSCARUSUARIO 
   Caption         =   "Busqueda de Usuario"
   ClientHeight    =   4875
   ClientLeft      =   5940
   ClientTop       =   3240
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   9900
   Begin MSDataGridLib.DataGrid Grillausuarios 
      Height          =   2655
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   8775
      _ExtentX        =   15478
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7440
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "Regresar"
      Height          =   735
      Left            =   7320
      Picture         =   "FRMBUSCARUSUARIO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "PARAMETRO DE BUSQUEDA"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      Begin VB.OptionButton OptApellido 
         Caption         =   "Apellido"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtParametroBusqueda 
         Height          =   405
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   4215
      End
   End
End
Attribute VB_Name = "FRMBUSCARUSUARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRegresar_Click()
MsgBox (CODUSUARIOS)
FRMUSUARIOS.Show
Unload Me

End Sub

Private Sub Form_Load()
USUARIOS
Adodc1.CursorLocation = adUseClient
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PUNTODEVENTA.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from USUARIOS"
Adodc1.Refresh
Set Grillausuarios.DataSource = Adodc1
BLOQUEAR_GRILLA

End Sub

Sub BUSCAUSUARIO()
Dim BUSCA As String
BUSCA = UCase(Trim(txtParametroBusqueda.Text)) & "%"
Adodc1.RecordSource = "select * from USUARIOS where [NOMBRE] like '" & BUSCA & "'"
Adodc1.Refresh
End Sub

Private Sub GrillaUsuarios_Click()
With RsUsuarios
    CODUSUARIOS = Grillausuarios.Columns(5).Text
End With
End Sub

Private Sub txtParametroBusqueda_Change()
If OptNombre.Value = True Then BUSCAUSUARIO
If OptApellido.Value = True Then BUSCAAPELLIDO
End Sub

Sub BUSCAAPELLIDO()
Dim BUSCA As String
BUSCA = UCase(Trim(txtParametroBusqueda.Text)) & "%"
Adodc1.RecordSource = "select * from USUARIOS where [APELLIDOS] like '" & BUSCA & "'"
Adodc1.Refresh
End Sub

Sub BLOQUEAR_GRILLA()
    Grillausuarios.Columns(5).Width = 0
    Grillausuarios.Columns(0).Locked = True
    Grillausuarios.Columns(1).Locked = True
    Grillausuarios.Columns(2).Locked = True
    Grillausuarios.Columns(3).Locked = True
    Grillausuarios.Columns(4).Locked = True
End Sub
