VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMBUSCARPRODUCTO 
   Caption         =   "Busqueda de Productos"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8160
      Top             =   1440
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
   Begin MSDataGridLib.DataGrid GrillaProductos 
      Height          =   2775
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   9375
      _ExtentX        =   16536
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
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      Height          =   735
      Left            =   8760
      Picture         =   "FRMBUSCARPRODUCTO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "Regresar"
      Height          =   735
      Left            =   7560
      Picture         =   "FRMBUSCARPRODUCTO.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "PARAMETRO DE BUSQUEDA"
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.TextBox txtParametroBusqueda 
         Height          =   405
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo Producto"
         Height          =   195
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OptNombre 
         Caption         =   "Nombre Producto"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FRMBUSCARPRODUCTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRegistrar_Click()
MsgBox (REGISTRAPRODUCTOS)
FRMSISTEMACOBRO.Show

End Sub

Private Sub cmdRegresar_Click()
MsgBox (CODPRODUCTOS)
FRMPRODUCTOS.Show
Unload Me

End Sub

Private Sub Form_Load()
PRODUCTOS
Adodc1.CursorLocation = adUseClient
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PUNTODEVENTA.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from PRODUCTOS"
Adodc1.Refresh
Set GrillaProductos.DataSource = Adodc1
BLOQUEAR_GRILLA

End Sub

Sub BLOQUEAR_GRILLA()
    GrillaProductos.Columns(0).Locked = True
    GrillaProductos.Columns(1).Locked = True
    GrillaProductos.Columns(2).Locked = True
    GrillaProductos.Columns(3).Locked = True
    GrillaProductos.Columns(4).Locked = True
    GrillaProductos.Columns(5).Locked = True
End Sub

Sub BUSCANOMBRE()
Dim BUSCA As String
BUSCA = UCase(Trim(txtParametroBusqueda.Text)) & "%"
Adodc1.RecordSource = "SELECT * FROM PRODUCTOS WHERE [PRODUCTO] LIKE '" & BUSCA & "'"
Adodc1.Refresh
End Sub

Private Sub GrillaProductos_Click()
With RsProductos
    CODPRODUCTOS = GrillaProductos.Columns(0).Text
    REGISTRAPRODUCTOS = GrillaProductos.Columns(0).Text
End With
End Sub

Private Sub txtParametroBusqueda_Change()
If OptNombre.Value = True Then BUSCANOMBRE
If OptCodigo.Value = True Then BUSCACODIGO
End Sub

Sub BUSCACODIGO()
Dim BUSCA As String
BUSCA = UCase(Trim(txtParametroBusqueda.Text)) & "%"
Adodc1.RecordSource = "SELECT * FROM PRODUCTOS WHERE [ID_PRODUCTO] LIKE '" & BUSCA & "'"
Adodc1.Refresh
End Sub

