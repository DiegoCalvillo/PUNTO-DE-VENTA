VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMSISTEMACOBRO 
   Caption         =   "Sistema de Cobro - Punto de venta"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   18960
   Begin MSDataGridLib.DataGrid GrillaVentas 
      Height          =   6375
      Left            =   1920
      TabIndex        =   14
      Top             =   840
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   11245
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
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdCobrar 
      Caption         =   "Cobrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtLector 
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label lblRegistro 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11520
      TabIndex        =   13
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblCambio 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   12
      Top             =   9720
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "CAMBIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   11
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   12960
      TabIndex        =   10
      Top             =   9000
      Width           =   3255
   End
   Begin VB.Label lblIVA 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   9
      Top             =   8280
      Width           =   3255
   End
   Begin VB.Label lblSubtotal 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   8
      Top             =   7560
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "SUB-TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   7
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      TabIndex        =   6
      Top             =   8160
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "TOTAL A PAGAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   5
      Top             =   8880
      Width           =   3135
   End
End
Attribute VB_Name = "FRMSISTEMACOBRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SUBTOTAL As Integer
Dim TOTAL As Double
Dim IVA As Double
Dim CANTIDAD1 As Integer
Dim PRODUCTO1 As String
Dim IDPRODUCTO As Double
Dim PRECIO As Double
Dim TOTALIVA As Double
Dim CAMBIO As Double
Dim EFECTIVO As Double

'Variables que estaran en el boton BORRAR PRODUCTO
Dim CODIPRODU As Double
Dim AUXSUBTOTAL As Double
Dim AUXIVA As Double
Dim AUXTOTAL As Double

Sub LIMPIAR()
txtLector.Text = ""
End Sub

Private Sub cmdBorrar_Click()
If CODIPRODU = 0 Then
    MsgBox "Debe elegir un producto", vbInformation, "AVISO"
    Exit Sub
Else
    With RsDetalleTemporal
        .Delete
        .Requery
        CODIPRODU = 0
    End With
End If
FORMATOGRILLA

'CALCULOS PARA RECALCULAR AL MOMENTO DE BORRAR UN PRODUCTO DEL SISTEMA DE COBRO
AUXIVA = AUXSUBTOTAL * 0.16
SUBTOTAL = Val(lblSubtotal.Caption) - AUXSUBTOTAL
lblSubtotal.Caption = SUBTOTAL
TOTALIVA = TOTALIVA - AUXIVA
lblIVA.Caption = TOTALIVA
AUXTOTAL = AUXIVA + AUXSUBTOTAL
TOTAL = TOTAL - AUXTOTAL
lblTotal.Caption = TOTAL

End Sub

Private Sub cmdBuscar_Click()
FRMBUSCARPRODUCTO.Show

End Sub

Private Sub cmdCobrar_Click()
'VERIFICA QUE SE HAYAN AGREGADO PRODUCTOS AL SISTEMA DE COBRO
With RsDetalleTemporal
    If .EOF Or .BOF Then
        MsgBox "No se ha ingresado ningun producto", vbInformation, "AVISO"
        FORMATOGRILLA
        Exit Sub
    Else
       EFECTIVO = InputBox("Ingrese el importe")
        If EFECTIVO < SUBTOTAL Then
            MsgBox "El IMPORTE ingresado es menor que el TOTAL, por favor verifique", vbCritical, "ERROR"
            Exit Sub
        Else
            CAMBIO = EFECTIVO - TOTAL
            lblCambio.Caption = CAMBIO
        End If
    End If
End With
'GUARDA LA VENTA EN LA TABLA VENTA DE LA BASE DE DATOS
'AGREGAR EN LA TABLA VENTA CADA UNO DE LOS PRODUCTOS VENDIDOS
'BORRA LA TABLA E INICIZALIZA TODAS LAS VARIABLES
BORRARTEMPORAL
SUBTOTAL = 0
TOTALIVA = 0
TOTAL = 0
CAMBIO = 0
End Sub

Private Sub Form_Load()
VENTADETALLE
PRODUCTOS
DETALLETEMPORAL
VENTAS
Set GrillaVentas.DataSource = RsDetalleTemporal
REGISTRAPRODUCTO
lblRegistro.Caption = REGISTRAPRODUCTOS
FORMATOGRILLA
End Sub

Sub FORMATOGRILLA()
With RsDetalleTemporal
    GrillaVentas.Columns(0).Width = 0
    GrillaVentas.Columns(1).Width = 0
    GrillaVentas.Columns(2).Width = 1000
    GrillaVentas.Columns(3).Width = 1600
    GrillaVentas.Columns(4).Width = 10000
    GrillaVentas.Columns(5).Width = 2000
End With
End Sub



Private Sub GrillaVentas_Click()
With RsDetalleTemporal
    CANTIDAD1 = GrillaVentas.Columns(2).Text
    CODIPRODU = GrillaVentas.Columns(3).Text
    AUXSUBTOTAL = GrillaVentas.Columns(5).Text
End With
End Sub

Sub ENCUENTRAPRODUCTO()
'Dim PRODUCTO1 As String
'Dim IDPRODUCTO As Double
'Dim PRECIO As Double
If txtLector.Text <> "" Then
BUSCAPRODUCTOS = Trim(txtLector.Text)
With RsProductos
    .Requery
    .Find "ID_PRODUCTO='" & Trim(BUSCAPRODUCTOS) & "'"
    If .EOF Then
        MsgBox "Producto no encontrado", vbCritical, "AVISO"
        .Requery
        Exit Sub
    Else
        PRODUCTO1 = !PRODUCTO
        IDPRODUCTO = !ID_PRODUCTO
        CANTIDAD1 = 1
        PRECIO = !PRECIO_VENTA
        'CALCULANDO EL IVA DEL PRECIO UNITARIO DE CADA PRODUCTO
        IVA = PRECIO * 0.16
        TOTALIVA = TOTALIVA + IVA
        lblIVA.Caption = TOTALIVA
        'CALCULO DE SUBTOTALES Y TOTALES
        SUBTOTAL = SUBTOTAL + (CANTIDAD1 * PRECIO)
        lblSubtotal.Caption = SUBTOTAL
        TOTAL = SUBTOTAL + TOTALIVA
        lblTotal.Caption = TOTAL
    End If
End With
FORMATOGRILLA
With RsDetalleTemporal
    .Requery
    .AddNew
        !ID_PRODUCTO = Val(IDPRODUCTO)
        !PRODUCTO = PRODUCTO1
        !PRECIO_UNITARIO = CDbl(PRECIO)
        !CANTIDAD = CANTIDAD1
    .Update
End With
End If
LIMPIAR
FORMATOGRILLA
End Sub

Sub BORRARTEMPORAL()
With RsDetalleTemporal
    .Requery
    If .EOF Or .BOF Then Exit Sub
    For X = 1 To .RecordCount
        .Delete
        If .EOF Or .BOF Then Exit Sub
            .MoveNext
    Next
End With
lblSubtotal.Caption = ""
lblIVA.Caption = ""
lblTotal.Caption = ""
FORMATOGRILLA
End Sub

Sub REGISTRAPRODUCTO()
With RsProductos
    .Requery
    .Find "ID_PRODUCTO='" & Trim(REGISTRAPRODUCTOS) & "'"
    If .EOF Then
        'MsgBox "Producto no encontrado", vbCritical, "AVISO"
        '.Requery
        Exit Sub
    Else
        PRODUCTO1 = !PRODUCTO
        IDPRODUCTO = !ID_PRODUCTO
        CANTIDAD1 = 1
        PRECIO = !PRECIO_VENTA
        'CALCULANDO EL IVA DEL PRECIO UNITARIO DE CADA PRODUCTO
        IVA = PRECIO * 0.16
        TOTALIVA = TOTALIVA + IVA
        lblIVA.Caption = TOTALIVA
        'CALCULO DE SUBTOTALES Y TOTALES
        SUBTOTAL = SUBTOTAL + (CANTIDAD1 * PRECIO)
        lblSubtotal.Caption = SUBTOTAL
        TOTAL = SUBTOTAL + TOTALIVA
        lblTotal.Caption = TOTAL
    End If
End With
FORMATOGRILLA
With RsDetalleTemporal
    .Requery
    .AddNew
        !ID_PRODUCTO = Val(IDPRODUCTO)
        !PRODUCTO = PRODUCTO1
        !PRECIO_UNITARIO = CDbl(PRECIO)
        !CANTIDAD = CANTIDAD1
    .Update
End With
FORMATOGRILLA
End Sub

Private Sub txtLector_Change()
ENCUENTRAPRODUCTO
End Sub


