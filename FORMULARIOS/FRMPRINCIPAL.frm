VERSION 5.00
Begin VB.MDIForm FRMPRINCIPAL 
   BackColor       =   &H8000000C&
   Caption         =   "PAGINA PRINCIPAL - PUNTO DE VENTA"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuUsuarios 
         Caption         =   "Usuarios"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuProductos 
         Caption         =   "Productos"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Cerrar Sesion"
      End
   End
   Begin VB.Menu mnuProveedores 
      Caption         =   "Proveedores"
      Begin VB.Menu mnuVerProveedores 
         Caption         =   "Ver Proveedores"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuRegistrar 
         Caption         =   "Registrar Pedido"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuSistema 
      Caption         =   "Sistema de Cobro"
   End
End
Attribute VB_Name = "FRMPRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuProductos_Click()
FRMPRODUCTOS.Show

End Sub

Private Sub mnuRegistrar_Click()
FRMPEDIDOS.Show

End Sub

Private Sub mnuSalir_Click()
If MsgBox("¿Desea cerrar sesion y salir del sistema?", vbInformation + vbYesNo) = vbYes Then
    Unload Me
    FRMLOGIN.Show
End If
End Sub

Private Sub mnuSistema_Click()
FRMSISTEMACOBRO.Show

End Sub

Private Sub mnuUsuarios_Click()
FRMUSUARIOS.Show

End Sub

Private Sub mnuVerProveedores_Click()
FRMPROVEEDORES.Show

End Sub
