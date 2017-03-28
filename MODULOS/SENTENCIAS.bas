Attribute VB_Name = "SENTENCIAS"
Sub main()
With BASE
    .CursorLocation = adUseClient
    .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PUNTODEVENTA.mdb;Persist Security Info=False"
    frmLogin.Show
End With
End Sub

Sub PRODUCTOS()
With RsProductos
    If .State = 1 Then .Close
    .Open "SELECT * FROM PRODUCTOS", BASE, adOpenStatic, adLockOptimistic
End With
End Sub

Sub USUARIOS()
With RsUsuarios
    If .State = 1 Then .Close
    .Open "SELECT * FROM USUARIOS", BASE, adOpenStatic, adLockOptimistic
    
End With
End Sub

Sub PROVEEDORES()
With RsProveedores
    If .State = 1 Then .Close
    .Open "SELECT * FROM PROVEEDORES", BASE, adOpenStatic, adLockOptimistic
End With
End Sub

Sub VENTADETALLE()
With RsVentaDetalle
    If .State = 1 Then .Close
    .Open "SELECT * FROM VENTA_DETALLE", BASE, adOpenStatic, adLockOptimistic
End With
End Sub

Sub DETALLETEMPORAL()
With RsDetalleTemporal
    If .State = 1 Then .Close
    .Open "SELECT * FROM VENTA_DETALLE_TEMPORAL", BASE, adOpenStatic, adLockOptimistic
End With
End Sub

Sub VENTAS()
With RsVentas
    If .State = 1 Then .Close
    .Open "SELECT * FROM VENTAS", BASE, adOpenStatic, adLockOptimistic
End With
End Sub
