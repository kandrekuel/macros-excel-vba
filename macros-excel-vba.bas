Attribute VB_Name = "Module6"
Sub AgregarProducto()
    ' Declarar variables
    Dim producto As String
    Dim cantidad As Integer
    Dim precio As Currency
    Dim total As Currency
    Dim ultimaFila As Long

    ' Solicitar datos al usuario
    producto = InputBox("Ingrese el nombre del producto:")
    cantidad = InputBox("Ingrese la cantidad del producto:")
    precio = InputBox("Ingrese el precio del producto:")

    ' Calcular el total
    total = cantidad * precio

    ' Determinar la œltima fila con datos en la columna A
    ultimaFila = Range("A" & Rows.Count).End(xlUp).Row

    ' Agregar los datos a la siguiente fila disponible
    Cells(ultimaFila + 1, 1).Value = producto
    Cells(ultimaFila + 1, 2).Value = cantidad
    Cells(ultimaFila + 1, 3).Value = precio
    Cells(ultimaFila + 1, 4).Value = total

    ' Aplicar el formato de la tabla a la nueva fila
    With Range("A" & ultimaFila + 1 & ":D" & ultimaFila + 1)
        .Borders.Weight = xlThin
        .Interior.Color = RGB(220, 230, 241) ' Color de fondo claro
        .HorizontalAlignment = xlCenter
    End With
End Sub

