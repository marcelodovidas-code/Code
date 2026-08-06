Attribute VB_Name = "Módulo1"
Function datosR(datos)
'pintar el rango de datos, incluído el título

n = Application.WorksheetFunction.Count(datos)

aux = datos(1) & "<-c("

For i = 2 To n - 1
aux = aux & datos(i) & ","
Next


datosR = aux & datos(n) & ")"

End Function
