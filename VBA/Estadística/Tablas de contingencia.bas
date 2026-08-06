Attribute VB_Name = "Tablas de Contingencia"
Function tabla2x2(datos)
'seleccionar los títulos

nvar = datos.Columns.Count + 0
nobs = datos.Rows.Count + 0

ReDim matriz(1 To nvar + 1, 1 To nvar + 1)
matriz(1, 1) = datos(1, 1) & "/" & datos(1, 2)
matriz(1, 2) = 0
matriz(1, 3) = 1
matriz(2, 1) = 0
matriz(3, 1) = 1

n00 = 0
n01 = 0
n10 = 0
n11 = 0

For k = 2 To nobs
    If (datos(k, 1) = 0 And datos(k, 2) = 0) Then
    n00 = n00 + 1
    ElseIf (datos(k, 1) = 0 And datos(k, 2)) = 1 Then
    n01 = n01 + 1
    ElseIf (datos(k, 1) = 1 And datos(k, 2) = 0) Then
    n10 = n10 + 1
    ElseIf (datos(k, 1) = 1 And datos(k, 2) = 1) Then
    n11 = n11 + 1
    End If
Next

matriz(2, 2) = n00
matriz(3, 2) = n10
matriz(2, 3) = n01
matriz(3, 3) = n11

tabla2x2 = matriz
End Function


Function tabla2x2bis(X1, X2)
'seleccionar los títulos

nvar = 2
nobs = X1.Rows.Count + 0

ReDim matriz(1 To nvar + 1, 1 To nvar + 1)
matriz(1, 1) = X1(1) & "/" & X2(1)
matriz(1, 2) = 0
matriz(1, 3) = 1
matriz(2, 1) = 0
matriz(3, 1) = 1

n00 = 0
n01 = 0
n10 = 0
n11 = 0

For k = 2 To nobs
    If (X1(k) = 0 And X2(k) = 0) Then
    n00 = n00 + 1
    ElseIf (X1(k) = 0 And X2(k) = 1) Then
    n01 = n01 + 1
    ElseIf (X1(k) = 1 And X2(k) = 0) Then
    n10 = n10 + 1
    ElseIf (X1(k) = 1 And X2(k) = 1) Then
    n11 = n11 + 1
    End If
Next

matriz(2, 2) = n00
matriz(3, 2) = n10
matriz(2, 3) = n01
matriz(3, 3) = n11

tabla2x2bis = matriz
End Function

