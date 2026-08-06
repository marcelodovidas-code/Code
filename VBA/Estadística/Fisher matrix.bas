Attribute VB_Name = "Fisher_Matrix"
Function Fishermatrix(datos)

'incluir títulos
nvar = datos.Columns.Count + 0
nobs = datos.Rows.Count + 0

ReDim matriz(1 To nvar, 1 To nvar)
For i = 1 To nvar
For j = 1 To nvar
    If i = j Then
    matriz(i, j) = datos(1, i)
    Else
    matriz(i, j) = ""
    End If
Next
Next

For i = 1 To nvar - 1
For j = (i + 1) To nvar

n00 = 0
n01 = 0
n10 = 0
n11 = 0

For k = 2 To nobs
    If (datos(k, i) = 0 And datos(k, j) = 0) Then
    n00 = n00 + 1
    ElseIf (datos(k, i) = 0 And datos(k, j)) = 1 Then
    n01 = n01 + 1
    ElseIf (datos(k, i) = 1 And datos(k, j) = 0) Then
    n10 = n10 + 1
    ElseIf (datos(k, i) = 1 And datos(k, j) = 1) Then
    n11 = n11 + 1
    End If
Next

n = n00 + n01 + n10 + n11
n0s = n00 + n01
ns0 = n00 + n10
n1s = n - n0s

pvalue0 = Application.WorksheetFunction.HypGeom_Dist(n00, ns0, n0s, n, False)
pvalue = Application.WorksheetFunction.HypGeom_Dist(n00, ns0, n0s, n, False)

For t = 0 To n0s
aux = Application.WorksheetFunction.HypGeom_Dist(t, ns0, n0s, n, False)
If aux < pvalue0 Then
pvalue = pvalue + aux
End If
Next

matriz(i, j) = pvalue
matriz(i, j) = Format(matriz(i, j), "0.000")

If pvalue < 0.05 Then
matriz(i, j) = matriz(i, j) & "*"
End If

Next
Next

Fishermatrix = matriz

End Function



