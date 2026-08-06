Attribute VB_Name = "Chi2 Matrix"
Function chi2matrix(datos)

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
n00esp = 0
n01esp = 0
n10esp = 0
n11esp = 0

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

'    If (n01 = 0 Or n10 = 0 Or n00 = 0 Or n11 = 0) Then
'        matriz(i, j) = (n00 + 0.5) * (n11 + 0.5) / (n01 + 0.5) / (n10 + 0.5)
'        Else
'        matriz(i, j) = n00 * n11 / n01 / n10
'    End If

n = n00 + n01 + n10 + n11
pri = (n10 + n11) / n
prj = (n01 + n11) / n

n00esp = n * (1 - pri) * (1 - prj)
n01esp = n * (1 - pri) * prj
n10esp = n * pri * (1 - prj)
n11esp = n * pri * prj

If (n00esp * n01esp * n10esp * n11esp * n00 * n01 * n01 * n11) = 0 Then
matriz(i, j) = "NA"
Else
matriz(i, j) = (n00 - n00esp) ^ 2 / n00esp + (n01 - n01esp) ^ 2 / n01esp + (n10 - n10esp) ^ 2 / n10esp + (n11 - n11esp) ^ 2 / n11esp
matriz(i, j) = Application.WorksheetFunction.ChiSq_Dist_RT(matriz(i, j), 1)
matriz(i, j) = Format(matriz(i, j), "0.00")
End If

Next
Next

'https://www.medcalc.org/calc/odds_ratio.php

chi2matrix = matriz

End Function

