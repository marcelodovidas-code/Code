Attribute VB_Name = "OddsRatio_Matrix"

Function OddsRatioMatrix(datos)
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

    If (n01 = 0 Or n10 = 0) Then
        matriz(i, j) = "Inf"
        Else
        matriz(i, j) = n00 * n11 / n01 / n10
    End If

Next
Next

'TypeName(datos), muy interesante!




OddsRatioMatrix = matriz

End Function




Function OddsRatioMatrixPlus(datos)
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

    If (n01 = 0 Or n10 = 0 Or n00 = 0 Or n11 = 0) Then
        matriz(i, j) = (n00 + 0.5) * (n11 + 0.5) / (n01 + 0.5) / (n10 + 0.5)
        Else
        matriz(i, j) = n00 * n11 / n01 / n10
    End If
matriz(i, j) = Format(matriz(i, j), "0.00")
Next
Next


'https://www.medcalc.org/calc/odds_ratio.php


OddsRatioMatrixPlus = matriz

End Function


Function OddsRatioMatrixPlusIC(datos)
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

    If (n01 = 0 Or n10 = 0 Or n00 = 0 Or n11 = 0) Then
        OORR = (n00 + 0.5) * (n11 + 0.5) / (n01 + 0.5) / (n10 + 0.5)
        SE = (1 / (n00 + 0.5) + 1 / (n01 + 0.5) + 1 / (n10 + 0.5) + 1 / (n11 + 0.5)) ^ 0.5
        Else
        OORR = n00 * n11 / n01 / n10
        SE = (1 / n00 + 1 / n01 + 1 / n10 + 1 / n11) ^ 0.5
    End If
    inf = Format(Exp(Log(OORR) - 1.96 * SE), "0.00")
    sup = Format(Exp(Log(OORR) + 1.96 * SE), "0.00")
    matriz(i, j) = inf & "-" & sup

Next
Next


'https://www.medcalc.org/calc/odds_ratio.php


OddsRatioMatrixPlusIC = matriz

End Function

