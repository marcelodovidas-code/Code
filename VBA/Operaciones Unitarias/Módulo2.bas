Attribute VB_Name = "Módulo2"


Function ResolverTridiagonalMatrix(rangoA As Range, _
                                   rangoB As Range, _
                                   rangoC As Range, _
                                   rangoD As Range) As Variant
    Dim n As Long
    Dim A() As Double, B() As Double, C() As Double, D() As Double
    Dim Cpr() As Double, Dpr() As Double, X() As Double
    Dim i As Long, denom As Double
    
    ' Cantidad de ecuaciones (usa la diagonal principal)
    n = rangoB.Count
    
    ' Usamos índices 0..n-1 para calzar con tus fórmulas (n = 0,1,...)
    ReDim A(0 To n - 1)
    ReDim B(0 To n - 1)
    ReDim C(0 To n - 1)
    ReDim D(0 To n - 1)
    ReDim Cpr(0 To n - 1)
    ReDim Dpr(0 To n - 1)
    ReDim X(0 To n - 1)
    
    ' Cargar coeficientes desde Excel
    For i = 0 To n - 1
    A(i) = rangoA.Cells(i + 1, 1).Value
        B(i) = rangoB.Cells(i + 1, 1).Value
        D(i) = rangoD.Cells(i + 1, 1).Value
        
        If i < n - 1 Then
            C(i) = rangoC.Cells(i + 1, 1).Value
        Else
            C(i) = 0
        End If
    Next i
    
    ' ============================
    '   FORWARD SWEEP (C', D')
    ' ============================
    ' Según tus fórmulas:
    ' C'0 = C0 / B0
    ' D'0 = D0 / B0
    ' C'n = Cn / (Bn - An C'n-1)
    ' D'n = (Dn - An D'n-1) / (Bn - An C'n-1)
    
    Cpr(0) = C(0) / B(0)
    Dpr(0) = D(0) / B(0)
    
    For i = 1 To n - 1
        denom = B(i) - A(i) * Cpr(i - 1)
        Cpr(i) = C(i) / denom
        Dpr(i) = (D(i) - A(i) * Dpr(i - 1)) / denom
    Next i
    
    ' ============================
    '   BACK SUBSTITUTION (x)
    ' ============================
    ' x_n = D'_n - C'_n x_{n+1}
    ' excepto el último:
    ' x_(n-1) = D'_(n-1)
    
    X(n - 1) = Dpr(n - 1)
    
    For i = n - 2 To 0 Step -1
        X(i) = Dpr(i) - Cpr(i) * X(i + 1)
    Next i
    
    ' Devolver vector columna
    ResolverTridiagonalMatrix = Application.Transpose(X)
End Function

