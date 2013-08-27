Attribute VB_Name = "TwoAssetBinomial"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.

Option Base 0       'The "Option Base" statement allows to specify 0 or 1 as the
                          'default first index of arrays.

' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug
                    
Public Function Max(X, y)
Attribute Max.VB_ProcData.VB_Invoke_Func = " \r14"
            Max = Application.Max(X, y)
End Function

Public Function Min(X, y)
Attribute Min.VB_ProcData.VB_Invoke_Func = " \r14"
            Min = Application.Min(X, y)
End Function


'// Three dimensional European only binomial tree
Public Function ThreeDimensionalBinomiaEuropean(TypeFlag As Integer, AmeEurFlag As String, CallPutFlag As String, _
                S1 As Double, S2 As Double, Q1 As Double, Q2 As Double, X1 As Double, X2 As Double, T As Double, r As Double, b1 As Double, _
                b2 As Double, v1 As Double, v2 As Double, rho As Double, n As Integer) As Double
    
  
    Dim dt As Double, u As Double, d As Double
    Dim my1 As Double, my2 As Double
    Dim y1 As Double, y2
    Dim NodeValueS1 As Double, NodeValueS2 As Double
    Dim i As Integer, j As Integer
    Dim sum As Double
    Dim PatheProbability As Double

    dt = T / n
    my1 = b1 - v1 ^ 2 / 2
    my2 = b2 - v2 ^ 2 / 2
    u = Exp(my1 * dt + v1 * Sqr(dt))
    d = Exp(my1 * dt - v1 * Sqr(dt))
    

    For j = 0 To n
        NodeValueS1 = S1 * u ^ j * d ^ (n - j)
        
        For i = 0 To n
            
             NodeValueS2 = S2 * Exp(my2 * T + v2 * (rho * (2 * j - n) + Sqr(1 - rho ^ 2) * (2 * i - n)) * Sqr(dt))
            
            PatheProbability = Application.Combin(n, i) * Application.Combin(n, j) * 0.25 ^ n
            sum = sum + PatheProbability * PayoffFunction(TypeFlag, CallPutFlag, NodeValueS1, NodeValueS2, Q1, Q2, X1, X2)
        Next
    
    Next
    
    ThreeDimensionalBinomiaEuropean = sum * Exp(-r * T)
End Function


'// Three dimensional binomial tree
Public Function ThreeDimensionalBinomial(TypeFlag As Integer, AmeEurFlag As String, CallPutFlag As String, _
                S1 As Double, S2 As Double, Q1 As Double, Q2 As Double, X1 As Double, X2 As Double, T As Double, r As Double, b1 As Double, _
                b2 As Double, v1 As Double, v2 As Double, rho As Double, n As Integer) As Double
Attribute ThreeDimensionalBinomial.VB_ProcData.VB_Invoke_Func = " \r14"
    
    Dim OptionValue() As Double
    Dim dt As Double, u As Double, d As Double
    Dim mu1 As Double, mu2 As Double
    Dim y1 As Double, y2 As Double
    Dim NodeValueS1 As Double, NodeValueS2 As Double
    Dim i As Integer, j As Integer, m As Integer
    
    ReDim OptionValue(0 To n + 1, 0 To n + 1)

    dt = T / n
    mu1 = b1 - v1 ^ 2 / 2
    mu2 = b2 - v2 ^ 2 / 2
    u = Exp(mu1 * dt + v1 * Sqr(dt))
    d = Exp(mu1 * dt - v1 * Sqr(dt))

    For j = 0 To n
        y1 = (2 * j - n) * Sqr(dt)
        NodeValueS1 = S1 * u ^ j * d ^ (n - j)
        For i = 0 To n
            NodeValueS2 = S2 * Exp(mu2 * n * dt) * Exp(v2 * (rho * y1 + Sqr(1 - rho ^ 2) * (2 * i - n) * Sqr(dt)))
            OptionValue(j, i) = PayoffFunction(TypeFlag, CallPutFlag, NodeValueS1, NodeValueS2, Q1, Q2, X1, X2)
        Next
    Next
    For m = n - 1 To 0 Step -1
        For j = 0 To m
            y1 = (2 * j - m) * Sqr(dt)
            NodeValueS1 = S1 * u ^ j * d ^ (m - j)
            For i = 0 To m
                y2 = rho * y1 + Sqr(1 - rho ^ 2) * (2 * i - m) * Sqr(dt)
                NodeValueS2 = S2 * Exp(mu2 * m * dt) * Exp(v2 * y2)
                OptionValue(j, i) = 1 / 4 * (OptionValue(j, i) + OptionValue(j + 1, i) + OptionValue(j, i + 1) _
                + OptionValue(j + 1, i + 1)) * Exp(-r * dt)
                If AmeEurFlag = "a" Then
                    OptionValue(j, i) = Max(OptionValue(j, i), PayoffFunction(TypeFlag, CallPutFlag, NodeValueS1, NodeValueS2, Q1, Q2, X1, X2))
                End If
            Next
        Next
    Next
    ThreeDimensionalBinomial = OptionValue(0, 0)
End Function



'// Payoff function used in three dimensional binomial tree
Public Function PayoffFunction(TypeFlag As Integer, CallPutFlag As String, S1 As Double, S2 As Double, _
        Q1 As Double, Q2 As Double, X1 As Double, X2 As Double) As Double
Attribute PayoffFunction.VB_ProcData.VB_Invoke_Func = " \r14"

    Dim z As Integer
    
    If CallPutFlag = "c" Then
        z = 1
        ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    If TypeFlag = 1 Then '// Spread option
        PayoffFunction = Max(0, z * (Q1 * S1 - Q2 * S2) - z * X1)
    ElseIf TypeFlag = 2 Then  '// Option on the maximum of two assets
        PayoffFunction = Max(0, z * Max(Q1 * S1, Q2 * S2) - z * X1)
    ElseIf TypeFlag = 3 Then  '// Option on the minimum of two assets
        PayoffFunction = Max(0, z * Min(Q1 * S1, Q2 * S2) - z * X1)
    ElseIf TypeFlag = 4 Then  '// Dual strike option
        PayoffFunction = Application.Max(0, z * (Q1 * S1 - X1), z * (Q2 * S2 - X2))
    ElseIf TypeFlag = 5 Then  '// Reverse-dual strike option
        PayoffFunction = Application.Max(0, z * (Q1 * S1 - X1), z * (X2 - Q2 * S2))
    ElseIf TypeFlag = 6 Then  '// Portfolio option
        PayoffFunction = Max(0, z * (Q1 * S1 + Q2 * S2) - z * X1)
    ElseIf TypeFlag = 7 Then  '// Exchange option
         PayoffFunction = PayoffFunction = Max(0, Q2 * S2 - Q1 * S1)
  ElseIf TypeFlag = 8 Then  '// Outperformance option
        PayoffFunction = Max(0, z * (Q1 * S1 / (Q2 * S2) - X1))
    ElseIf TypeFlag = 9 Then  '// Product option
        PayoffFunction = Max(0, z * (Q1 * S1 * Q2 * S2 - X1))
    End If
End Function


