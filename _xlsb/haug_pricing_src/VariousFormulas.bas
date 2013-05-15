Attribute VB_Name = "VariousFormulas"
Option Explicit     'Requirs that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.

Option Base 1       'The "Option Base" statment alowws to specify 0 or 1 as the
                            'default first index of arrays.
                            
' Programmer Espen Gaarder Haug Copyright 2006

'// Basket volatility
Public Function BasketVolatility(Weights As Variant, Vols As Variant, Correlations As Variant) As Double
        
        Dim n As Integer, i As Integer, j As Integer
        Dim sum As Double
        
        n = Application.Count(Weights)
        For i = 1 To n
                sum = sum + Weights(i) ^ 2 * Vols(i) ^ 2
            For j = i + 1 To n
                sum = sum + 2 * Weights(i) * Weights(j) * Vols(i) * Vols(j) * Correlations(i, j)
            Next
        Next
        BasketVolatility = Sqr(sum)
        
End Function


Public Function ImpliedCorCurOpt(v1 As Double, v2 As Double, v3 As Double) As Double
    'v1: Volatility currency 1: e.g. USD/EUR
    'v2: Volatility currency 2: e.g. USD/JPY
    'v3: Cross volatility currency 3: e.g. EUR/JPY
    
    ImpliedCorCurOpt = (v1 ^ 2 + v2 ^ 2 - v3 ^ 2) / (2 * v1 * v2)
    
End Function

Public Function ImpliedForwardVolatility(v1 As Double, v2 As Double, T1 As Double, T2 As Double) As Double
    
    ImpliedForwardVolatility = Sqr((v2 ^ 2 * T2 - v1 ^ 2 * T1) / (T2 - T1))

End Function

Public Function LowerBoundaryVolatility(v1 As Double, T1 As Double, T2 As Double) As Double
    
    LowerBoundaryVolatility = v1 * Sqr(T1 / T2)

End Function
