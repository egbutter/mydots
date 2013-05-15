Attribute VB_Name = "VarSwapVBA"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.

Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
                    

' Programmer Espen Gaarder Haug
' Copyright Espen Gaarder Haug 2006

Public Function WeightVarSwap(S As Double, x As Double, T As Double, r As Double, b As Double) As Double

    WeightVarSwap = 2 / T * ((x - S) / S - Log(x / S))

End Function

Public Function GVarVega(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    GVarVega = S * Exp((b - r) * T) * Sqr(T) / (2 * v) * Exp(-1 * d1 ^ 2 / 2) / (Sqr(2 * Pi))

End Function


'// Variance-vomma for the generalized Black and Scholes formula
Public Function GVarianceVomma(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GVarianceVomma = S * Exp((b - r) * T) * Sqr(T) / (4 * v ^ 3) * ND(d1) * (d1 * d2 - 1)

End Function

'// Variance-ultima for the generalized Black and Scholes formula
Public Function GVarianceUltima(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GVarianceUltima = S * Exp((b - r) * T) * Sqr(T) / (8 * v ^ 5) * ND(d1) * ((d1 * d2 - 1) * (d1 * d2 - 3) - (d1 ^ 2 + d2 ^ 2))
    
End Function

'// The generalized Black and Scholes formula
Public Function GBlackScholes(CallPutFlag As String, S As Double, x _
                As Double, T As Double, r As Double, b As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then
        GBlackScholes = S * Exp((b - r) * T) * CND(d1) - x * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
        GBlackScholes = x * Exp(-r * T) * CND(-d2) - S * Exp((b - r) * T) * CND(-d1)
    End If
End Function























