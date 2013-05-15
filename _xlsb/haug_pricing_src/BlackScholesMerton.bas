Attribute VB_Name = "BlackScholesMerton"
Option Explicit

' Programmer Espen Gaarder Haug
' Copyright Espen G. Haug 2006

'//  The generalized Black and Scholes formula
Public Function GBlackScholes(CallPutFlag As String, S As Double, X _
                As Double, T As Double, r As Double, b As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    d1 = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then
        GBlackScholes = S * Exp((b - r) * T) * CND(d1) - X * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
        GBlackScholes = X * Exp(-r * T) * CND(-d2) - S * Exp((b - r) * T) * CND(-d1)
    End If
    
End Function


Public Function ImpliedVolGBlackScholes(CallPutFlag As String, S As Double, _
                X As Double, T As Double, r As Double, b As Double, cm As Double) As Double

    Dim vLow As Double, vHigh As Double, vi As Double
    Dim cLow As Double, cHigh As Double, epsilon As Double
    Dim N As Integer
    
    vLow = 0.005
    vHigh = 4
    epsilon = 0.00000001
    cLow = GBlackScholes(CallPutFlag, S, X, T, r, b, vLow)
    cHigh = GBlackScholes(CallPutFlag, S, X, T, r, b, vHigh)
    N = 0
    vi = vLow + (cm - cLow) * (vHigh - vLow) / (cHigh - cLow)
    While Abs(cm - GBlackScholes(CallPutFlag, S, X, T, r, b, vi)) > epsilon
        N = N + 1
        If N > 20 Then Exit Function
        
        If GBlackScholes(CallPutFlag, S, X, T, r, b, vi) < cm Then
            vLow = vi
        Else
            vHigh = vi
        End If
        cLow = GBlackScholes(CallPutFlag, S, X, T, r, b, vLow)
        cHigh = GBlackScholes(CallPutFlag, S, X, T, r, b, vHigh)
        vi = vLow + (cm - cLow) * (vHigh - vLow) / (cHigh - cLow)
    Wend
    ImpliedVolGBlackScholes = vi
    
End Function
