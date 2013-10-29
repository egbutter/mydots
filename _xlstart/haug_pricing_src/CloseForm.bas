Attribute VB_Name = "CloseForm"
Option Explicit

' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug



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

'// Two asset correlation options
Public Function TwoAssetCorrelation(CallPutFlag As String, S1 As Double, S2 As Double, x1 As Double, x2 As Double, T As Double, _
                b1 As Double, b2 As Double, r As Double, v1 As Double, v2 As Double, rho As Double)

    Dim y1 As Double, y2 As Double
   
    y1 = (Log(S1 / x1) + (b1 - v1 ^ 2 / 2) * T) / (v1 * Sqr(T))
    y2 = (Log(S2 / x2) + (b2 - v2 ^ 2 / 2) * T) / (v2 * Sqr(T))
    
    If CallPutFlag = "c" Then
        TwoAssetCorrelation = S2 * Exp((b2 - r) * T) * CBND(y2 + v2 * Sqr(T), y1 + rho * v1 * Sqr(T), rho) _
        - x2 * Exp(-r * T) * CBND(y2, y1, rho)
    ElseIf CallPutFlag = "p" Then
         TwoAssetCorrelation = x2 * Exp(-r * T) * CBND(-y2, -y1, rho) _
         - S2 * Exp((b2 - r) * T) * CBND(-y2 - v2 * Sqr(T), -y1 - rho * v2 * Sqr(T), rho)
    End If
End Function


