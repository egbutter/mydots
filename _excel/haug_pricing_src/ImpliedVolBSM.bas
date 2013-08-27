Attribute VB_Name = "ImpliedVolBSM"

Global Const Pi = 3.14159265358979
Option Explicit     'Requirs that all variables to be declared explicitly.

' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug

Public Function GBlackScholesImpVolBisection(CallPutFlag As String, S As Double, _
                X As Double, T As Double, r As Double, b As Double, cm As Double) As Double

    Dim vLow As Double, vHigh As Double, vi As Double
    Dim cLow As Double, cHigh As Double, epsilon As Double
    Dim counter As Integer
    
    vLow = 0.005
    vHigh = 4
    epsilon = 0.00000001
    cLow = GBlackScholes(CallPutFlag, S, X, T, r, b, vLow)
    cHigh = GBlackScholes(CallPutFlag, S, X, T, r, b, vHigh)
    counter = 0
    vi = vLow + (cm - cLow) * (vHigh - vLow) / (cHigh - cLow)
    While Abs(cm - GBlackScholes(CallPutFlag, S, X, T, r, b, vi)) > epsilon
        counter = counter + 1
        If counter = 100 Then Exit Function
        If GBlackScholes(CallPutFlag, S, X, T, r, b, vi) < cm Then
            vLow = vi
        Else
            vHigh = vi
        End If
        cLow = GBlackScholes(CallPutFlag, S, X, T, r, b, vLow)
        cHigh = GBlackScholes(CallPutFlag, S, X, T, r, b, vHigh)
        vi = vLow + (cm - cLow) * (vHigh - vLow) / (cHigh - cLow)
    Wend
    GBlackScholesImpVolBisection = vi
    
End Function


Public Function GImpliedVolatilityNR(CallPutFlag As String, S As Double, X _
As Double, T As Double, r As Double, b As Double, cm As Double, epsilon As Double)

    Dim vi As Double, ci As Double
    Dim vegai As Double
    Dim minDiff As Double

    'Manaster and Koehler seed value (vi)
    vi = Sqr(Abs(Log(S / X) + r * T) * 2 / T)
    ci = GBlackScholes(CallPutFlag, S, X, T, r, b, vi)
    vegai = GVega(S, X, T, r, b, vi)
    minDiff = Abs(cm - ci)
    
    While Abs(cm - ci) >= epsilon And Abs(cm - ci) <= minDiff
        vi = vi - (ci - cm) / vegai
        ci = GBlackScholes(CallPutFlag, S, X, T, r, b, vi)
        vegai = GVega(S, X, T, r, b, vi)
        minDiff = Abs(cm - ci)
    Wend

    If Abs(cm - ci) < epsilon Then GImpliedVolatilityNR = vi Else GImpliedVolatilityNR = "NA"
End Function



