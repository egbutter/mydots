Attribute VB_Name = "TwoAssetOptions"
Option Explicit

' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug


'// Geometric average rate option
Public Function GeometricAverageRateOption(CallPutFlag As String, S As Double, X As Double, _
                T As Double, r As Double, b As Double, v As Double) As Double


    Dim bA As Double, vA As Double
    
    bA = 1 / 2 * (b - v ^ 2 / 6)
    vA = v / Sqr(3)

    GeometricAverageRateOption = GBlackScholes(CallPutFlag, S, X, T, r, bA, vA)


End Function


' Asian call option on the minimum of 2 Geometric averages
Public Function MinMaxTwoAverges(TypeFlag As String, S1 As Double, S2 As Double, X As Double, T As Double, r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double) As Double

    Dim a1 As Double, a2 As Double, a3 As Double, a4 As Double, a5 As Double, a6 As Double
    Dim mu1 As Double, mu2 As Double, v1s As Double, v2s As Double, vS As Double
    Dim S1s As Double, S2s As Double, CallMin As Double
    Dim rho1 As Double, rho2 As Double, v As Double
    Dim CallPutFlag As String
    
    If TypeFlag = "cmin" Or TypeFlag = "cmax" Then
        CallPutFlag = "c"
    Else
        CallPutFlag = "p"
    End If
        
    v1s = v1 / Sqr(3)
    v2s = v2 / Sqr(3)
    v = Sqr(v1 ^ 2 + v2 ^ 2 - 2 * rho * v1 * v2)
    vS = v / Sqr(3)
    
    mu1 = 0.5 * (b1 - v1 ^ 2 / 2) + 1 / 6 * v1 ^ 2
    mu2 = 0.5 * (b2 - v2 ^ 2 / 2) + 1 / 6 * v2 ^ 2

    S1s = S1 * Exp((mu1 - r) * T)
    S2s = S2 * Exp((mu2 - r) * T)

    a1 = (Log(S1s / X) + (b1 + v1s ^ 2 / 2) * T) / (v1s * Sqr(T))
    a2 = a1 - v1s * Sqr(T)
    a3 = (Log(S2s / X) + (b2 + v2s ^ 2 / 2) * T) / (v2s * Sqr(T))
    a4 = a3 - v2s * Sqr(T)
    a5 = (Log(S1s / S2s) - vS ^ 2 / 2 * T) / (vS * Sqr(T))
    a6 = (Log(S2s / S1s) - vS ^ 2 / 2 * T) / (vS * Sqr(T))
    
    rho1 = (rho * v2s - v1s) / vS
    rho2 = (rho * v1s - v2s) / vS
    
    CallMin = S1s * CBND(a1, a6, rho1) _
                       + S2s * CBND(a3, a5, rho2) _
                      - X * Exp(-r * T) * CBND(a2, a4, rho)
                      
    If TypeFlag = "cmin" Then
        MinMaxTwoAverges = CallMin
    ElseIf TypeFlag = "cmax" Then
       MinMaxTwoAverges = GeometricAverageRateOption(CallPutFlag, S1, X, T, r, b1, v1) _
         + GeometricAverageRateOption(CallPutFlag, S2, X, T, r, b2, v2) _
         - CallMin
    ElseIf TypeFlag = "pmin" Then
       MinMaxTwoAverges = X * Exp(-r * T) _
       - MinMaxTwoAverges("cmin", S1, S2, 0.0000001, T, r, b1, b2, v1, v2, rho) _
        + CallMin
     ElseIf TypeFlag = "pmax" Then
       MinMaxTwoAverges = X * Exp(-r * T) _
       - MinMaxTwoAverges("cmax", S1, S2, 0.0000001, T, r, b1, b2, v1, v2, rho) _
        + MinMaxTwoAverges("cmax", S1, S2, X, T, r, b1, b2, v1, v2, rho)
    End If
                      
End Function


'// Spread option approximation with drifts and Quantiities same results as above
Function SpreadApproximation(CallPutFlag As String, S1 As Double, S2 As Double, Q1 As Double, Q2 As Double, X As Double, T As Double, _
                r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double) As Double

    Dim v As Double, S As Double
    Dim d1 As Double, d2 As Double
    Dim F As Double
    F = Q2 * S2 * Exp((b2 - r) * T) / (Q2 * S2 * Exp((b2 - r) * T) + X * Exp(-r * T))
    v = Sqr(v1 ^ 2 + (v2 * F) ^ 2 - 2 * rho * v1 * v2 * F)
    S = Q1 * S1 * Exp((b1 - r) * T) / (Q2 * S2 * Exp((b2 - r) * T) + X * Exp(-r * T))
    d1 = (Log(S) + v ^ 2 / 2 * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then
         SpreadApproximation = (Q2 * S2 * Exp((b2 - r) * T) + X * Exp(-r * T)) * (S * CND(d1) - CND(d2))
    Else
        SpreadApproximation = (Q2 * S2 * Exp((b2 - r) * T) + X * Exp(-r * T)) * (CND(-d2) - S * CND(-d1))
    End If
    
End Function



'// Exchange options on exchange options
Public Function ExchangeExchangeOption(TypeFlag As Integer, S1 As Double, S2 As Double, q As Double, t1 As Double, T2 As Double, r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double) As Double
    
    Dim i As Double, I1 As Double
    Dim d1 As Double, d2 As Double
    Dim d3 As Double, d4 As Double
    Dim y1 As Double, y2 As Double
    Dim y3 As Double, y4 As Double
    Dim v As Double, id As Integer
    
    v = Sqr(v1 ^ 2 + v2 ^ 2 - 2 * rho * v1 * v2)
    I1 = S1 * Exp((b1 - r) * (T2 - t1)) / (S2 * Exp((b2 - r) * (T2 - t1)))
    
    If TypeFlag = 1 Or TypeFlag = 2 Then
        id = 1
    Else
        id = 2
    End If
    
    i = CriticalPrice(id, I1, t1, T2, v, q)
    d1 = (Log(S1 / (i * S2)) + (b1 - b2 + v ^ 2 / 2) * t1) / (v * Sqr(t1))
    d2 = d1 - v * Sqr(t1)
    d3 = (Log((i * S2) / S1) + (b2 - b1 + v ^ 2 / 2) * t1) / (v * Sqr(t1))
    d4 = d3 - v * Sqr(t1)
    y1 = (Log(S1 / S2) + (b1 - b2 + v ^ 2 / 2) * T2) / (v * Sqr(T2))
    y2 = y1 - v * Sqr(T2)
    y3 = (Log(S2 / S1) + (b2 - b1 + v ^ 2 / 2) * T2) / (v * Sqr(T2))
    y4 = y3 - v * Sqr(T2)
    
    If TypeFlag = 1 Then
        ExchangeExchangeOption = -S2 * Exp((b2 - r) * T2) * CBND(d2, y2, Sqr(t1 / T2)) + S1 * Exp((b1 - r) * T2) * CBND(d1, y1, Sqr(t1 / T2)) - q * S2 * Exp((b2 - r) * t1) * CND(d2)
    ElseIf TypeFlag = 2 Then
        ExchangeExchangeOption = S2 * Exp((b2 - r) * T2) * CBND(d3, y2, -Sqr(t1 / T2)) - S1 * Exp((b1 - r) * T2) * CBND(d4, y1, -Sqr(t1 / T2)) + q * S2 * Exp((b2 - r) * t1) * CND(d3)
    ElseIf TypeFlag = 3 Then
        ExchangeExchangeOption = S2 * Exp((b2 - r) * T2) * CBND(d3, y3, Sqr(t1 / T2)) - S1 * Exp((b1 - r) * T2) * CBND(d4, y4, Sqr(t1 / T2)) - q * S2 * Exp((b2 - r) * t1) * CND(d3)
    ElseIf TypeFlag = 4 Then
        ExchangeExchangeOption = -S2 * Exp((b2 - r) * T2) * CBND(d2, y3, -Sqr(t1 / T2)) + S1 * Exp((b1 - r) * T2) * CBND(d1, y4, -Sqr(t1 / T2)) + q * S2 * Exp((b2 - r) * t1) * CND(d2)
    End If
End Function
'// Numerical search algorithm to find critical price I
Private Function CriticalPrice(id As Integer, I1 As Double, t1 As Double, T2 As Double, v As Double, q As Double) As Double
    Dim Ii As Double, yi As Double, di As Double
    Dim epsilon As Double
    
    Ii = I1
    yi = CriticalPart3(id, Ii, t1, T2, v)
    di = CriticalPart2(id, Ii, t1, T2, v)
    epsilon = 0.00001
    While Abs(yi - q) > epsilon
        Ii = Ii - (yi - q) / di
        yi = CriticalPart3(id, Ii, t1, T2, v)
        di = CriticalPart2(id, Ii, t1, T2, v)
    Wend
    CriticalPrice = Ii
End Function
Private Function CriticalPart2(id As Integer, i As Double, t1 As Double, T2 As Double, v As Double) As Double
    Dim z1 As Double, z2 As Double
    If id = 1 Then
        z1 = (Log(i) + v ^ 2 / 2 * (T2 - t1)) / (v * Sqr(T2 - t1))
        CriticalPart2 = CND(z1)
    ElseIf id = 2 Then
        z2 = (-Log(i) - v ^ 2 / 2 * (T2 - t1)) / (v * Sqr(T2 - t1))
        CriticalPart2 = -CND(z2)
    End If
End Function
Private Function CriticalPart3(id As Integer, i As Double, t1 As Double, T2 As Double, v As Double) As Double
    Dim z1 As Double, z2 As Double
    If id = 1 Then
        z1 = (Log(i) + v ^ 2 / 2 * (T2 - t1)) / (v * Sqr(T2 - t1))
        z2 = (Log(i) - v ^ 2 / 2 * (T2 - t1)) / (v * Sqr(T2 - t1))
        CriticalPart3 = i * CND(z1) - CND(z2)
    ElseIf id = 2 Then
        z1 = (-Log(i) + v ^ 2 / 2 * (T2 - t1)) / (v * Sqr(T2 - t1))
        z2 = (-Log(i) - v ^ 2 / 2 * (T2 - t1)) / (v * Sqr(T2 - t1))
        CriticalPart3 = CND(z1) - i * CND(z2)
    End If
End Function

Public Function EExchangeExchangeOption(OutPutFlag As String, TypeFlag As Integer, S1 As Double, S2 As Double, q As Double, t1 As Double, T2 As Double, _
                                    r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        EExchangeExchangeOption = ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "d1" Then 'Delta S1
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "d2" Then 'Delta S2
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e1" Then 'Elasticity S1
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS) * S1 / ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "e2" Then 'Elasticity S2
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS) * S2 / ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "g1" Then 'Gamma S1
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho)) / dS ^ 2
      ElseIf OutPutFlag = "g2" Then 'Gamma S2
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv1" Then 'DGammaDVol S1
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) + ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) _
      - ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1 - dv, v2, rho) + 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 - dv, v2, rho) - ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "gv2" Then 'DGammaDVol S2
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) + ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) _
      - ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1, v2 - dv, rho) + 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 - dv, rho) - ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgv1" Then 'Cross GammaDvol S1v2 Cross Zomma
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) + ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) _
      - ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1, v2 - dv, rho) + 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 - dv, rho) - ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "cgv2" Then 'Cross GammaDvol S2v1 Cross Zomma
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) + ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) _
      - ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1 - dv, v2, rho) + 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 - dv, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
 
    ElseIf OutPutFlag = "gp1" Then 'GammaP S1
        EExchangeExchangeOption = S1 / 100 * (ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gp2" Then 'GammaP S2
        EExchangeExchangeOption = S2 / 100 * (ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "cg" Then 'Cross gamma
        EExchangeExchangeOption = 1 / (4 * dS * dS) * (ExchangeExchangeOption(TypeFlag, S1 + dS, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1 + dS, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho) _
        - ExchangeExchangeOption(TypeFlag, S1 - dS, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1 - dS, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "dddv1" Then 'DDeltaDvol S1 v1
        EExchangeExchangeOption = 1 / (4 * dS * dv) * (ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) - ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1 - dv, v2, rho) _
        - ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) + ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "dddv2" Then 'DDeltaDvol S2 v2
        EExchangeExchangeOption = 1 / (4 * dS * dv) * (ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) - ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1, v2 - dv, rho) _
        - ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) + ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / 100

   ElseIf OutPutFlag = "cv1" Then 'Cross vanna 1
        EExchangeExchangeOption = 1 / (4 * dS * dv) * (ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) - ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1, v2 - dv, rho) _
        - ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) + ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / 100
     ElseIf OutPutFlag = "cv2" Then 'Cross vanna 2
        EExchangeExchangeOption = 1 / (4 * dS * dv) * (ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1 - dv, v2, rho) _
        - ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) + ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "v1" Then 'Vega v1
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / 2
      ElseIf OutPutFlag = "v2" Then 'Vega v2
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "vv1" Then 'DvegaDvol/vomma v1
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / dv ^ 2 / 10000
      ElseIf OutPutFlag = "vv2" Then 'DvegaDvol/vomma v2
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / dv ^ 2 / 10000
   ElseIf OutPutFlag = "cv" Then 'Cross vomma
        EExchangeExchangeOption = 1 / (4 * dv * dv) * (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 + dv, v2 + dv, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 - dv, v2 + dv, rho) _
        - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 + dv, v2 - dv, rho) + ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 - dv, v2 - dv, rho)) / 10000
  
    ElseIf OutPutFlag = "vp1" Then 'VegaP v1
         EExchangeExchangeOption = v1 / 0.1 * (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / 2
        ElseIf OutPutFlag = "vp2" Then 'VegaP v2
         EExchangeExchangeOption = v2 / 0.1 * (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "dvdv1" Then 'DvegaDvol v1
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 + dv, v2, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1 - dv, v2, rho))
      ElseIf OutPutFlag = "dvdv2" Then 'DvegaDvol v2
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 + dv, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2 - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho + 0.01) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho - 0.01)) / 2
   ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho + 0.01) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If t1 <= 1 / 365 Then
                EExchangeExchangeOption = ExchangeExchangeOption(TypeFlag, S1, S2, q, 0.00001, T2 - 1 / 365, r, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho)
        Else
                EExchangeExchangeOption = ExchangeExchangeOption(TypeFlag, S1, S2, q, t1 - 1 / 365, T2 - 1 / 365, r, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r + 0.01, b1 + 0.01, b2 + 0.01, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r - 0.01, b1 - 0.01, b2 - 0.01, v1, v2, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r + 0.01, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r - 0.01, b1, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1 - 0.01, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1 - 0.01, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
  ElseIf OutPutFlag = "f2" Then 'Rho2
         EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2 - 0.01, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2 + 0.01, v1, v2, rho)) / (2)
  
    ElseIf OutPutFlag = "b1" Then 'Carry S1
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1 + 0.01, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1 - 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "b2" Then 'Carry S2
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2 + 0.01, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2 - 0.01, v1, v2, rho)) / (2)
   
    ElseIf OutPutFlag = "s1" Then 'Speed S1
        EExchangeExchangeOption = 1 / dS ^ 3 * (ExchangeExchangeOption(TypeFlag, S1 + 2 * dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho) - 3 * ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho) _
                                + 3 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "s2" Then 'Speed S2
        EExchangeExchangeOption = 1 / dS ^ 3 * (ExchangeExchangeOption(TypeFlag, S1, S2 + 2 * dS, q, t1, T2, r, b1, b2, v1, v2, rho) - 3 * ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) _
                                + 3 * ExchangeExchangeOption(TypeFlag, S1, S2, q, t1, T2, r, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "cs1" Then 'Cross speed S1^2S2
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1 + dS, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1 - dS, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) _
        - ExchangeExchangeOption(TypeFlag, S1 + dS, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho) + 2 * ExchangeExchangeOption(TypeFlag, S1, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1 - dS, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "cs2" Then 'Cross speed S2^2S1
        EExchangeExchangeOption = (ExchangeExchangeOption(TypeFlag, S1 + dS, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) - 2 * ExchangeExchangeOption(TypeFlag, S1 + dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho) + ExchangeExchangeOption(TypeFlag, S1 + dS, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho) _
        - ExchangeExchangeOption(TypeFlag, S1 - dS, S2 + dS, q, t1, T2, r, b1, b2, v1, v2, rho) + 2 * ExchangeExchangeOption(TypeFlag, S1 - dS, S2, q, t1, T2, r, b1, b2, v1, v2, rho) - ExchangeExchangeOption(TypeFlag, S1 - dS, S2 - dS, q, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   
    End If
End Function




Public Function ESpreadApproximation(OutPutFlag As String, CallPutFlag As String, S1 As Double, S2 As Double, Q1 As Double, Q2 As Double, X As Double, T As Double, _
                                    r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        ESpreadApproximation = SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "d1" Then 'Delta S1
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "d2" Then 'Delta S2
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e1" Then 'Elasticity S1
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S1 / SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "e2" Then 'Elasticity S2
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S2 / SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "g1" Then 'Gamma S1
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
      ElseIf OutPutFlag = "g2" Then 'Gamma S2
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv1" Then 'DGammaDVol S1
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) + SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) _
      - SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho) + 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho) - SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "gv2" Then 'DGammaDVol S2
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) + SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) _
      - SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho) + 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho) - SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgv1" Then 'Cross GammaDvol S1v2 Cross Zomma
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) + SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) _
      - SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho) + 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho) - SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "cgv2" Then 'Cross GammaDvol S2v1 Cross Zomma
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) + SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) _
      - SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho) + 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
 
    ElseIf OutPutFlag = "gp1" Then 'GammaP S1
        ESpreadApproximation = S1 / 100 * (SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gp2" Then 'GammaP S2
        ESpreadApproximation = S2 / 100 * (SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "cg" Then 'Cross gamma
        ESpreadApproximation = 1 / (4 * dS * dS) * (SpreadApproximation(CallPutFlag, S1 + dS, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1 + dS, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) _
        - SpreadApproximation(CallPutFlag, S1 - dS, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1 - dS, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "dddv1" Then 'DDeltaDvol S1 v1
        ESpreadApproximation = 1 / (4 * dS * dv) * (SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) - SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho) _
        - SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) + SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "dddv2" Then 'DDeltaDvol S2 v2
        ESpreadApproximation = 1 / (4 * dS * dv) * (SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) - SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho) _
        - SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) + SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 100

   ElseIf OutPutFlag = "cv1" Then 'Cross vanna 1
        ESpreadApproximation = 1 / (4 * dS * dv) * (SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) - SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho) _
        - SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) + SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 100
     ElseIf OutPutFlag = "cv2" Then 'Cross vanna 2
        ESpreadApproximation = 1 / (4 * dS * dv) * (SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho) _
        - SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) + SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "v1" Then 'Vega v1
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 2
      ElseIf OutPutFlag = "v2" Then 'Vega v2
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "vv1" Then 'DvegaDvol/vomma v1
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho)) / dv ^ 2 / 10000
      ElseIf OutPutFlag = "vv2" Then 'DvegaDvol/vomma v2
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho)) / dv ^ 2 / 10000
   ElseIf OutPutFlag = "cv" Then 'Cross vomma
        ESpreadApproximation = 1 / (4 * dv * dv) * (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2 + dv, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2 + dv, rho) _
        - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2 - dv, rho) + SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2 - dv, rho)) / 10000
  
    ElseIf OutPutFlag = "vp1" Then 'VegaP v1
         ESpreadApproximation = v1 / 0.1 * (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 2
        ElseIf OutPutFlag = "vp2" Then 'VegaP v2
         ESpreadApproximation = v2 / 0.1 * (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "dvdv1" Then 'DvegaDvol v1
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1 - dv, v2, rho))
      ElseIf OutPutFlag = "dvdv2" Then 'DvegaDvol v2
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2 - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho + 0.01) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho - 0.01)) / 2
   ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho + 0.01) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ESpreadApproximation = SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, 0.00001, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)
        Else
                ESpreadApproximation = SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T - 1 / 365, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r + 0.01, b1 + 0.01, b2 + 0.01, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r - 0.01, b1 - 0.01, b2 - 0.01, v1, v2, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r + 0.01, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r - 0.01, b1, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1 - 0.01, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1 - 0.01, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
  ElseIf OutPutFlag = "f2" Then 'Rho2
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2 - 0.01, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2 + 0.01, v1, v2, rho)) / (2)
  
    ElseIf OutPutFlag = "b1" Then 'Carry S1
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1 + 0.01, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1 - 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "b2" Then 'Carry S2
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2 + 0.01, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2 - 0.01, v1, v2, rho)) / (2)
   
    ElseIf OutPutFlag = "s1" Then 'Speed S1
        ESpreadApproximation = 1 / dS ^ 3 * (SpreadApproximation(CallPutFlag, S1 + 2 * dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - 3 * SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) _
                                + 3 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "s2" Then 'Speed S2
        ESpreadApproximation = 1 / dS ^ 3 * (SpreadApproximation(CallPutFlag, S1, S2 + 2 * dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - 3 * SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) _
                                + 3 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "cs1" Then 'Cross speed S1^2S2
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1 + dS, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1 - dS, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) _
        - SpreadApproximation(CallPutFlag, S1 + dS, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + 2 * SpreadApproximation(CallPutFlag, S1, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1 - dS, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "cs2" Then 'Cross speed S2^2S1
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1 + dS, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - 2 * SpreadApproximation(CallPutFlag, S1 + dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1 + dS, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) _
        - SpreadApproximation(CallPutFlag, S1 - dS, S2 + dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + 2 * SpreadApproximation(CallPutFlag, S1 - dS, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1 - dS, S2 - dS, Q1, Q2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X + dS, T, r, b1, b2, v1, v2, rho) - SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X - dS, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        ESpreadApproximation = (SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X + dS, T, r, b1, b2, v1, v2, rho) - 2 * SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X, T, r, b1, b2, v1, v2, rho) + SpreadApproximation(CallPutFlag, S1, S2, Q1, Q2, X - dS, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    End If
End Function





'// European option to exchange one asset for another
Public Function EuropeanExchangeOption(CallPutFlag As String, S1 As Double, S2 As Double, Q1 As Double, Q2 As Double, T As Double, r As Double, b1 As Double, _
                b2 As Double, v1 As Double, v2 As Double, rho As Double) As Double

    Dim v As Double, d1 As Double, d2 As Double

    v = Sqr(v1 ^ 2 + v2 ^ 2 - 2 * rho * v1 * v2)
    d1 = (Log(Q1 * S1 / (Q2 * S2)) + (b1 - b2 + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    If CallPutFlag = "c" Then
        EuropeanExchangeOption = Q1 * S1 * Exp((b1 - r) * T) * CND(d1) - Q2 * S2 * Exp((b2 - r) * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
                EuropeanExchangeOption = Q2 * S2 * Exp((b2 - r) * T) * CND(-d2) - Q1 * S1 * Exp((b1 - r) * T) * CND(-d1)
    End If
    
End Function
    
Public Function EEuropeanExchangeOption(OutPutFlag As String, CallPutFlag As String, S1 As Double, S2 As Double, Q1 As Double, Q2 As Double, T As Double, _
                                    r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        EEuropeanExchangeOption = EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "d1" Then 'Delta S1
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "d2" Then 'Delta S2
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e1" Then 'Elasticity S1
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S1 / EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "e2" Then 'Elasticity S2
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S2 / EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "g1" Then 'Gamma S1
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
      ElseIf OutPutFlag = "g2" Then 'Gamma S2
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv1" Then 'DGammaDVol S1
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) _
      - EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho) + 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "gv2" Then 'DGammaDVol S2
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) _
      - EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho) + 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgv1" Then 'Cross GammaDvol S1v2 Cross Zomma
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) + EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) _
      - EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho) + 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho) - EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "cgv2" Then 'Cross GammaDvol S2v1 Cross Zomma
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) _
      - EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho) + 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
 
    ElseIf OutPutFlag = "gp1" Then 'GammaP S1
        EEuropeanExchangeOption = S1 / 100 * (EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gp2" Then 'GammaP S2
        EEuropeanExchangeOption = S2 / 100 * (EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "cg" Then 'Cross gamma
        EEuropeanExchangeOption = 1 / (4 * dS * dS) * (EuropeanExchangeOption(CallPutFlag, S1 + dS, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1 + dS, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) _
        - EuropeanExchangeOption(CallPutFlag, S1 - dS, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1 - dS, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "dddv1" Then 'DDeltaDvol S1 v1
        EEuropeanExchangeOption = 1 / (4 * dS * dv) * (EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho) _
        - EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "dddv2" Then 'DDeltaDvol S2 v2
        EEuropeanExchangeOption = 1 / (4 * dS * dv) * (EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho) _
        - EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho)) / 100

   ElseIf OutPutFlag = "cv1" Then 'Cross vanna 1
        EEuropeanExchangeOption = 1 / (4 * dS * dv) * (EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) - EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho) _
        - EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) + EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho)) / 100
     ElseIf OutPutFlag = "cv2" Then 'Cross vanna 2
        EEuropeanExchangeOption = 1 / (4 * dS * dv) * (EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho) _
        - EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "v1" Then 'Vega v1
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho)) / 2
      ElseIf OutPutFlag = "v2" Then 'Vega v2
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "vv1" Then 'DvegaDvol/vomma v1
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho)) / dv ^ 2 / 10000
      ElseIf OutPutFlag = "vv2" Then 'DvegaDvol/vomma v2
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho)) / dv ^ 2 / 10000
   ElseIf OutPutFlag = "cv" Then 'Cross vomma
        EEuropeanExchangeOption = 1 / (4 * dv * dv) * (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2 + dv, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2 + dv, rho) _
        - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2 - dv, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2 - dv, rho)) / 10000
  
    ElseIf OutPutFlag = "vp1" Then 'VegaP v1
         EEuropeanExchangeOption = v1 / 0.1 * (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho)) / 2
        ElseIf OutPutFlag = "vp2" Then 'VegaP v2
         EEuropeanExchangeOption = v2 / 0.1 * (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "dvdv1" Then 'DvegaDvol v1
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 + dv, v2, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1 - dv, v2, rho))
      ElseIf OutPutFlag = "dvdv2" Then 'DvegaDvol v2
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 + dv, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2 - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho + 0.01) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho - 0.01)) / 2
   ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho + 0.01) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EEuropeanExchangeOption = EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, 0.00001, r, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho)
        Else
                EEuropeanExchangeOption = EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T - 1 / 365, r, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r + 0.01, b1 + 0.01, b2 + 0.01, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r - 0.01, b1 - 0.01, b2 - 0.01, v1, v2, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r + 0.01, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r - 0.01, b1, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1 - 0.01, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1 - 0.01, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
  ElseIf OutPutFlag = "f2" Then 'Rho2
         EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2 - 0.01, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2 + 0.01, v1, v2, rho)) / (2)
  
    ElseIf OutPutFlag = "b1" Then 'Carry S1
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1 + 0.01, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1 - 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "b2" Then 'Carry S2
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2 + 0.01, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2 - 0.01, v1, v2, rho)) / (2)
   
    ElseIf OutPutFlag = "s1" Then 'Speed S1
        EEuropeanExchangeOption = 1 / dS ^ 3 * (EuropeanExchangeOption(CallPutFlag, S1 + 2 * dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) - 3 * EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) _
                                + 3 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "s2" Then 'Speed S2
        EEuropeanExchangeOption = 1 / dS ^ 3 * (EuropeanExchangeOption(CallPutFlag, S1, S2 + 2 * dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) - 3 * EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) _
                                + 3 * EuropeanExchangeOption(CallPutFlag, S1, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "cs1" Then 'Cross speed S1^2S2
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1 + dS, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1 - dS, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) _
        - EuropeanExchangeOption(CallPutFlag, S1 + dS, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) + 2 * EuropeanExchangeOption(CallPutFlag, S1, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1 - dS, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "cs2" Then 'Cross speed S2^2S1
        EEuropeanExchangeOption = (EuropeanExchangeOption(CallPutFlag, S1 + dS, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) - 2 * EuropeanExchangeOption(CallPutFlag, S1 + dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) + EuropeanExchangeOption(CallPutFlag, S1 + dS, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) _
        - EuropeanExchangeOption(CallPutFlag, S1 - dS, S2 + dS, Q1, Q2, T, r, b1, b2, v1, v2, rho) + 2 * EuropeanExchangeOption(CallPutFlag, S1 - dS, S2, Q1, Q2, T, r, b1, b2, v1, v2, rho) - EuropeanExchangeOption(CallPutFlag, S1 - dS, S2 - dS, Q1, Q2, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   
    End If
End Function




'// Two asset correlation options
Public Function TwoAssetCorrelation(CallPutFlag As String, S1 As Double, S2 As Double, X1 As Double, X2 As Double, T As Double, _
                b1 As Double, b2 As Double, r As Double, v1 As Double, v2 As Double, rho As Double)

    Dim y1 As Double, y2 As Double
   
    y1 = (Log(S1 / X1) + (b1 - v1 ^ 2 / 2) * T) / (v1 * Sqr(T))
    y2 = (Log(S2 / X2) + (b2 - v2 ^ 2 / 2) * T) / (v2 * Sqr(T))
    
    If CallPutFlag = "c" Then
        TwoAssetCorrelation = S2 * Exp((b2 - r) * T) * CBND(y2 + v2 * Sqr(T), y1 + rho * v2 * Sqr(T), rho) _
        - X2 * Exp(-r * T) * CBND(y2, y1, rho)
    ElseIf CallPutFlag = "p" Then
         TwoAssetCorrelation = X2 * Exp(-r * T) * CBND(-y2, -y1, rho) _
         - S2 * Exp((b2 - r) * T) * CBND(-y2 - v2 * Sqr(T), -y1 - rho * v2 * Sqr(T), rho)
    End If
End Function


Public Function ETwoAssetCorrelation(OutPutFlag As String, CallPutFlag As String, S1 As Double, S2 As Double, X1 As Double, X2 As Double, T As Double, _
                                    r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        ETwoAssetCorrelation = TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "d1" Then 'Delta S1
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "d2" Then 'Delta S2
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e1" Then 'Elasticity S1
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S1 / TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "e2" Then 'Elasticity S2
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S2 / TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "g1" Then 'Gamma S1
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
      ElseIf OutPutFlag = "g2" Then 'Gamma S2
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv1" Then 'DGammaDVol S1
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) _
      - TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1 - dv, v2, rho) + 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 - dv, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "gv2" Then 'DGammaDVol S2
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) _
      - TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1, v2 - dv, rho) + 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 - dv, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgv1" Then 'Cross GammaDvol S1v2 Cross Zomma
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) _
      - TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1, v2 - dv, rho) + 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 - dv, rho) - TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "cgv2" Then 'Cross GammaDvol S2v1 Cross Zomma
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) _
      - TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1 - dv, v2, rho) + 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 - dv, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
 
    ElseIf OutPutFlag = "gp1" Then 'GammaP S1
        ETwoAssetCorrelation = S1 / 100 * (TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gp2" Then 'GammaP S2
        ETwoAssetCorrelation = S2 / 100 * (TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "cg" Then 'Cross gamma
        ETwoAssetCorrelation = 1 / (4 * dS * dS) * (TwoAssetCorrelation(CallPutFlag, S1 + dS, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1 + dS, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho) _
        - TwoAssetCorrelation(CallPutFlag, S1 - dS, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1 - dS, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "dddv1" Then 'DDeltaDvol S1 v1
        ETwoAssetCorrelation = 1 / (4 * dS * dv) * (TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1 - dv, v2, rho) _
        - TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "dddv2" Then 'DDeltaDvol S2 v2
        ETwoAssetCorrelation = 1 / (4 * dS * dv) * (TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1, v2 - dv, rho) _
        - TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1, v2 - dv, rho)) / 100

   ElseIf OutPutFlag = "cv1" Then 'Cross vanna 1
        ETwoAssetCorrelation = 1 / (4 * dS * dv) * (TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1, v2 - dv, rho) _
        - TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1, v2 - dv, rho)) / 100
     ElseIf OutPutFlag = "cv2" Then 'Cross vanna 2
        ETwoAssetCorrelation = 1 / (4 * dS * dv) * (TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1 - dv, v2, rho) _
        - TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "v1" Then 'Vega v1
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 - dv, v2, rho)) / 2
      ElseIf OutPutFlag = "v2" Then 'Vega v2
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "vv1" Then 'DvegaDvol/vomma v1
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 - dv, v2, rho)) / dv ^ 2 / 10000
      ElseIf OutPutFlag = "vv2" Then 'DvegaDvol/vomma v2
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 - dv, rho)) / dv ^ 2 / 10000
   ElseIf OutPutFlag = "cv" Then 'Cross vomma
        ETwoAssetCorrelation = 1 / (4 * dv * dv) * (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 + dv, v2 + dv, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 - dv, v2 + dv, rho) _
        - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 + dv, v2 - dv, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 - dv, v2 - dv, rho)) / 10000
  
    ElseIf OutPutFlag = "vp1" Then 'VegaP v1
         ETwoAssetCorrelation = v1 / 0.1 * (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 - dv, v2, rho)) / 2
        ElseIf OutPutFlag = "vp2" Then 'VegaP v2
         ETwoAssetCorrelation = v2 / 0.1 * (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "dvdv1" Then 'DvegaDvol v1
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1 - dv, v2, rho))
      ElseIf OutPutFlag = "dvdv2" Then 'DvegaDvol v2
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2 - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho + 0.01) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho - 0.01)) / 2
   ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho + 0.01) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ETwoAssetCorrelation = TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, 0.00001, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho)
        Else
                ETwoAssetCorrelation = TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T - 1 / 365, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r + 0.01, b1 + 0.01, b2 + 0.01, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r - 0.01, b1 - 0.01, b2 - 0.01, v1, v2, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r + 0.01, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r - 0.01, b1, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1 - 0.01, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1 - 0.01, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
  ElseIf OutPutFlag = "f2" Then 'Rho2
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2 - 0.01, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2 + 0.01, v1, v2, rho)) / (2)
  
    ElseIf OutPutFlag = "b1" Then 'Carry S1
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1 + 0.01, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1 - 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "b2" Then 'Carry S2
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2 + 0.01, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2 - 0.01, v1, v2, rho)) / (2)
   
    ElseIf OutPutFlag = "s1" Then 'Speed S1
        ETwoAssetCorrelation = 1 / dS ^ 3 * (TwoAssetCorrelation(CallPutFlag, S1 + 2 * dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho) - 3 * TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho) _
                                + 3 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "s2" Then 'Speed S2
        ETwoAssetCorrelation = 1 / dS ^ 3 * (TwoAssetCorrelation(CallPutFlag, S1, S2 + 2 * dS, X1, X2, T, r, b1, b2, v1, v2, rho) - 3 * TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) _
                                + 3 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "cs1" Then 'Cross speed S1^2S2
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1 + dS, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1 - dS, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) _
        - TwoAssetCorrelation(CallPutFlag, S1 + dS, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho) + 2 * TwoAssetCorrelation(CallPutFlag, S1, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1 - dS, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "cs2" Then 'Cross speed S2^2S1
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1 + dS, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1 + dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1 + dS, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho) _
        - TwoAssetCorrelation(CallPutFlag, S1 - dS, S2 + dS, X1, X2, T, r, b1, b2, v1, v2, rho) + 2 * TwoAssetCorrelation(CallPutFlag, S1 - dS, S2, X1, X2, T, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1 - dS, S2 - dS, X1, X2, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx1" Then 'Strike Delta X!
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1 + dS, X2, T, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1 - dS, X2, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
        ElseIf OutPutFlag = "dx2" Then 'Strike Delta X2
         ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2 + dS, T, r, b1, b2, v1, v2, rho) - TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2 - dS, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
 
     ElseIf OutPutFlag = "dxdx1" Then 'GammaX1
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1 + dS, X2, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2, X1 - dS, X2, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
       ElseIf OutPutFlag = "dxdx2" Then 'GammaX2
        ETwoAssetCorrelation = (TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2 + dS, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2, T, r, b1, b2, v1, v2, rho) + TwoAssetCorrelation(CallPutFlag, S1, S2, X1, X2 - dS, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     
    End If
End Function

'// Options on the maximum or minimum of two risky assets
Public Function OptionsOnTheMaxMin(TypeFlag As String, S1 As Double, S2 As Double, X As Double, T As Double, r As Double, _
        b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double) As Double

    Dim v As Double, rho1 As Double, rho2 As Double
    Dim d As Double, y1 As Double, y2 As Double
    
    v = Sqr(v1 ^ 2 + v2 ^ 2 - 2 * rho * v1 * v2)
    rho1 = (v1 - rho * v2) / v
    rho2 = (v2 - rho * v1) / v
    d = (Log(S1 / S2) + (b1 - b2 + v ^ 2 / 2) * T) / (v * Sqr(T))
    y1 = (Log(S1 / X) + (b1 + v1 ^ 2 / 2) * T) / (v1 * Sqr(T))
    y2 = (Log(S2 / X) + (b2 + v2 ^ 2 / 2) * T) / (v2 * Sqr(T))
  
    If TypeFlag = "cmin" Then
        OptionsOnTheMaxMin = S1 * Exp((b1 - r) * T) * CBND(y1, -d, -rho1) + S2 * Exp((b2 - r) * T) * CBND(y2, d - v * Sqr(T), -rho2) - X * Exp(-r * T) * CBND(y1 - v1 * Sqr(T), y2 - v2 * Sqr(T), rho)
    ElseIf TypeFlag = "cmax" Then
        OptionsOnTheMaxMin = S1 * Exp((b1 - r) * T) * CBND(y1, d, rho1) + S2 * Exp((b2 - r) * T) * CBND(y2, -d + v * Sqr(T), rho2) - X * Exp(-r * T) * (1 - CBND(-y1 + v1 * Sqr(T), -y2 + v2 * Sqr(T), rho))
    ElseIf TypeFlag = "pmin" Then
        OptionsOnTheMaxMin = X * Exp(-r * T) - S1 * Exp((b1 - r) * T) + EuropeanExchangeOption("c", S1, S2, 1, 1, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin("cmin", S1, S2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf TypeFlag = "pmax" Then
        OptionsOnTheMaxMin = X * Exp(-r * T) - S2 * Exp((b2 - r) * T) - EuropeanExchangeOption("c", S1, S2, 1, 1, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin("cmax", S1, S2, X, T, r, b1, b2, v1, v2, rho)
    End If
End Function



Public Function EOptionsOnTheMaxMin(OutPutFlag As String, CallPutFlag As String, S1 As Double, S2 As Double, X As Double, T As Double, _
                                    r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        EOptionsOnTheMaxMin = OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "d1" Then 'Delta S1
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "d2" Then 'Delta S2
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e1" Then 'Elasticity S1
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S1 / OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "e2" Then 'Elasticity S2
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S2 / OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "g1" Then 'Gamma S1
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
      ElseIf OutPutFlag = "g2" Then 'Gamma S2
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv1" Then 'DGammaDVol S1
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) _
      - OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho) + 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "gv2" Then 'DGammaDVol S2
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 + dv, rho) _
      - OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 - dv, rho) + 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgv1" Then 'Cross GammaDvol S1v2 Cross Zomma
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) + OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) _
      - OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho) + 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho) - OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "cgv2" Then 'Cross GammaDvol S2v1 Cross Zomma
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 + dv, v2, rho) _
      - OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 - dv, v2, rho) + 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
 
    ElseIf OutPutFlag = "gp1" Then 'GammaP S1
        EOptionsOnTheMaxMin = S1 / 100 * (OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gp2" Then 'GammaP S2
        EOptionsOnTheMaxMin = S2 / 100 * (OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "cg" Then 'Cross gamma
        EOptionsOnTheMaxMin = 1 / (4 * dS * dS) * (OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho) _
        - OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "dddv1" Then 'DDeltaDvol S1 v1
        EOptionsOnTheMaxMin = 1 / (4 * dS * dv) * (OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho) _
        - OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "dddv2" Then 'DDeltaDvol S2 v2
        EOptionsOnTheMaxMin = 1 / (4 * dS * dv) * (OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 + dv, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 - dv, rho) _
        - OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 + dv, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 - dv, rho)) / 100

   ElseIf OutPutFlag = "cv1" Then 'Cross vanna 1
        EOptionsOnTheMaxMin = 1 / (4 * dS * dv) * (OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho) _
        - OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) + OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 100
     ElseIf OutPutFlag = "cv2" Then 'Cross vanna 2
        EOptionsOnTheMaxMin = 1 / (4 * dS * dv) * (OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 + dv, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 - dv, v2, rho) _
        - OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 + dv, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "v1" Then 'Vega v1
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 2
      ElseIf OutPutFlag = "v2" Then 'Vega v2
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "vv1" Then 'DvegaDvol/vomma v1
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / dv ^ 2 / 10000
      ElseIf OutPutFlag = "vv2" Then 'DvegaDvol/vomma v2
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / dv ^ 2 / 10000
   ElseIf OutPutFlag = "cv" Then 'Cross vomma
        EOptionsOnTheMaxMin = 1 / (4 * dv * dv) * (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2 + dv, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2 + dv, rho) _
        - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2 - dv, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2 - dv, rho)) / 10000
  
    ElseIf OutPutFlag = "vp1" Then 'VegaP v1
         EOptionsOnTheMaxMin = v1 / 0.1 * (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 2
        ElseIf OutPutFlag = "vp2" Then 'VegaP v2
         EOptionsOnTheMaxMin = v2 / 0.1 * (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "dvdv1" Then 'DvegaDvol v1
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho))
      ElseIf OutPutFlag = "dvdv2" Then 'DvegaDvol v2
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho + 0.01) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho - 0.01)) / 2
   ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho + 0.01) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EOptionsOnTheMaxMin = OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, 0.00001, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
        Else
                EOptionsOnTheMaxMin = OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T - 1 / 365, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r + 0.01, b1 + 0.01, b2 + 0.01, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r - 0.01, b1 - 0.01, b2 - 0.01, v1, v2, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r + 0.01, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r - 0.01, b1, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1 - 0.01, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1 - 0.01, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
  ElseIf OutPutFlag = "f2" Then 'Rho2
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2 - 0.01, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2 + 0.01, v1, v2, rho)) / (2)
  
    ElseIf OutPutFlag = "b1" Then 'Carry S1
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1 + 0.01, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1 - 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "b2" Then 'Carry S2
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2 + 0.01, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2 - 0.01, v1, v2, rho)) / (2)
   
    ElseIf OutPutFlag = "s1" Then 'Speed S1
        EOptionsOnTheMaxMin = 1 / dS ^ 3 * (OptionsOnTheMaxMin(CallPutFlag, S1 + 2 * dS, S2, X, T, r, b1, b2, v1, v2, rho) - 3 * OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) _
                                + 3 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "s2" Then 'Speed S2
        EOptionsOnTheMaxMin = 1 / dS ^ 3 * (OptionsOnTheMaxMin(CallPutFlag, S1, S2 + 2 * dS, X, T, r, b1, b2, v1, v2, rho) - 3 * OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) _
                                + 3 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "cs1" Then 'Cross speed S1^2S2
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) _
        - OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho) + 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "cs2" Then 'Cross speed S2^2S1
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1 + dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho) _
        - OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) + 2 * OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1 - dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X + dS, T, r, b1, b2, v1, v2, rho) - OptionsOnTheMaxMin(CallPutFlag, S1, S2, X - dS, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EOptionsOnTheMaxMin = (OptionsOnTheMaxMin(CallPutFlag, S1, S2, X + dS, T, r, b1, b2, v1, v2, rho) - 2 * OptionsOnTheMaxMin(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OptionsOnTheMaxMin(CallPutFlag, S1, S2, X - dS, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    End If
End Function





'// Product Option
Public Function ProductOption(CallPutFlag As String, S1 As Double, S2 As Double, X As Double, T As Double, r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double) As Double
    Dim d1 As Double, d2 As Double, F As Double, v As Double
    
    F = S1 * S2 * Exp((b1 + b2 + rho * v1 * v2) * T)
    v = Sqr(v1 ^ 2 + v2 ^ 2 + 2 * rho * v1 * v2)
    d1 = (Log(F / X) + T * v ^ 2 / 2) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then
        ProductOption = Exp(-r * T) * (F * CND(d1) - X * CND(d2))
    Else
        ProductOption = Exp(-r * T) * (X * CND(-d2) - F * CND(-d1))
    End If
End Function



Public Function EProductOption(OutPutFlag As String, CallPutFlag As String, S1 As Double, S2 As Double, X As Double, T As Double, _
                                    r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        EProductOption = ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "d1" Then 'Delta S1
         EProductOption = (ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "d2" Then 'Delta S2
         EProductOption = (ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e1" Then 'Elasticity S1
         EProductOption = (ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S1 / ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "e2" Then 'Elasticity S2
         EProductOption = (ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S2 / ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "g1" Then 'Gamma S1
        EProductOption = (ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
      ElseIf OutPutFlag = "g2" Then 'Gamma S2
        EProductOption = (ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv1" Then 'DGammaDVol S1
        EProductOption = (ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) + ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) _
      - ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho) + 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho) - ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "gv2" Then 'DGammaDVol S2
        EProductOption = (ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) + ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 + dv, rho) _
      - ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 - dv, rho) + 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho) - ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgv1" Then 'Cross GammaDvol S1v2 Cross Zomma
        EProductOption = (ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) + ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) _
      - ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho) + 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho) - ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "cgv2" Then 'Cross GammaDvol S2v1 Cross Zomma
        EProductOption = (ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) + ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 + dv, v2, rho) _
      - ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 - dv, v2, rho) + 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho) - ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
 
    ElseIf OutPutFlag = "gp1" Then 'GammaP S1
        EProductOption = S1 / 100 * (ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gp2" Then 'GammaP S2
        EProductOption = S2 / 100 * (ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "cg" Then 'Cross gamma
        EProductOption = 1 / (4 * dS * dS) * (ProductOption(CallPutFlag, S1 + dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1 + dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho) _
        - ProductOption(CallPutFlag, S1 - dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1 - dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "dddv1" Then 'DDeltaDvol S1 v1
        EProductOption = 1 / (4 * dS * dv) * (ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho) _
        - ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) + ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "dddv2" Then 'DDeltaDvol S2 v2
        EProductOption = 1 / (4 * dS * dv) * (ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 + dv, rho) - ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 - dv, rho) _
        - ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 + dv, rho) + ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 - dv, rho)) / 100

   ElseIf OutPutFlag = "cv1" Then 'Cross vanna 1
        EProductOption = 1 / (4 * dS * dv) * (ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho) _
        - ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) + ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 100
     ElseIf OutPutFlag = "cv2" Then 'Cross vanna 2
        EProductOption = 1 / (4 * dS * dv) * (ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 + dv, v2, rho) - ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 - dv, v2, rho) _
        - ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 + dv, v2, rho) + ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "v1" Then 'Vega v1
         EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 2
      ElseIf OutPutFlag = "v2" Then 'Vega v2
         EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "vv1" Then 'DvegaDvol/vomma v1
        EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / dv ^ 2 / 10000
      ElseIf OutPutFlag = "vv2" Then 'DvegaDvol/vomma v2
        EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / dv ^ 2 / 10000
   ElseIf OutPutFlag = "cv" Then 'Cross vomma
        EProductOption = 1 / (4 * dv * dv) * (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2 + dv, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2 + dv, rho) _
        - ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2 - dv, rho) + ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2 - dv, rho)) / 10000
  
    ElseIf OutPutFlag = "vp1" Then 'VegaP v1
         EProductOption = v1 / 0.1 * (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 2
        ElseIf OutPutFlag = "vp2" Then 'VegaP v2
         EProductOption = v2 / 0.1 * (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "dvdv1" Then 'DvegaDvol v1
        EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho))
      ElseIf OutPutFlag = "dvdv2" Then 'DvegaDvol v2
        EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho + 0.01) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho - 0.01)) / 2
   ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho + 0.01) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EProductOption = ProductOption(CallPutFlag, S1, S2, X, 0.00001, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
        Else
                EProductOption = ProductOption(CallPutFlag, S1, S2, X, T - 1 / 365, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r + 0.01, b1 + 0.01, b2 + 0.01, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r - 0.01, b1 - 0.01, b2 - 0.01, v1, v2, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r + 0.01, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r - 0.01, b1, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1 - 0.01, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1 - 0.01, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
  ElseIf OutPutFlag = "f2" Then 'Rho2
         EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2 - 0.01, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2 + 0.01, v1, v2, rho)) / (2)
  
    ElseIf OutPutFlag = "b1" Then 'Carry S1
        EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1 + 0.01, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1 - 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "b2" Then 'Carry S2
        EProductOption = (ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2 + 0.01, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2 - 0.01, v1, v2, rho)) / (2)
   
    ElseIf OutPutFlag = "s1" Then 'Speed S1
        EProductOption = 1 / dS ^ 3 * (ProductOption(CallPutFlag, S1 + 2 * dS, S2, X, T, r, b1, b2, v1, v2, rho) - 3 * ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) _
                                + 3 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "s2" Then 'Speed S2
        EProductOption = 1 / dS ^ 3 * (ProductOption(CallPutFlag, S1, S2 + 2 * dS, X, T, r, b1, b2, v1, v2, rho) - 3 * ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) _
                                + 3 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "cs1" Then 'Cross speed S1^2S2
        EProductOption = (ProductOption(CallPutFlag, S1 + dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * ProductOption(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1 - dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) _
        - ProductOption(CallPutFlag, S1 + dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho) + 2 * ProductOption(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1 - dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "cs2" Then 'Cross speed S2^2S1
        EProductOption = (ProductOption(CallPutFlag, S1 + dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * ProductOption(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1 + dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho) _
        - ProductOption(CallPutFlag, S1 - dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) + 2 * ProductOption(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1 - dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EProductOption = (ProductOption(CallPutFlag, S1, S2, X + dS, T, r, b1, b2, v1, v2, rho) - ProductOption(CallPutFlag, S1, S2, X - dS, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EProductOption = (ProductOption(CallPutFlag, S1, S2, X + dS, T, r, b1, b2, v1, v2, rho) - 2 * ProductOption(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + ProductOption(CallPutFlag, S1, S2, X - dS, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    End If
End Function



 

'// Relative out-performance options
Public Function OutPerformance(CallPutFlag As String, S1 As Double, S2 As Double, X As Double, T As Double, r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double) As Double
    Dim d1 As Double, d2 As Double, F As Double, v As Double
    
    F = S1 / S2 * Exp((b1 - b2 + v2 ^ 2 - rho * v1 * v2) * T)
    v = Sqr(v1 ^ 2 + v2 ^ 2 - 2 * rho * v1 * v2)
    d1 = (Log(F / X) + T * v ^ 2 / 2) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then
        OutPerformance = Exp(-r * T) * (F * CND(d1) - X * CND(d2))
    Else
        OutPerformance = Exp(-r * T) * (X * CND(-d2) - F * CND(-d1))
    End If
End Function




Public Function EOutPerformance(OutPutFlag As String, CallPutFlag As String, S1 As Double, S2 As Double, X As Double, T As Double, _
                                    r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        EOutPerformance = OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "d1" Then 'Delta S1
         EOutPerformance = (OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "d2" Then 'Delta S2
         EOutPerformance = (OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e1" Then 'Elasticity S1
         EOutPerformance = (OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S1 / OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "e2" Then 'Elasticity S2
         EOutPerformance = (OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S2 / OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "g1" Then 'Gamma S1
        EOutPerformance = (OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
      ElseIf OutPutFlag = "g2" Then 'Gamma S2
        EOutPerformance = (OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv1" Then 'DGammaDVol S1
        EOutPerformance = (OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) + OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) _
      - OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho) + 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho) - OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "gv2" Then 'DGammaDVol S2
        EOutPerformance = (OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) + OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 + dv, rho) _
      - OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 - dv, rho) + 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho) - OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgv1" Then 'Cross GammaDvol S1v2 Cross Zomma
        EOutPerformance = (OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) + OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) _
      - OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho) + 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho) - OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "cgv2" Then 'Cross GammaDvol S2v1 Cross Zomma
        EOutPerformance = (OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) + OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 + dv, v2, rho) _
      - OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 - dv, v2, rho) + 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho) - OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
 
    ElseIf OutPutFlag = "gp1" Then 'GammaP S1
        EOutPerformance = S1 / 100 * (OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gp2" Then 'GammaP S2
        EOutPerformance = S2 / 100 * (OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "cg" Then 'Cross gamma
        EOutPerformance = 1 / (4 * dS * dS) * (OutPerformance(CallPutFlag, S1 + dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1 + dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho) _
        - OutPerformance(CallPutFlag, S1 - dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1 - dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "dddv1" Then 'DDeltaDvol S1 v1
        EOutPerformance = 1 / (4 * dS * dv) * (OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho) _
        - OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 + dv, v2, rho) + OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "dddv2" Then 'DDeltaDvol S2 v2
        EOutPerformance = 1 / (4 * dS * dv) * (OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 + dv, rho) - OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2 - dv, rho) _
        - OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 + dv, rho) + OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2 - dv, rho)) / 100

   ElseIf OutPutFlag = "cv1" Then 'Cross vanna 1
        EOutPerformance = 1 / (4 * dS * dv) * (OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho) _
        - OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 + dv, rho) + OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 100
     ElseIf OutPutFlag = "cv2" Then 'Cross vanna 2
        EOutPerformance = 1 / (4 * dS * dv) * (OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 + dv, v2, rho) - OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1 - dv, v2, rho) _
        - OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 + dv, v2, rho) + OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "v1" Then 'Vega v1
         EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 2
      ElseIf OutPutFlag = "v2" Then 'Vega v2
         EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "vv1" Then 'DvegaDvol/vomma v1
        EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / dv ^ 2 / 10000
      ElseIf OutPutFlag = "vv2" Then 'DvegaDvol/vomma v2
        EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / dv ^ 2 / 10000
   ElseIf OutPutFlag = "cv" Then 'Cross vomma
        EOutPerformance = 1 / (4 * dv * dv) * (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2 + dv, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2 + dv, rho) _
        - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2 - dv, rho) + OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2 - dv, rho)) / 10000
  
    ElseIf OutPutFlag = "vp1" Then 'VegaP v1
         EOutPerformance = v1 / 0.1 * (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho)) / 2
        ElseIf OutPutFlag = "vp2" Then 'VegaP v2
         EOutPerformance = v2 / 0.1 * (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "dvdv1" Then 'DvegaDvol v1
        EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 + dv, v2, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1 - dv, v2, rho))
      ElseIf OutPutFlag = "dvdv2" Then 'DvegaDvol v2
        EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 + dv, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2 - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho + 0.01) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho - 0.01)) / 2
   ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho + 0.01) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EOutPerformance = OutPerformance(CallPutFlag, S1, S2, X, 0.00001, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
        Else
                EOutPerformance = OutPerformance(CallPutFlag, S1, S2, X, T - 1 / 365, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r + 0.01, b1 + 0.01, b2 + 0.01, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r - 0.01, b1 - 0.01, b2 - 0.01, v1, v2, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r + 0.01, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r - 0.01, b1, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1 - 0.01, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1 - 0.01, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
  ElseIf OutPutFlag = "f2" Then 'Rho2
         EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2 - 0.01, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2 + 0.01, v1, v2, rho)) / (2)
  
    ElseIf OutPutFlag = "b1" Then 'Carry S1
        EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1 + 0.01, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1 - 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "b2" Then 'Carry S2
        EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2 + 0.01, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2 - 0.01, v1, v2, rho)) / (2)
   
    ElseIf OutPutFlag = "s1" Then 'Speed S1
        EOutPerformance = 1 / dS ^ 3 * (OutPerformance(CallPutFlag, S1 + 2 * dS, S2, X, T, r, b1, b2, v1, v2, rho) - 3 * OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) _
                                + 3 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "s2" Then 'Speed S2
        EOutPerformance = 1 / dS ^ 3 * (OutPerformance(CallPutFlag, S1, S2 + 2 * dS, X, T, r, b1, b2, v1, v2, rho) - 3 * OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) _
                                + 3 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "cs1" Then 'Cross speed S1^2S2
        EOutPerformance = (OutPerformance(CallPutFlag, S1 + dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * OutPerformance(CallPutFlag, S1, S2 + dS, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1 - dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) _
        - OutPerformance(CallPutFlag, S1 + dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho) + 2 * OutPerformance(CallPutFlag, S1, S2 - dS, X, T, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1 - dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "cs2" Then 'Cross speed S2^2S1
        EOutPerformance = (OutPerformance(CallPutFlag, S1 + dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) - 2 * OutPerformance(CallPutFlag, S1 + dS, S2, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1 + dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho) _
        - OutPerformance(CallPutFlag, S1 - dS, S2 + dS, X, T, r, b1, b2, v1, v2, rho) + 2 * OutPerformance(CallPutFlag, S1 - dS, S2, X, T, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1 - dS, S2 - dS, X, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X + dS, T, r, b1, b2, v1, v2, rho) - OutPerformance(CallPutFlag, S1, S2, X - dS, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EOutPerformance = (OutPerformance(CallPutFlag, S1, S2, X + dS, T, r, b1, b2, v1, v2, rho) - 2 * OutPerformance(CallPutFlag, S1, S2, X, T, r, b1, b2, v1, v2, rho) + OutPerformance(CallPutFlag, S1, S2, X - dS, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    End If
End Function


