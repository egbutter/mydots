Attribute VB_Name = "Barriers"
Option Explicit

' Programmer Espen Gaarder Haug
' Copyright Espen Gaarder Haug  2006




Public Function ESoftBarrier(OutPutFlag As String, TypeFlag As String, S As Double, X As Double, L As Double, U As Double, T As Double, _
            r As Double, b As Double, v As Double, Optional dS)

    If IsMissing(dS) Then
        dS = 0.0001
    End If
  
    
    
    If OutPutFlag = "p" Then 'Value
            ESoftBarrier = SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v)
    ElseIf OutPutFlag = "d" Then ' Delta
            ESoftBarrier = (SoftBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v) _
                 - SoftBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dddv" Then ' DeltaDVol
                 ESoftBarrier = 1 / (4 * dS * 0.01) * (SoftBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v + 0.01) - SoftBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v - 0.01) _
                - SoftBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v + 0.01) + SoftBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v - 0.01)) / 100
     ElseIf OutPutFlag = "g" Then ' Gamma
            ESoftBarrier = (SoftBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v) _
            - 2 * SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v) _
            + SoftBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v)) / (dS ^ 2)
      ElseIf OutPutFlag = "gp" Then ' GammaP
            ESoftBarrier = S / 100 * ESoftBarrier("g", TypeFlag, S + dS, X, L, U, T, r, b, v)
    ElseIf OutPutFlag = "gv" Then ' DGammaDVol
             ESoftBarrier = (SoftBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v + 0.01) - 2 * SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v + 0.01) + SoftBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v + 0.01) _
                - SoftBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v - 0.01) + 2 * SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v - 0.01) - SoftBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "v" Then ' Vega
            ESoftBarrier = (SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v + 0.01) _
                - SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v - 0.01)) / 2
    ElseIf OutPutFlag = "dvdv" Then ' DVegaDVol/Vomma
                  ESoftBarrier = (SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v + 0.01) - 2 * SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v) _
                        + SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then ' VegaP
            ESoftBarrier = v / 0.1 * ESoftBarrier("v", TypeFlag, S + dS, X, L, U, T, r, b, v)
    ElseIf OutPutFlag = "r" Then ' Rho
            ESoftBarrier = (SoftBarrier(TypeFlag, S, X, L, U, T, r + 0.01, b + 0.01, v) _
            - SoftBarrier(TypeFlag, S, X, L, U, T, r - 0.01, b - 0.01, v)) / 2
    ElseIf OutPutFlag = "f" Then 'Rho2/Phi
            ESoftBarrier = (SoftBarrier(TypeFlag, S, X, L, U, T, r, b - 0.01, v) _
            - SoftBarrier(TypeFlag, S, X, L, U, T, r, b + 0.01, v)) / 2
    ElseIf OutPutFlag = "b" Then ' Carry sensitivity
            ESoftBarrier = (SoftBarrier(TypeFlag, S, X, L, U, T, r, b + 0.01, v) _
            - SoftBarrier(TypeFlag, S, X, L, U, T, r, b - 0.01, v)) / 2
    ElseIf OutPutFlag = "t" Then 'Theta
            If T <= 1 / 365 Then
                ESoftBarrier = SoftBarrier(TypeFlag, S, X, L, U, 0.00001, r, b, v) _
                    - SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v)
              Else
                ESoftBarrier = SoftBarrier(TypeFlag, S, X, L, U, T - 1 / 365, r, b, v) _
                    - SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v)
                End If
    ElseIf OutPutFlag = "dx" Then 'Strike Delta
            ESoftBarrier = (SoftBarrier(TypeFlag, S, X + dS, L, U, T, r, b, v) _
            - SoftBarrier(TypeFlag, S, X - dS, L, U, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "dxdx" Then ' Strike Gamma
            ESoftBarrier = (SoftBarrier(TypeFlag, S, X + dS, L, U, T, r, b, v) _
            - 2 * SoftBarrier(TypeFlag, S, X, L, U, T, r, b, v) _
            + SoftBarrier(TypeFlag, S, X - dS, L, U, T, r, b, v)) / (dS ^ 2)
    End If
End Function

'// Soft barrier options
Public Function SoftBarrier(TypeFlag As String, S As Double, X As Double, _
                            L As Double, U As Double, T As Double, r As Double, b As Double, v As Double) As Double

    Dim mu As Double
    Dim d1 As Double, d2 As Double
    Dim d3 As Double, d4 As Double
    Dim e1 As Double, e2 As Double
    Dim e3 As Double, e4 As Double
    Dim Lambda1 As Double, Lambda2 As Double
    Dim Value As Double, eta As Integer
    
    If TypeFlag = "cdi" Or TypeFlag = "cdo" Then
        eta = 1
    Else
        eta = -1
    End If
    
    mu = (b + v ^ 2 / 2) / v ^ 2
    Lambda1 = Exp(-1 / 2 * v ^ 2 * T * (mu + 0.5) * (mu - 0.5))
    Lambda2 = Exp(-1 / 2 * v ^ 2 * T * (mu - 0.5) * (mu - 1.5))
    d1 = Log(U ^ 2 / (S * X)) / (v * Sqr(T)) + mu * v * Sqr(T)
    d2 = d1 - (mu + 0.5) * v * Sqr(T)
    d3 = Log(U ^ 2 / (S * X)) / (v * Sqr(T)) + (mu - 1) * v * Sqr(T)
    d4 = d3 - (mu - 0.5) * v * Sqr(T)
    e1 = Log(L ^ 2 / (S * X)) / (v * Sqr(T)) + mu * v * Sqr(T)
    e2 = e1 - (mu + 0.5) * v * Sqr(T)
    e3 = Log(L ^ 2 / (S * X)) / (v * Sqr(T)) + (mu - 1) * v * Sqr(T)
    e4 = e3 - (mu - 0.5) * v * Sqr(T)
    
    Value = eta * 1 / (U - L) * (S * Exp((b - r) * T) * S ^ (-2 * mu) _
    * (S * X) ^ (mu + 0.5) / (2 * (mu + 0.5)) _
    * ((U ^ 2 / (S * X)) ^ (mu + 0.5) * CND(eta * d1) - Lambda1 * CND(eta * d2) _
    - (L ^ 2 / (S * X)) ^ (mu + 0.5) * CND(eta * e1) + Lambda1 * CND(eta * e2)) _
    - X * Exp(-r * T) * S ^ (-2 * (mu - 1)) _
    * (S * X) ^ (mu - 0.5) / (2 * (mu - 0.5)) _
    * ((U ^ 2 / (S * X)) ^ (mu - 0.5) * CND(eta * d3) - Lambda2 * CND(eta * d4) _
    - (L ^ 2 / (S * X)) ^ (mu - 0.5) * CND(eta * e3) + Lambda2 * CND(eta * e4)))
    
    If TypeFlag = "cdi" Or TypeFlag = "pui" Then
        SoftBarrier = Value
    ElseIf TypeFlag = "cdo" Then
        SoftBarrier = GBlackScholes("c", S, X, T, r, b, v) - Value
    ElseIf TypeFlag = "puo" Then
        SoftBarrier = GBlackScholes("p", S, X, T, r, b, v) - Value
    End If
    
End Function

Public Function ELookBarrier(OutPutFlag As String, TypeFlag As String, S As Double, X As Double, H As Double, t1 As Double, T2 As Double, _
            r As Double, b As Double, v As Double, Optional dS)

    Dim CallPutFlag As String
            
    If IsMissing(dS) Then
        dS = 0.0001
    End If
     
    CallPutFlag = Left(TypeFlag, 1)
    
    If (TypeFlag = "cuo" And S >= H) Or (TypeFlag = "pdo" And S <= H) Then
        ELookBarrier = 0
    ElseIf (TypeFlag = "cui" And S >= H) Or (TypeFlag = "pdi" And S <= H) Then
        ELookBarrier = PartialFixedLB(CallPutFlag, S, X, t1, T2, r, b, v)
        Exit Function
    End If
      
    If OutPutFlag = "p" Then 'Value
            ELookBarrier = LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
            ELookBarrier = (LookBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v) _
                         - LookBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "dddv" Then 'DeltaDVol
            ELookBarrier = 1 / (4 * dS * 0.01) * (LookBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v + 0.01) - LookBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v - 0.01) _
        - LookBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v + 0.01) + LookBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v - 0.01)) / 100
     ElseIf OutPutFlag = "g" Then 'Gamma
            ELookBarrier = (LookBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v) _
            - 2 * LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v) _
            + LookBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v)) / (dS ^ 2)
    ElseIf OutPutFlag = "gp" Then ' GammaP
            ELookBarrier = S / 100 * ELookBarrier("g", TypeFlag, S + dS, X, H, t1, T2, r, b, v)
    ElseIf OutPutFlag = "gv" Then 'DGammaDvol
            ELookBarrier = (LookBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v + 0.01) - 2 * LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v + 0.01) + LookBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v + 0.01) _
                - LookBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v - 0.01) + 2 * LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v - 0.01) - LookBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "v" Then ' Vega
            ELookBarrier = (LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v + 0.01) _
                    - LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vp" Then ' VegaP
            ELookBarrier = v / 0.1 * ELookBarrier("v", TypeFlag, S + dS, X, H, t1, T2, r, b, v)
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol/vomma
            ELookBarrier = (LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v + 0.01) - 2 * LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v) _
            + LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "r" Then 'Rho
            ELookBarrier = (LookBarrier(TypeFlag, S, X, H, t1, T2, r + 0.01, b + 0.01, v) _
            - LookBarrier(TypeFlag, S, X, H, t1, T2, r - 0.01, b - 0.01, v)) / 2
    ElseIf OutPutFlag = "f" Then 'Rho2 Phi
            ELookBarrier = (LookBarrier(TypeFlag, S, X, H, t1, T2, r, b - 0.01, v) _
            - LookBarrier(TypeFlag, S, X, H, t1, T2, r, b + 0.01, v)) / 2
    ElseIf OutPutFlag = "b" Then ' Carry sensitivity
            ELookBarrier = (LookBarrier(TypeFlag, S, X, H, t1, T2, r, b + 0.01, v) _
            - LookBarrier(TypeFlag, S, X, H, t1, T2, r, b - 0.01, v)) / 2
      ElseIf OutPutFlag = "t" Then 'Theta
            If t1 <= 1 / 365 Then
                ELookBarrier = LookBarrier(TypeFlag, S, X, H, 0.00001, T2 - 1 / 365, r, b, v) _
                    - LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v)
              Else
                    ELookBarrier = LookBarrier(TypeFlag, S, X, H, t1 - 1 / 365, T2 - 1 / 365, r, b, v) _
                    - LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v)
            End If
     ElseIf OutPutFlag = "dx" Then 'Strike Delta
            ELookBarrier = (LookBarrier(TypeFlag, S, X + dS, H, t1, T2, r, b, v) _
                - LookBarrier(TypeFlag, S, X - dS, H, t1, T2, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
      
            ELookBarrier = (LookBarrier(TypeFlag, S, X + dS, H, t1, T2, r, b, v) _
                - 2 * LookBarrier(TypeFlag, S, X, H, t1, T2, r, b, v) _
                + LookBarrier(TypeFlag, S, X - dS, H, t1, T2, r, b, v)) / (dS ^ 2)
    End If
    
End Function

'// Look-barrier options
Public Function LookBarrier(TypeFlag As String, S As Double, X As Double, H As Double, t1 As Double, T2 As Double, r As Double, b As Double, v As Double) As Double

    Dim hh As Double
    Dim k As Double, mu1 As Double, mu2 As Double
    Dim rho As Double, eta As Double, m As Double
    Dim g1 As Double, g2 As Double
    Dim OutValue As Double, part1 As Double, part2 As Double, part3 As Double, part4 As Double
    
    hh = Log(H / S)
    k = Log(X / S)
    mu1 = b - v ^ 2 / 2
    mu2 = b + v ^ 2 / 2
    rho = Sqr(t1 / T2)
    
    If TypeFlag = "cuo" Or TypeFlag = "cui" Then
        eta = 1
        m = Min(hh, k)
    ElseIf TypeFlag = "pdo" Or TypeFlag = "pdi" Then
        eta = -1
        m = Max(hh, k)
    End If
    
    g1 = (CND(eta * (hh - mu2 * t1) / (v * Sqr(t1))) - Exp(2 * mu2 * hh / v ^ 2) * CND(eta * (-hh - mu2 * t1) / (v * Sqr(t1)))) _
        - (CND(eta * (m - mu2 * t1) / (v * Sqr(t1))) - Exp(2 * mu2 * hh / v ^ 2) * CND(eta * (m - 2 * hh - mu2 * t1) / (v * Sqr(t1))))
    g2 = (CND(eta * (hh - mu1 * t1) / (v * Sqr(t1))) - Exp(2 * mu1 * hh / v ^ 2) * CND(eta * (-hh - mu1 * t1) / (v * Sqr(t1)))) _
        - (CND(eta * (m - mu1 * t1) / (v * Sqr(t1))) - Exp(2 * mu1 * hh / v ^ 2) * CND(eta * (m - 2 * hh - mu1 * t1) / (v * Sqr(t1))))

    part1 = S * Exp((b - r) * T2) * (1 + v ^ 2 / (2 * b)) * (CBND(eta * (m - mu2 * t1) / (v * Sqr(t1)), eta * (-k + mu2 * T2) / (v * Sqr(T2)), -rho) - Exp(2 * mu2 * hh / v ^ 2) _
        * CBND(eta * (m - 2 * hh - mu2 * t1) / (v * Sqr(t1)), eta * (2 * hh - k + mu2 * T2) / (v * Sqr(T2)), -rho))
    part2 = -Exp(-r * T2) * X * (CBND(eta * (m - mu1 * t1) / (v * Sqr(t1)), eta * (-k + mu1 * T2) / (v * Sqr(T2)), -rho) _
        - Exp(2 * mu1 * hh / v ^ 2) * CBND(eta * (m - 2 * hh - mu1 * t1) / (v * Sqr(t1)), eta * (2 * hh - k + mu1 * T2) / (v * Sqr(T2)), -rho))
    part3 = -Exp(-r * T2) * v ^ 2 / (2 * b) * (S * (S / X) ^ (-2 * b / v ^ 2) * CBND(eta * (m + mu1 * t1) / (v * Sqr(t1)), eta * (-k - mu1 * T2) / (v * Sqr(T2)), -rho) _
        - H * (H / X) ^ (-2 * b / v ^ 2) * CBND(eta * (m - 2 * hh + mu1 * t1) / (v * Sqr(t1)), eta * (2 * hh - k - mu1 * T2) / (v * Sqr(T2)), -rho))
    part4 = S * Exp((b - r) * T2) * ((1 + v ^ 2 / (2 * b)) * CND(eta * mu2 * (T2 - t1) / (v * Sqr(T2 - t1))) + Exp(-b * (T2 - t1)) * (1 - v ^ 2 / (2 * b)) _
        * CND(eta * (-mu1 * (T2 - t1)) / (v * Sqr(T2 - t1)))) * g1 - Exp(-r * T2) * X * g2
    OutValue = eta * (part1 + part2 + part3 + part4)

    If TypeFlag = "cuo" Or TypeFlag = "pdo" Then
        LookBarrier = OutValue
    ElseIf TypeFlag = "cui" Then
        LookBarrier = PartialFixedLB("c", S, X, t1, T2, r, b, v) - OutValue
    ElseIf TypeFlag = "pdi" Then
        LookBarrier = PartialFixedLB("p", S, X, t1, T2, r, b, v) - OutValue
    End If
    
End Function


Public Function EPartialTimeBarrier(OutPutFlag As String, TypeFlag As String, S As Double, X As Double, H As Double, t1 As Double, T2 As Double, _
            r As Double, b As Double, v As Double, Optional dS)

    If IsMissing(dS) Then
        dS = 0.0001
    End If
    
    If (TypeFlag = "cuoA" And S >= H) Or (TypeFlag = "puoA" And S >= H) _
        Or (TypeFlag = "cdoA" And S <= H) Or (TypeFlag = "pdoA" And S <= H) Then
            EPartialTimeBarrier = 0
        Exit Function
    End If
      
    If OutPutFlag = "p" Then 'Value
            EPartialTimeBarrier = PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
            EPartialTimeBarrier = (PartialTimeBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v) _
                         - PartialTimeBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "dddv" Then 'DeltaDVol
            EPartialTimeBarrier = 1 / (4 * dS * 0.01) * (PartialTimeBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v + 0.01) - PartialTimeBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v - 0.01) _
        - PartialTimeBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v + 0.01) + PartialTimeBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v - 0.01)) / 100
     ElseIf OutPutFlag = "g" Then 'Gamma
            EPartialTimeBarrier = (PartialTimeBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v) _
            - 2 * PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v) _
            + PartialTimeBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v)) / (dS ^ 2)
    ElseIf OutPutFlag = "gp" Then ' GammaP
            EPartialTimeBarrier = S / 100 * EPartialTimeBarrier("g", TypeFlag, S + dS, X, H, t1, T2, r, b, v)
    ElseIf OutPutFlag = "gv" Then 'DGammaDvol
            EPartialTimeBarrier = (PartialTimeBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v + 0.01) - 2 * PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v + 0.01) + PartialTimeBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v + 0.01) _
                - PartialTimeBarrier(TypeFlag, S + dS, X, H, t1, T2, r, b, v - 0.01) + 2 * PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v - 0.01) - PartialTimeBarrier(TypeFlag, S - dS, X, H, t1, T2, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "v" Then ' Vega
            EPartialTimeBarrier = (PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v + 0.01) _
                    - PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vp" Then ' VegaP
            EPartialTimeBarrier = v / 0.1 * EPartialTimeBarrier("v", TypeFlag, S + dS, X, H, t1, T2, r, b, v)
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol/vomma
            EPartialTimeBarrier = (PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v + 0.01) - 2 * PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v) _
            + PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "r" Then 'Rho
            EPartialTimeBarrier = (PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r + 0.01, b + 0.01, v) _
            - PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r - 0.01, b - 0.01, v)) / 2
      ElseIf OutPutFlag = "fr" Then 'Futures option Rho
            EPartialTimeBarrier = (PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r + 0.01, 0, v) _
            - PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r - 0.01, 0, v)) / 2
    ElseIf OutPutFlag = "f" Then 'Rho2 Phi
            EPartialTimeBarrier = (PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b - 0.01, v) _
            - PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b + 0.01, v)) / 2
    ElseIf OutPutFlag = "b" Then ' Carry sensitivity
            EPartialTimeBarrier = (PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b + 0.01, v) _
            - PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b - 0.01, v)) / 2
      ElseIf OutPutFlag = "t" Then 'Theta
            If t1 <= 1 / 365 Then
                EPartialTimeBarrier = PartialTimeBarrier(TypeFlag, S, X, H, 0.00001, T2 - 1 / 365, r, b, v) _
                    - PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v)
              Else
                    EPartialTimeBarrier = PartialTimeBarrier(TypeFlag, S, X, H, t1 - 1 / 365, T2 - 1 / 365, r, b, v) _
                    - PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v)
            End If
     ElseIf OutPutFlag = "dx" Then 'Strike Delta
            EPartialTimeBarrier = (PartialTimeBarrier(TypeFlag, S, X + dS, H, t1, T2, r, b, v) _
                - PartialTimeBarrier(TypeFlag, S, X - dS, H, t1, T2, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
      
            EPartialTimeBarrier = (PartialTimeBarrier(TypeFlag, S, X + dS, H, t1, T2, r, b, v) _
                - 2 * PartialTimeBarrier(TypeFlag, S, X, H, t1, T2, r, b, v) _
                + PartialTimeBarrier(TypeFlag, S, X - dS, H, t1, T2, r, b, v)) / (dS ^ 2)
    End If
    
End Function

'// Partial-time singel asset barrier options
Public Function PartialTimeBarrier(TypeFlag As String, S As Double, X As Double, H As Double, _
                t1 As Double, T2 As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    Dim f1 As Double, f2 As Double
    Dim e1 As Double, e2 As Double
    Dim e3 As Double, e4 As Double
    Dim g1 As Double, g2 As Double
    Dim g3 As Double, g4 As Double
    Dim mu As Double, rho As Double, eta As Integer
    Dim z1 As Double, z2 As Double, z3 As Double
    Dim z4 As Double, z5 As Double, z6 As Double
    Dim z7 As Double, z8 As Double
    
    If TypeFlag = "cdoA" Then
        eta = 1
    ElseIf TypeFlag = "cuoA" Then
        eta = -1
    End If
    
    d1 = (Log(S / X) + (b + v ^ 2 / 2) * T2) / (v * Sqr(T2))
    d2 = d1 - v * Sqr(T2)
    f1 = (Log(S / X) + 2 * Log(H / S) + (b + v ^ 2 / 2) * T2) / (v * Sqr(T2))
    f2 = f1 - v * Sqr(T2)
    e1 = (Log(S / H) + (b + v ^ 2 / 2) * t1) / (v * Sqr(t1))
    e2 = e1 - v * Sqr(t1)
    e3 = e1 + 2 * Log(H / S) / (v * Sqr(t1))
    e4 = e3 - v * Sqr(t1)
    mu = (b - v ^ 2 / 2) / v ^ 2
    rho = Sqr(t1 / T2)
    g1 = (Log(S / H) + (b + v ^ 2 / 2) * T2) / (v * Sqr(T2))
    g2 = g1 - v * Sqr(T2)
    g3 = g1 + 2 * Log(H / S) / (v * Sqr(T2))
    g4 = g3 - v * Sqr(T2)
    
    z1 = CND(e2) - (H / S) ^ (2 * mu) * CND(e4)
    z2 = CND(-e2) - (H / S) ^ (2 * mu) * CND(-e4)
    z3 = CBND(g2, e2, rho) - (H / S) ^ (2 * mu) * CBND(g4, -e4, -rho)
    z4 = CBND(-g2, -e2, rho) - (H / S) ^ (2 * mu) * CBND(-g4, e4, -rho)
    z5 = CND(e1) - (H / S) ^ (2 * (mu + 1)) * CND(e3)
    z6 = CND(-e1) - (H / S) ^ (2 * (mu + 1)) * CND(-e3)
    z7 = CBND(g1, e1, rho) - (H / S) ^ (2 * (mu + 1)) * CBND(g3, -e3, -rho)
    z8 = CBND(-g1, -e1, rho) - (H / S) ^ (2 * (mu + 1)) * CBND(-g3, e3, -rho)
    
    If TypeFlag = "cdoA" Or TypeFlag = "cuoA" Then '// call down-and out and up-and-out type A
        PartialTimeBarrier = S * Exp((b - r) * T2) * (CBND(d1, eta * e1, eta * rho) - (H / S) ^ (2 * (mu + 1)) * CBND(f1, eta * e3, eta * rho)) _
        - X * Exp(-r * T2) * (CBND(d2, eta * e2, eta * rho) - (H / S) ^ (2 * mu) * CBND(f2, eta * e4, eta * rho))
    ElseIf TypeFlag = "cdoB2" And X < H Then  '// call down-and-out type B2
        PartialTimeBarrier = S * Exp((b - r) * T2) * (CBND(g1, e1, rho) - (H / S) ^ (2 * (mu + 1)) * CBND(g3, -e3, -rho)) _
        - X * Exp(-r * T2) * (CBND(g2, e2, rho) - (H / S) ^ (2 * mu) * CBND(g4, -e4, -rho))
    ElseIf TypeFlag = "cdoB2" And X > H Then
        PartialTimeBarrier = PartialTimeBarrier("coB1", S, X, H, t1, T2, r, b, v)
    ElseIf TypeFlag = "cuoB2" And X < H Then  '// call up-and-out type B2
        PartialTimeBarrier = S * Exp((b - r) * T2) * (CBND(-g1, -e1, rho) - (H / S) ^ (2 * (mu + 1)) * CBND(-g3, e3, -rho)) _
        - X * Exp(-r * T2) * (CBND(-g2, -e2, rho) - (H / S) ^ (2 * mu) * CBND(-g4, e4, -rho)) _
        - S * Exp((b - r) * T2) * (CBND(-d1, -e1, rho) - (H / S) ^ (2 * (mu + 1)) * CBND(e3, -f1, -rho)) _
        + X * Exp(-r * T2) * (CBND(-d2, -e2, rho) - (H / S) ^ (2 * mu) * CBND(e4, -f2, -rho))
    ElseIf TypeFlag = "coB1" And X > H Then  '// call out type B1
        PartialTimeBarrier = S * Exp((b - r) * T2) * (CBND(d1, e1, rho) - (H / S) ^ (2 * (mu + 1)) * CBND(f1, -e3, -rho)) _
        - X * Exp(-r * T2) * (CBND(d2, e2, rho) - (H / S) ^ (2 * mu) * CBND(f2, -e4, -rho))
    ElseIf TypeFlag = "coB1" And X < H Then
        PartialTimeBarrier = S * Exp((b - r) * T2) * (CBND(-g1, -e1, rho) - (H / S) ^ (2 * (mu + 1)) * CBND(-g3, e3, -rho)) _
        - X * Exp(-r * T2) * (CBND(-g2, -e2, rho) - (H / S) ^ (2 * mu) * CBND(-g4, e4, -rho)) _
        - S * Exp((b - r) * T2) * (CBND(-d1, -e1, rho) - (H / S) ^ (2 * (mu + 1)) * CBND(-f1, e3, -rho)) _
        + X * Exp(-r * T2) * (CBND(-d2, -e2, rho) - (H / S) ^ (2 * mu) * CBND(-f2, e4, -rho)) _
        + S * Exp((b - r) * T2) * (CBND(g1, e1, rho) - (H / S) ^ (2 * (mu + 1)) * CBND(g3, -e3, -rho)) _
        - X * Exp(-r * T2) * (CBND(g2, e2, rho) - (H / S) ^ (2 * mu) * CBND(g4, -e4, -rho))
    ElseIf TypeFlag = "pdoA" Then  '// put down-and out and up-and-out type A
        PartialTimeBarrier = PartialTimeBarrier("cdoA", S, X, H, t1, T2, r, b, v) - S * Exp((b - r) * T2) * z5 + X * Exp(-r * T2) * z1
    ElseIf TypeFlag = "puoA" Then
        PartialTimeBarrier = PartialTimeBarrier("cuoA", S, X, H, t1, T2, r, b, v) - S * Exp((b - r) * T2) * z6 + X * Exp(-r * T2) * z2
    ElseIf TypeFlag = "poB1" Then  '// put out type B1
        PartialTimeBarrier = PartialTimeBarrier("coB1", S, X, H, t1, T2, r, b, v) - S * Exp((b - r) * T2) * z8 + X * Exp(-r * T2) * z4 - S * Exp((b - r) * T2) * z7 + X * Exp(-r * T2) * z3
    ElseIf TypeFlag = "pdoB2" Then  '// put down-and-out type B2
        PartialTimeBarrier = PartialTimeBarrier("cdoB2", S, X, H, t1, T2, r, b, v) - S * Exp((b - r) * T2) * z7 + X * Exp(-r * T2) * z3
    ElseIf TypeFlag = "puoB2" Then  '// put up-and-out type B2
        PartialTimeBarrier = PartialTimeBarrier("cuoB2", S, X, H, t1, T2, r, b, v) - S * Exp((b - r) * T2) * z8 + X * Exp(-r * T2) * z4
    End If
    
End Function






'// Double barrier options
Function DoubleBarrier(TypeFlag As String, S As Double, X As Double, L As Double, U As Double, T As Double, _
        r As Double, b As Double, v As Double, delta1 As Double, delta2 As Double) As Double
    
    Dim E As Double, F As Double
    Dim Sum1 As Double, Sum2 As Double
    Dim d1 As Double, d2 As Double
    Dim d3 As Double, d4 As Double
    Dim mu1 As Double, mu2 As Double, mu3 As Double
    Dim OutValue As Double, n As Integer
    
    F = U * Exp(delta1 * T)
    E = L * Exp(delta2 * T)
    Sum1 = 0
    Sum2 = 0
    
    If TypeFlag = "co" Or TypeFlag = "ci" Then
        For n = -5 To 5
            d1 = (Log(S * U ^ (2 * n) / (X * L ^ (2 * n))) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
            d2 = (Log(S * U ^ (2 * n) / (F * L ^ (2 * n))) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
            d3 = (Log(L ^ (2 * n + 2) / (X * S * U ^ (2 * n))) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
            d4 = (Log(L ^ (2 * n + 2) / (F * S * U ^ (2 * n))) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
            mu1 = 2 * (b - delta2 - n * (delta1 - delta2)) / v ^ 2 + 1
            mu2 = 2 * n * (delta1 - delta2) / v ^ 2
            mu3 = 2 * (b - delta2 + n * (delta1 - delta2)) / v ^ 2 + 1
            Sum1 = Sum1 + (U ^ n / L ^ n) ^ mu1 * (L / S) ^ mu2 * (CND(d1) - CND(d2)) - (L ^ (n + 1) / (U ^ n * S)) ^ mu3 * (CND(d3) - CND(d4))
            Sum2 = Sum2 + (U ^ n / L ^ n) ^ (mu1 - 2) * (L / S) ^ mu2 * (CND(d1 - v * Sqr(T)) - CND(d2 - v * Sqr(T))) - (L ^ (n + 1) / (U ^ n * S)) ^ (mu3 - 2) * (CND(d3 - v * Sqr(T)) - CND(d4 - v * Sqr(T)))
        Next
        OutValue = S * Exp((b - r) * T) * Sum1 - X * Exp(-r * T) * Sum2
    ElseIf TypeFlag = "po" Or TypeFlag = "pi" Then
        For n = -5 To 5
            d1 = (Log(S * U ^ (2 * n) / (E * L ^ (2 * n))) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
            d2 = (Log(S * U ^ (2 * n) / (X * L ^ (2 * n))) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
            d3 = (Log(L ^ (2 * n + 2) / (E * S * U ^ (2 * n))) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
            d4 = (Log(L ^ (2 * n + 2) / (X * S * U ^ (2 * n))) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
            mu1 = 2 * (b - delta2 - n * (delta1 - delta2)) / v ^ 2 + 1
            mu2 = 2 * n * (delta1 - delta2) / v ^ 2
            mu3 = 2 * (b - delta2 + n * (delta1 - delta2)) / v ^ 2 + 1
            Sum1 = Sum1 + (U ^ n / L ^ n) ^ mu1 * (L / S) ^ mu2 * (CND(d1) - CND(d2)) - (L ^ (n + 1) / (U ^ n * S)) ^ mu3 * (CND(d3) - CND(d4))
            Sum2 = Sum2 + (U ^ n / L ^ n) ^ (mu1 - 2) * (L / S) ^ mu2 * (CND(d1 - v * Sqr(T)) - CND(d2 - v * Sqr(T))) - (L ^ (n + 1) / (U ^ n * S)) ^ (mu3 - 2) * (CND(d3 - v * Sqr(T)) - CND(d4 - v * Sqr(T)))
        Next
        OutValue = X * Exp(-r * T) * Sum2 - S * Exp((b - r) * T) * Sum1
    End If
    If TypeFlag = "co" Or TypeFlag = "po" Then
        DoubleBarrier = OutValue
    ElseIf TypeFlag = "ci" Then
        DoubleBarrier = GBlackScholes("c", S, X, T, r, b, v) - OutValue
    ElseIf TypeFlag = "pi" Then
        DoubleBarrier = GBlackScholes("p", S, X, T, r, b, v) - OutValue
    End If
End Function



'// Standard barrier options
Public Function StandardBarrier(TypeFlag As String, S As Double, X As Double, H As Double, k As Double, T As Double, _
            r As Double, b As Double, v As Double)

    'TypeFlag:      The "TypeFlag" gives you 8 different standard barrier options:
    '               1) "cdi"=Down-and-in call,    2) "cui"=Up-and-in call
    '               3) "pdi"=Down-and-in put,     4) "pui"=Up-and-in put
    '               5) "cdo"=Down-and-out call,   6) "cuo"=Up-out-in call
    '               7) "pdo"=Down-and-out put,    8) "puo"=Up-out-in put
    
    Dim mu As Double
    Dim lambda As Double
    Dim X1 As Double, X2 As Double
    Dim y1 As Double, y2 As Double
    Dim z As Double
    
    Dim eta As Integer    'Binary variable that can take the value of 1 or -1
    Dim phi As Integer    'Binary variable that can take the value of 1 or -1
    
    Dim f1 As Double    'Equal to formula "A" in the book
    Dim f2 As Double    'Equal to formula "B" in the book
    Dim f3 As Double    'Equal to formula "C" in the book
    Dim f4 As Double    'Equal to formula "D" in the book
    Dim f5 As Double    'Equal to formula "E" in the book
    Dim f6 As Double    'Equal to formula "F" in the book

    mu = (b - v ^ 2 / 2) / v ^ 2
    lambda = Sqr(mu ^ 2 + 2 * r / v ^ 2)
    X1 = Log(S / X) / (v * Sqr(T)) + (1 + mu) * v * Sqr(T)
    X2 = Log(S / H) / (v * Sqr(T)) + (1 + mu) * v * Sqr(T)
    y1 = Log(H ^ 2 / (S * X)) / (v * Sqr(T)) + (1 + mu) * v * Sqr(T)
    y2 = Log(H / S) / (v * Sqr(T)) + (1 + mu) * v * Sqr(T)
    z = Log(H / S) / (v * Sqr(T)) + lambda * v * Sqr(T)
    
    If TypeFlag = "cdi" Or TypeFlag = "cdo" Then
        eta = 1
        phi = 1
    ElseIf TypeFlag = "cui" Or TypeFlag = "cuo" Then
        eta = -1
        phi = 1
    ElseIf TypeFlag = "pdi" Or TypeFlag = "pdo" Then
        eta = 1
        phi = -1
    ElseIf TypeFlag = "pui" Or TypeFlag = "puo" Then
        eta = -1
        phi = -1
    End If
    
    f1 = phi * S * Exp((b - r) * T) * CND(phi * X1) - phi * X * Exp(-r * T) * CND(phi * X1 - phi * v * Sqr(T))
    f2 = phi * S * Exp((b - r) * T) * CND(phi * X2) - phi * X * Exp(-r * T) * CND(phi * X2 - phi * v * Sqr(T))
    f3 = phi * S * Exp((b - r) * T) * (H / S) ^ (2 * (mu + 1)) * CND(eta * y1) - phi * X * Exp(-r * T) * (H / S) ^ (2 * mu) * CND(eta * y1 - eta * v * Sqr(T))
    f4 = phi * S * Exp((b - r) * T) * (H / S) ^ (2 * (mu + 1)) * CND(eta * y2) - phi * X * Exp(-r * T) * (H / S) ^ (2 * mu) * CND(eta * y2 - eta * v * Sqr(T))
    f5 = k * Exp(-r * T) * (CND(eta * X2 - eta * v * Sqr(T)) - (H / S) ^ (2 * mu) * CND(eta * y2 - eta * v * Sqr(T)))
    f6 = k * ((H / S) ^ (mu + lambda) * CND(eta * z) + (H / S) ^ (mu - lambda) * CND(eta * z - 2 * eta * lambda * v * Sqr(T)))
    
    
    If X > H Then
        Select Case TypeFlag
            Case Is = "cdi"      '1a) cdi
                StandardBarrier = f3 + f5
            Case Is = "cui"   '2a) cui
                StandardBarrier = f1 + f5
            Case Is = "pdi"    '3a) pdi
                StandardBarrier = f2 - f3 + f4 + f5
            Case Is = "pui" '4a) pui
                StandardBarrier = f1 - f2 + f4 + f5
            Case Is = "cdo"    '5a) cdo
                StandardBarrier = f1 - f3 + f6
            Case Is = "cuo"   '6a) cuo
                StandardBarrier = f6
            Case Is = "pdo"   '7a) pdo
                StandardBarrier = f1 - f2 + f3 - f4 + f6
            Case Is = "puo" '8a) puo
                StandardBarrier = f2 - f4 + f6
            End Select
    ElseIf X < H Then
        Select Case TypeFlag
            Case Is = "cdi" '1b) cdi
                StandardBarrier = f1 - f2 + f4 + f5
            Case Is = "cui"  '2b) cui
                StandardBarrier = f2 - f3 + f4 + f5
            Case Is = "pdi" '3b) pdi
                StandardBarrier = f1 + f5
            Case Is = "pui"   '4b) pui
                StandardBarrier = f3 + f5
            Case Is = "cdo" '5b) cdo
                StandardBarrier = f2 + f6 - f4
            Case Is = "cuo" '6b) cuo
                StandardBarrier = f1 - f2 + f3 - f4 + f6
            Case Is = "pdo"   '7b) pdo
                StandardBarrier = f6
            Case Is = "puo"  '8b) puo
                StandardBarrier = f1 - f3 + f6
        End Select
    End If
End Function



Public Function EDoubleBarrier(OutPutFlag As String, TypeFlag As String, S As Double, X As Double, L As Double, U As Double, T As Double, _
            r As Double, b As Double, v As Double, delta1 As Double, delta2 As Double, Optional dS)

    If IsMissing(dS) Then
        dS = 0.0001
    End If
  
    
    Dim OutInnFlag As String
    Dim CallPutFlag As String
    
    OutInnFlag = Right(TypeFlag, 1)
    CallPutFlag = Left(TypeFlag, 1)
    
    If OutInnFlag = "o" And (S <= L Or S >= U) Then
            EDoubleBarrier = 0
        Exit Function
        ElseIf OutInnFlag = "i" And (S <= L Or S >= U) Then
            EDoubleBarrier = EGBlackScholes(OutPutFlag, CallPutFlag, S, X, T, r, b, v)
            Exit Function
    End If
    
    If OutPutFlag = "p" Then 'Value
            EDoubleBarrier = DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v, delta1, delta2)
    ElseIf OutPutFlag = "d" Then ' Delta
            EDoubleBarrier = (DoubleBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v, delta1, delta2) _
                 - DoubleBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v, delta1, delta2)) / (2 * dS)
     ElseIf OutPutFlag = "dddv" Then ' DeltaDVol
                 EDoubleBarrier = 1 / (4 * dS * 0.01) * (DoubleBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v + 0.01, delta1, delta2) - DoubleBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v - 0.01, delta1, delta2) _
                - DoubleBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v + 0.01, delta1, delta2) + DoubleBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v - 0.01, delta1, delta2)) / 100
     ElseIf OutPutFlag = "g" Then ' Gamma
            EDoubleBarrier = (DoubleBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v, delta1, delta2) _
            - 2 * DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v, delta1, delta2) _
            + DoubleBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v, delta1, delta2)) / (dS ^ 2)
      ElseIf OutPutFlag = "gp" Then ' GammaP
            EDoubleBarrier = S / 100 * EDoubleBarrier("g", TypeFlag, S + dS, X, L, U, T, r, b, v, delta1, delta2)
    ElseIf OutPutFlag = "gv" Then ' DGammaDVol
             EDoubleBarrier = (DoubleBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v + 0.01, delta1, delta2) - 2 * DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v + 0.01, delta1, delta2) + DoubleBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v + 0.01, delta1, delta2) _
                - DoubleBarrier(TypeFlag, S + dS, X, L, U, T, r, b, v - 0.01, delta1, delta2) + 2 * DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v - 0.01, delta1, delta2) - DoubleBarrier(TypeFlag, S - dS, X, L, U, T, r, b, v - 0.01, delta1, delta2)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "v" Then ' Vega
            EDoubleBarrier = (DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v + 0.01, delta1, delta2) _
                - DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v - 0.01, delta1, delta2)) / 2
    ElseIf OutPutFlag = "dvdv" Then ' DVegaDVol/Vomma
                  EDoubleBarrier = (DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v + 0.01, delta1, delta2) - 2 * DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v, delta1, delta2) _
                        + DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v - 0.01, delta1, delta2)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then ' VegaP
            EDoubleBarrier = v / 0.1 * EDoubleBarrier("v", TypeFlag, S + dS, X, L, U, T, r, b, v, delta1, delta2)
    ElseIf OutPutFlag = "r" Then ' Rho
            EDoubleBarrier = (DoubleBarrier(TypeFlag, S, X, L, U, T, r + 0.01, b + 0.01, v, delta1, delta2) _
            - DoubleBarrier(TypeFlag, S, X, L, U, T, r - 0.01, b - 0.01, v, delta1, delta2)) / 2
    ElseIf OutPutFlag = "fr" Then ' Futures option Rho
            EDoubleBarrier = (DoubleBarrier(TypeFlag, S, X, L, U, T, r + 0.01, 0, v, delta1, delta2) _
            - DoubleBarrier(TypeFlag, S, X, L, U, T, r - 0.01, 0, v, delta1, delta2)) / 2
    ElseIf OutPutFlag = "f" Then 'Rho2/Phi
            EDoubleBarrier = (DoubleBarrier(TypeFlag, S, X, L, U, T, r, b - 0.01, v, delta1, delta2) _
            - DoubleBarrier(TypeFlag, S, X, L, U, T, r, b + 0.01, v, delta1, delta2)) / 2
    ElseIf OutPutFlag = "b" Then ' Carry sensitivity
            EDoubleBarrier = (DoubleBarrier(TypeFlag, S, X, L, U, T, r, b + 0.01, v, delta1, delta2) _
            - DoubleBarrier(TypeFlag, S, X, L, U, T, r, b - 0.01, v, delta1, delta2)) / 2
    ElseIf OutPutFlag = "t" Then 'Theta
            If T <= 1 / 365 Then
                EDoubleBarrier = DoubleBarrier(TypeFlag, S, X, L, U, 0.00001, r, b, v, delta1, delta2) _
                    - DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v, delta1, delta2)
              Else
                EDoubleBarrier = DoubleBarrier(TypeFlag, S, X, L, U, T - 1 / 365, r, b, v, delta1, delta2) _
                    - DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v, delta1, delta2)
                End If
    ElseIf OutPutFlag = "dx" Then 'Strike Delta
            EDoubleBarrier = (DoubleBarrier(TypeFlag, S, X + dS, L, U, T, r, b, v, delta1, delta2) _
            - DoubleBarrier(TypeFlag, S, X - dS, L, U, T, r, b, v, delta1, delta2)) / (2 * dS)
    ElseIf OutPutFlag = "dxdx" Then ' Strike Gamma
            EDoubleBarrier = (DoubleBarrier(TypeFlag, S, X + dS, L, U, T, r, b, v, delta1, delta2) _
            - 2 * DoubleBarrier(TypeFlag, S, X, L, U, T, r, b, v, delta1, delta2) _
            + DoubleBarrier(TypeFlag, S, X - dS, L, U, T, r, b, v, delta1, delta2)) / (dS ^ 2)
    End If
End Function


Public Function EStandardBarrier(OutPutFlag As String, TypeFlag As String, S As Double, X As Double, H As Double, k As Double, T As Double, _
            r As Double, b As Double, v As Double, Optional dS)

    If IsMissing(dS) Then
        dS = 0.0001
    End If
    
   
    Dim OutInnFlag As String
    Dim CallPutFlag As String
    
    OutInnFlag = Right(TypeFlag, 2)
    CallPutFlag = Left(TypeFlag, 1)
    
     
    If (OutInnFlag = "do" And S <= H) Or (OutInnFlag = "uo" And S >= H) Then
        If OutPutFlag = "p" Then
            EStandardBarrier = k
        Else
            EStandardBarrier = 0
        End If
        Exit Function
    ElseIf (OutInnFlag = "di" And S <= H) Or (OutInnFlag = "ui" And S >= H) Then
            EStandardBarrier = EGBlackScholes(OutPutFlag, CallPutFlag, S, X, T, r, b, v)
            Exit Function
    End If
      
    
    If OutPutFlag = "p" Then 'Value
            EStandardBarrier = StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
            EStandardBarrier = (StandardBarrier(TypeFlag, S + dS, X, H, k, T, r, b, v) _
                         - StandardBarrier(TypeFlag, S - dS, X, H, k, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "dddv" Then 'DeltaDVol
            EStandardBarrier = 1 / (4 * dS * 0.01) * (StandardBarrier(TypeFlag, S + dS, X, H, k, T, r, b, v + 0.01) - StandardBarrier(TypeFlag, S + dS, X, H, k, T, r, b, v - 0.01) _
        - StandardBarrier(TypeFlag, S - dS, X, H, k, T, r, b, v + 0.01) + StandardBarrier(TypeFlag, S - dS, X, H, k, T, r, b, v - 0.01)) / 100
     ElseIf OutPutFlag = "g" Then 'Gamma
            EStandardBarrier = (StandardBarrier(TypeFlag, S + dS, X, H, k, T, r, b, v) _
            - 2 * StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v) _
            + StandardBarrier(TypeFlag, S - dS, X, H, k, T, r, b, v)) / (dS ^ 2)
    ElseIf OutPutFlag = "gp" Then ' GammaP
            EStandardBarrier = S / 100 * EStandardBarrier("g", TypeFlag, S + dS, X, H, k, T, r, b, v)
    ElseIf OutPutFlag = "gv" Then 'DGammaDvol
            EStandardBarrier = (StandardBarrier(TypeFlag, S + dS, X, H, k, T, r, b, v + 0.01) - 2 * StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v + 0.01) + StandardBarrier(TypeFlag, S - dS, X, H, k, T, r, b, v + 0.01) _
                - StandardBarrier(TypeFlag, S + dS, X, H, k, T, r, b, v - 0.01) + 2 * StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v - 0.01) - StandardBarrier(TypeFlag, S - dS, X, H, k, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "v" Then ' Vega
            EStandardBarrier = (StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v + 0.01) _
                    - StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vp" Then ' VegaP
            EStandardBarrier = v / 0.1 * EStandardBarrier("v", TypeFlag, S + dS, X, H, k, T, r, b, v)
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol/vomma
            EStandardBarrier = (StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v + 0.01) - 2 * StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v) _
            + StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "r" Then 'Rho
            EStandardBarrier = (StandardBarrier(TypeFlag, S, X, H, k, T, r + 0.01, b + 0.01, v) _
            - StandardBarrier(TypeFlag, S, X, H, k, T, r - 0.01, b - 0.01, v)) / 2
      ElseIf OutPutFlag = "fr" Then 'Futures option Rho
            EStandardBarrier = (StandardBarrier(TypeFlag, S, X, H, k, T, r + 0.01, 0, v) _
            - StandardBarrier(TypeFlag, S, X, H, k, T, r - 0.01, 0, v)) / 2
    ElseIf OutPutFlag = "f" Then 'Rho2 Phi
            EStandardBarrier = (StandardBarrier(TypeFlag, S, X, H, k, T, r, b - 0.01, v) _
            - StandardBarrier(TypeFlag, S, X, H, k, T, r, b + 0.01, v)) / 2
    ElseIf OutPutFlag = "b" Then ' Carry sensitivity
            EStandardBarrier = (StandardBarrier(TypeFlag, S, X, H, k, T, r, b + 0.01, v) _
            - StandardBarrier(TypeFlag, S, X, H, k, T, r, b - 0.01, v)) / 2
      ElseIf OutPutFlag = "t" Then 'Theta
            If T <= 1 / 365 Then
                EStandardBarrier = StandardBarrier(TypeFlag, S, X, H, k, 0.00001, r, b, v) _
                    - StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v)
              Else
                    EStandardBarrier = StandardBarrier(TypeFlag, S, X, H, k, T - 1 / 365, r, b, v) _
                    - StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v)
            End If
     ElseIf OutPutFlag = "dx" Then 'Strike Delta
            EStandardBarrier = (StandardBarrier(TypeFlag, S, X + dS, H, k, T, r, b, v) _
                - StandardBarrier(TypeFlag, S, X - dS, H, k, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
      
            EStandardBarrier = (StandardBarrier(TypeFlag, S, X + dS, H, k, T, r, b, v) _
                - 2 * StandardBarrier(TypeFlag, S, X, H, k, T, r, b, v) _
                + StandardBarrier(TypeFlag, S, X - dS, H, k, T, r, b, v)) / (dS ^ 2)
    End If
    
End Function

    

'// Discrete barrier monitoring adjustment
Public Function DiscreteAdjustedBarrier(S As Double, H As Double, v As Double, dt As Double) As Double

    If H > S Then
        DiscreteAdjustedBarrier = H * Exp(0.5826 * v * Sqr(dt))
    ElseIf H < S Then
        DiscreteAdjustedBarrier = H * Exp(-0.5826 * v * Sqr(dt))
    End If
End Function
