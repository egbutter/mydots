Attribute VB_Name = "StochasticVol"
Option Explicit


Public Function HullWhite87SV(CallPutFlag As String, S As Double, X As Double, T As Double, _
    r As Double, b As Double, v As Double, Vvol As Double) As Double

    '// v: intitial volatility/standard deviation
    '// VVol: volatility of volatility
    '// rho: correlation between asset price and volatility
    
    Dim d1 As Double, d2 As Double, k As Double
    Dim CallValue As Double, cgbs As Double, ek As Double
    Dim cVV As Double, cVVV As Double
    
    k = Vvol ^ 2 * T
    ek = Exp(k)
    
    d1 = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    cgbs = GBlackScholes("c", S, X, T, r, b, v)
    
    '//Partial derivatives
    cVV = S * Exp((b - r) * T) * Sqr(T) * ND(d1) * (d1 * d2 - 1) / (4 * v ^ 3)
    cVVV = S * Exp((b - r) * T) * Sqr(T) * ND(d1) * ((d1 * d2 - 1) _
        * (d1 * d2 - 3) - (d1 ^ 2 + d2 ^ 2)) / (8 * v ^ 5)

    CallValue = cgbs + 1 / 2 * cVV * (2 * v ^ 4 * (ek - k - 1) / k ^ 2 - v ^ 4) _
            + 1 / 6 * cVVV * v ^ 6 * (ek ^ 3 - (9 + 18 * k) * ek + 8 _
            + 24 * k + 18 * k ^ 2 + 6 * k ^ 3) / (3 * k ^ 3)
  
    If CallPutFlag = "c" Then
        HullWhite87SV = CallValue
    ElseIf CallPutFlag = "p" Then '// Use put call parity to find put
        HullWhite87SV = CallValue - S * Exp((b - r) * T) + X * Exp(-r * T)
    End If

End Function


Public Function EHullWhite87SV(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, Vvol As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EHullWhite87SV = HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol)
    ElseIf OutPutFlag = "d" Then 'Delta
         EHullWhite87SV = (HullWhite87SV(CallPutFlag, S + dS, X, T, r, b, v, Vvol) - HullWhite87SV(CallPutFlag, S - dS, X, T, r, b, v, Vvol)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EHullWhite87SV = (HullWhite87SV(CallPutFlag, S + dS, X, T, r, b, v, Vvol) - HullWhite87SV(CallPutFlag, S - dS, X, T, r, b, v, Vvol)) / (2 * dS) * S / HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EHullWhite87SV = (HullWhite87SV(CallPutFlag, S + dS, X, T, r, b, v, Vvol) - 2 * HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol) + HullWhite87SV(CallPutFlag, S - dS, X, T, r, b, v, Vvol)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EHullWhite87SV = (HullWhite87SV(CallPutFlag, S + dS, X, T, r, b, v + 0.01, Vvol) - 2 * HullWhite87SV(CallPutFlag, S, X, T, r, b, v + 0.01, Vvol) + HullWhite87SV(CallPutFlag, S - dS, X, T, r, b, v + 0.01, Vvol) _
      - HullWhite87SV(CallPutFlag, S + dS, X, T, r, b, v - 0.01, Vvol) + 2 * HullWhite87SV(CallPutFlag, S, X, T, r, b, v - 0.01, Vvol) - HullWhite87SV(CallPutFlag, S - dS, X, T, r, b, v - 0.01, Vvol)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EHullWhite87SV = S / 100 * (HullWhite87SV(CallPutFlag, S + dS, X, T, r, b, v, Vvol) - 2 * HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol) + HullWhite87SV(CallPutFlag, S - dS, X, T, r, b, v, Vvol)) / dS ^ 2
    ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EHullWhite87SV = 1 / (4 * dS * 0.01) * (HullWhite87SV(CallPutFlag, S + dS, X, T, r, b, v + 0.01, Vvol) - HullWhite87SV(CallPutFlag, S + dS, X, T, r, b, v - 0.01, Vvol) _
        - HullWhite87SV(CallPutFlag, S - dS, X, T, r, b, v + 0.01, Vvol) + HullWhite87SV(CallPutFlag, S - dS, X, T, r, b, v - 0.01, Vvol)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EHullWhite87SV = (HullWhite87SV(CallPutFlag, S, X, T, r, b, v + 0.01, Vvol) - HullWhite87SV(CallPutFlag, S, X, T, r, b, v - 0.01, Vvol)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EHullWhite87SV = (HullWhite87SV(CallPutFlag, S, X, T, r, b, v + 0.01, Vvol) - 2 * HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol) + HullWhite87SV(CallPutFlag, S, X, T, r, b, v - 0.01, Vvol)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EHullWhite87SV = v / 0.1 * (HullWhite87SV(CallPutFlag, S, X, T, r, b, v + 0.01, Vvol) - HullWhite87SV(CallPutFlag, S, X, T, r, b, v - 0.01, Vvol)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EHullWhite87SV = (HullWhite87SV(CallPutFlag, S, X, T, r, b, v + 0.01, Vvol) - 2 * HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol) + HullWhite87SV(CallPutFlag, S, X, T, r, b, v - 0.01, Vvol))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EHullWhite87SV = HullWhite87SV(CallPutFlag, S, X, 0.00001, r, b, v, Vvol) - HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol)
        Else
                EHullWhite87SV = HullWhite87SV(CallPutFlag, S, X, T - 1 / 365, r, b, v, Vvol) - HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EHullWhite87SV = (HullWhite87SV(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v, Vvol) - HullWhite87SV(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v, Vvol)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EHullWhite87SV = (HullWhite87SV(CallPutFlag, S, X, T, r + 0.01, b, v, Vvol) - HullWhite87SV(CallPutFlag, S, X, T, r - 0.01, b, v, Vvol)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EHullWhite87SV = (HullWhite87SV(CallPutFlag, S, X, T, r, b - 0.01, v, Vvol) - HullWhite87SV(CallPutFlag, S, X, T, r, b + 0.01, v, Vvol)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EHullWhite87SV = (HullWhite87SV(CallPutFlag, S, X, T, r, b + 0.01, v, Vvol) - HullWhite87SV(CallPutFlag, S, X, T, r, b - 0.01, v, Vvol)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EHullWhite87SV = 1 / dS ^ 3 * (HullWhite87SV(CallPutFlag, S + 2 * dS, X, T, r, b, v, Vvol) - 3 * HullWhite87SV(CallPutFlag, S + dS, X, T, r, b, v, Vvol) _
                                + 3 * HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol) - HullWhite87SV(CallPutFlag, S - dS, X, T, r, b, v, Vvol))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EHullWhite87SV = (HullWhite87SV(CallPutFlag, S, X + dS, T, r, b, v, Vvol) - HullWhite87SV(CallPutFlag, S, X - dS, T, r, b, v, Vvol)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EHullWhite87SV = (HullWhite87SV(CallPutFlag, S, X + dS, T, r, b, v, Vvol) - 2 * HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol) + HullWhite87SV(CallPutFlag, S, X - dS, T, r, b, v, Vvol)) / dS ^ 2
    ElseIf OutPutFlag = "vvVega" Then 'VolVol Vega
         EHullWhite87SV = (HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol + 0.01) - HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol - 0.01)) / 2
    ElseIf OutPutFlag = "diff" Then 'HullWhite87 minus Black-Scholes-Merton
         EHullWhite87SV = HullWhite87SV(CallPutFlag, S, X, T, r, b, v, Vvol) - GBlackScholes(CallPutFlag, S, X, T, r, b, v)
   
    End If

End Function


Public Function HullWhite88SV(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, _
            sig0 As Double, sigLR As Double, HL As Double, Vvol As Double, rho As Double) As Double

    
    '// sig0: intitial volatility
    '// sigLR: the long run mean reversion level of volatility
    '// HL: half-life of volatilit deviation
    '// VVol: volatility of volatility
    '// rho: correlation between asset price and volatility
    
    Dim phi1 As Double, phi2 As Double, phi3 As Double, phi4 As Double
    Dim f0 As Double, f1 As Double, f2 As Double, d1 As Double, d2 As Double
    Dim cSV As Double, cVV As Double, cSVV As Double, cVVV As Double, ed As Double
    Dim delta As Double, Beta As Double, A As Double, v As Double, Vbar As Double
    Dim CallValue As Double
    
    
    Beta = -Log(2) / HL '// Find  constant, beta, from Half Life
    A = -Beta * sigLR ^ 2 '// Find constant, a, from long run volatility
    delta = Beta * T
    ed = Exp(delta)
    v = sig0 ^ 2
    
    If Abs(Beta) < 0.0001 Then
        Vbar = v + 0.5 * A * T '// Average expected variance
     Else
        Vbar = (v + A / Beta) * (ed - 1) / delta - A / Beta '// Average expected variance
    End If
    
    d1 = (Log(S / X) + (b + Vbar / 2) * T) / Sqr(Vbar * T)
    d2 = d1 - Sqr(Vbar * T)

    '// Partial derivatives
    cSV = -S * Exp((b - r) * T) * ND(d1) * d2 / (2 * Vbar)
    cVV = S * Exp((b - r) * T) * ND(d1) * Sqr(T) / (4 * Vbar ^ 1.5) * (d1 * d2 - 1)
    cSVV = S * Exp((b - r) * T) / (4 * Vbar ^ 2) * ND(d1) * _
            (-d1 * d2 ^ 2 + d1 + 2 * d2)
    cVVV = S * Exp((b - r) * T) * ND(d1) * Sqr(T) / (8 * Vbar ^ 2.5) _
            * ((d1 * d2 - 1) * (d1 * d2 - 3) - (d1 ^ 2 + d2 ^ 2))

    If Abs(Beta) < 0.0001 Then
        f1 = rho * (A * T / 3 + v) * T / 2 * cSV
        phi1 = rho ^ 2 * (A * T / 4 + v) * T ^ 3 / 6
        phi2 = (2 + 1 / rho ^ 2) * phi1
        phi3 = rho ^ 2 * (A * T / 3 + v) ^ 2 * T ^ 4 / 8
        phi4 = 2 * phi3
    Else '// Beta different from zero
        phi1 = rho ^ 2 / Beta ^ 4 * ((A + Beta * v) * (ed * (delta ^ 2 / 2 - delta + 1) - 1) _
        + A * (ed * (2 - delta) - (2 + delta)))
        phi2 = 2 * phi1 + 1 / (2 * Beta ^ 4) * ((A + Beta * v) * (ed ^ 2 - 2 * delta * ed - 1) _
        - A / 2 * (ed ^ 2 - 4 * ed + 2 * delta + 3))
        phi3 = rho ^ 2 / (2 * Beta ^ 6) * ((A + Beta * v) * (ed - delta * ed - 1) _
        - A * (1 + delta - ed)) ^ 2
        phi4 = 2 * phi3
        f1 = rho / (Beta ^ 3 * T) * ((A + Beta * v) * (1 - ed + delta * ed) _
        + A * (1 + delta - ed)) * cSV
    End If
    
    f0 = S * Exp((b - r) * T) * CND(d1) - X * Exp(-r * T) * CND(d2)
    f2 = phi1 / T * cSV + phi2 / T ^ 2 * cVV + phi3 / T ^ 2 * cSVV + phi4 / T ^ 3 * cVVV

    CallValue = f0 + f1 * Vvol + f2 * Vvol ^ 2
    
    If CallPutFlag = "c" Then
        HullWhite88SV = CallValue
    ElseIf CallPutFlag = "p" Then '// Use put call parity to find put
        HullWhite88SV = CallValue - S * Exp((b - r) * T) + X * Exp(-r * T)
    End If

End Function


Public Function EHullWhite88SV(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, sig0 As Double, sigLR As Double, HL As Double, Vvol As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EHullWhite88SV = HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho)
    ElseIf OutPutFlag = "d" Then 'Delta
         EHullWhite88SV = (HullWhite88SV(CallPutFlag, S + dS, X, T, r, b, sig0, sigLR, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S - dS, X, T, r, b, sig0, sigLR, HL, Vvol, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EHullWhite88SV = (HullWhite88SV(CallPutFlag, S + dS, X, T, r, b, sig0, sigLR, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S - dS, X, T, r, b, sig0, sigLR, HL, Vvol, rho)) / (2 * dS) * S / HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EHullWhite88SV = (HullWhite88SV(CallPutFlag, S + dS, X, T, r, b, sig0, sigLR, HL, Vvol, rho) - 2 * HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho) + HullWhite88SV(CallPutFlag, S - dS, X, T, r, b, sig0, sigLR, HL, Vvol, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EHullWhite88SV = (HullWhite88SV(CallPutFlag, S + dS, X, T, r, b, sig0, sigLR + 0.01, HL, Vvol, rho) - 2 * HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR + 0.01, HL, Vvol, rho) + HullWhite88SV(CallPutFlag, S - dS, X, T, r, b, sig0, sigLR + 0.01, HL, Vvol, rho) _
      - HullWhite88SV(CallPutFlag, S + dS, X, T, r, b, sig0, sigLR - 0.01, HL, Vvol, rho) + 2 * HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR - 0.01, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S - dS, X, T, r, b, sig0, sigLR - 0.01, HL, Vvol, rho)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EHullWhite88SV = S / 100 * (HullWhite88SV(CallPutFlag, S + dS, X, T, r, b, sig0, sigLR, HL, Vvol, rho) - 2 * HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho) + HullWhite88SV(CallPutFlag, S - dS, X, T, r, b, sig0, sigLR, HL, Vvol, rho)) / dS ^ 2
    ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EHullWhite88SV = 1 / (4 * dS * 0.01) * (HullWhite88SV(CallPutFlag, S + dS, X, T, r, b, sig0, sigLR + 0.01, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S + dS, X, T, r, b, sig0, sigLR - 0.01, HL, Vvol, rho) _
        - HullWhite88SV(CallPutFlag, S - dS, X, T, r, b, sig0, sigLR + 0.01, HL, Vvol, rho) + HullWhite88SV(CallPutFlag, S - dS, X, T, r, b, sig0, sigLR - 0.01, HL, Vvol, rho)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR + 0.01, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR - 0.01, HL, Vvol, rho)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR + 0.01, HL, Vvol, rho) - 2 * HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho) + HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR - 0.01, HL, Vvol, rho)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EHullWhite88SV = sigLR / 0.1 * (HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR + 0.01, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR - 0.01, HL, Vvol, rho)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR + 0.01, HL, Vvol, rho) - 2 * HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho) + HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR - 0.01, HL, Vvol, rho))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EHullWhite88SV = HullWhite88SV(CallPutFlag, S, X, 0.00001, r, b, sig0, sigLR, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho)
        Else
                EHullWhite88SV = HullWhite88SV(CallPutFlag, S, X, T - 1 / 365, r, b, sig0, sigLR, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X, T, r + 0.01, b + 0.01, sig0, sigLR, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S, X, T, r - 0.01, b - 0.01, sig0, sigLR, HL, Vvol, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X, T, r + 0.01, b, sig0, sigLR, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S, X, T, r - 0.01, b, sig0, sigLR, HL, Vvol, rho)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X, T, r, b - 0.01, sig0, sigLR, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S, X, T, r, b + 0.01, sig0, sigLR, HL, Vvol, rho)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X, T, r, b + 0.01, sig0, sigLR, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S, X, T, r, b - 0.01, sig0, sigLR, HL, Vvol, rho)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EHullWhite88SV = 1 / dS ^ 3 * (HullWhite88SV(CallPutFlag, S + 2 * dS, X, T, r, b, sig0, sigLR, HL, Vvol, rho) - 3 * HullWhite88SV(CallPutFlag, S + dS, X, T, r, b, sig0, sigLR, HL, Vvol, rho) _
                                + 3 * HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S - dS, X, T, r, b, sig0, sigLR, HL, Vvol, rho))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X + dS, T, r, b, sig0, sigLR, HL, Vvol, rho) - HullWhite88SV(CallPutFlag, S, X - dS, T, r, b, sig0, sigLR, HL, Vvol, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X + dS, T, r, b, sig0, sigLR, HL, Vvol, rho) - 2 * HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho) + HullWhite88SV(CallPutFlag, S, X - dS, T, r, b, sig0, sigLR, HL, Vvol, rho)) / dS ^ 2
    ElseIf OutPutFlag = "vvVega" Then 'VolVol Vega
         EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol + 0.01, rho) - HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol - 0.01, rho)) / 2
      ElseIf OutPutFlag = "corr" Then 'Correlation delta
         EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho + 0.01) - HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho - 0.01)) / 2
  ElseIf OutPutFlag = "corrVomma" Then ' Correlation vomma
        EHullWhite88SV = (HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho + 0.01) - 2 * HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho) + HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "diff" Then 'HullWhite88 minus Black-Scholes-Merton
        Dim kappa As Double, AverageVol As Double
        kappa = Log(2) / HL
        AverageVol = Sqr(((1 - Exp(-kappa * T)) / (kappa * T)) * (sig0 ^ 2 - sigLR ^ 2) + sigLR ^ 2)
        EHullWhite88SV = HullWhite88SV(CallPutFlag, S, X, T, r, b, sig0, sigLR, HL, Vvol, rho) - GBlackScholes(CallPutFlag, S, X, T, r, b, AverageVol)
    End If

End Function
