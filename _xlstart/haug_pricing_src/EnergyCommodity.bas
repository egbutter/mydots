Attribute VB_Name = "EnergyCommodity"
Option Explicit

' Programmmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug
    
'// Energy Swaption
Public Function EnergySwaption(CallPutFlag As String, F As Double, X _
      As Double, T As Double, Tb As Double, rj As Double, rb As Double, j As Integer, n As Integer, v As Double) As Double

    Dim d1 As Double, d2 As Double, Df As Double
    '// T: years to option expiry
    '// Tb: years to start of swap delivery period (Tb>=T)
    '// rb: zero coupon rate from now to start of swap delivery period
    '// rj: swap rate covering delivery period with j compoundings per year
    '// j: number of compoundings per year
    '// n: number of days in delivery period
    
     Df = (1 - 1 / (1 + rj / j) ^ n) / rj * j / n
  
     d1 = (Log(F / X) + v ^ 2 / 2 * T) / (v * Sqr(T))
     d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then
        EnergySwaption = Df * Exp(-rb * Tb) * (F * CND(d1) - X * CND(d2))
    ElseIf CallPutFlag = "p" Then
        EnergySwaption = Df * Exp(-rb * Tb) * (X * CND(-d2) - F * CND(-d1))
    End If
    
End Function


'// Calculates implied energy swap price from energy swaption prices
Public Function ImpliedSwapPrice(c As Double, p As Double, X As Double, Tb As Double, rj As Double, rb As Double, j As Integer, n As Integer) As Double

    Dim Df As Double
    '// Tb: years to start of swap delivery period (Tb>=T)
    '// rb: zero coupon rate from now to start of swap delivery period
    '// rj: swap rate covering delivery period with j compoundings per year
    '// j: number of compoundings per year
    '// n: number of days in delivery period
    
     Df = (1 - 1 / (1 + rj / j) ^ n) / rj * j / n
    ImpliedSwapPrice = (c - p) / (Df * Exp(-rb * Tb)) + X
    
End Function




'// Energy swaption with numerical Greeks
Public Function EEnergySwaption(OutPutFlag As String, CallPutFlag As String, F As Double, X As Double, T As Double, Tb As Double, _
                rj As Double, rb As Double, j As Integer, n As Integer, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        EEnergySwaption = EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EEnergySwaption = (EnergySwaption(CallPutFlag, F + dS, X, T, Tb, rj, rb, j, n, v) - EnergySwaption(CallPutFlag, F - dS, X, T, Tb, rj, rb, j, n, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EEnergySwaption = (EnergySwaption(CallPutFlag, F + dS, X, T, Tb, rj, rb, j, n, v) - EnergySwaption(CallPutFlag, F - dS, X, T, Tb, rj, rb, j, n, v)) / (2 * dS) * F / EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EEnergySwaption = (EnergySwaption(CallPutFlag, F + dS, X, T, Tb, rj, rb, j, n, v) - 2 * EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v) + EnergySwaption(CallPutFlag, F - dS, X, T, Tb, rj, rb, j, n, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EEnergySwaption = (EnergySwaption(CallPutFlag, F + dS, X, T, Tb, rj, rb, j, n, v + 0.01) - 2 * EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v + 0.01) + EnergySwaption(CallPutFlag, F - dS, X, T, Tb, rj, rb, j, n, v + 0.01) _
      - EnergySwaption(CallPutFlag, F + dS, X, T, Tb, rj, rb, j, n, v - 0.01) + 2 * EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v - 0.01) - EnergySwaption(CallPutFlag, F - dS, X, T, Tb, rj, rb, j, n, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "gp" Then 'GammaP
        EEnergySwaption = F / 100 * (EnergySwaption(CallPutFlag, F + dS, X, T, Tb, rj, rb, j, n, v) - 2 * EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v) + EnergySwaption(CallPutFlag, F - dS, X, T, Tb, rj, rb, j, n, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EEnergySwaption = 1 / (4 * dS * 0.01) * (EnergySwaption(CallPutFlag, F + dS, X, T, Tb, rj, rb, j, n, v + 0.01) - EnergySwaption(CallPutFlag, F + dS, X, T, Tb, rj, rb, j, n, v - 0.01) _
        - EnergySwaption(CallPutFlag, F - dS, X, T, Tb, rj, rb, j, n, v + 0.01) + EnergySwaption(CallPutFlag, F - dS, X, T, Tb, rj, rb, j, n, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EEnergySwaption = (EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v + 0.01) - EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EEnergySwaption = (EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v + 0.01) - 2 * EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v) + EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EEnergySwaption = v / 0.1 * (EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v + 0.01) - EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EEnergySwaption = (EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v + 0.01) - 2 * EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v) + EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EEnergySwaption = EnergySwaption(CallPutFlag, F, X, 0.00001, Tb - 1 / 365, rj, rb, j, n, v) - EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v)
        Else
                EEnergySwaption = EnergySwaption(CallPutFlag, F, X, T - 1 / 365, Tb - 1 / 365, rj, rb, j, n, v) - EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v)
        End If
     ElseIf OutPutFlag = "fr" Then 'Futurbs options rho
         EEnergySwaption = (EnergySwaption(CallPutFlag, F, X, T, Tb, rj + 0.01, rb + 0.01, j, n, v) - EnergySwaption(CallPutFlag, F, X, T, Tb, rj + 0.01, rb - 0.01, j, n, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EEnergySwaption = 1 / dS ^ 3 * (EnergySwaption(CallPutFlag, F + 2 * dS, X, T, Tb, rj, rb, j, n, v) - 3 * EnergySwaption(CallPutFlag, F + dS, X, T, Tb, rj, rb, j, n, v) _
                                + 3 * EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v) - EnergySwaption(CallPutFlag, F - dS, X, T, Tb, rj, rb, j, n, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EEnergySwaption = (EnergySwaption(CallPutFlag, F, X + dS, T, Tb, rj, rb, j, n, v) - EnergySwaption(CallPutFlag, F, X - dS, T, Tb, rj, rb, j, n, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EEnergySwaption = (EnergySwaption(CallPutFlag, F, X + dS, T, Tb, rj, rb, j, n, v) - 2 * EnergySwaption(CallPutFlag, F, X, T, Tb, rj, rb, j, n, v) + EnergySwaption(CallPutFlag, F, X - dS, T, Tb, rj, rb, j, n, v)) / dS ^ 2
    End If
End Function


'// Energy swaption approximation
Public Function EnergySwaptionApproximation(CallPutFlag As String, F As Double, X _
                As Double, T As Double, Tm As Double, re As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    '// T: years to option expiry
    '// Tm: years to start of swap delivery period (Tm>=T)
    '// re: zero coupon rate from now to start of swap delivery period
    
    d1 = (Log(F / X) + v ^ 2 / 2 * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then
        EnergySwaptionApproximation = Exp(-re * Tm) * (F * CND(d1) - X * CND(d2))
    ElseIf CallPutFlag = "p" Then
        EnergySwaptionApproximation = Exp(-re * Tm) * (X * CND(-d2) - F * CND(-d1))
    End If
    
End Function


'// Energy swaption approximation with numerical Greeks
Public Function EEnergySwaptionApproximation(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, Tm As Double, _
                re As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        EEnergySwaptionApproximation = EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EEnergySwaptionApproximation = (EnergySwaptionApproximation(CallPutFlag, S + dS, X, T, Tm, re, v) - EnergySwaptionApproximation(CallPutFlag, S - dS, X, T, Tm, re, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EEnergySwaptionApproximation = (EnergySwaptionApproximation(CallPutFlag, S + dS, X, T, Tm, re, v) - EnergySwaptionApproximation(CallPutFlag, S - dS, X, T, Tm, re, v)) / (2 * dS) * S / EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EEnergySwaptionApproximation = (EnergySwaptionApproximation(CallPutFlag, S + dS, X, T, Tm, re, v) - 2 * EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v) + EnergySwaptionApproximation(CallPutFlag, S - dS, X, T, Tm, re, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EEnergySwaptionApproximation = (EnergySwaptionApproximation(CallPutFlag, S + dS, X, T, Tm, re, v + 0.01) - 2 * EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v + 0.01) + EnergySwaptionApproximation(CallPutFlag, S - dS, X, T, Tm, re, v + 0.01) _
      - EnergySwaptionApproximation(CallPutFlag, S + dS, X, T, Tm, re, v - 0.01) + 2 * EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v - 0.01) - EnergySwaptionApproximation(CallPutFlag, S - dS, X, T, Tm, re, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "gp" Then 'GammaP
        EEnergySwaptionApproximation = S / 100 * (EnergySwaptionApproximation(CallPutFlag, S + dS, X, T, Tm, re, v) - 2 * EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v) + EnergySwaptionApproximation(CallPutFlag, S - dS, X, T, Tm, re, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EEnergySwaptionApproximation = 1 / (4 * dS * 0.01) * (EnergySwaptionApproximation(CallPutFlag, S + dS, X, T, Tm, re, v + 0.01) - EnergySwaptionApproximation(CallPutFlag, S + dS, X, T, Tm, re, v - 0.01) _
        - EnergySwaptionApproximation(CallPutFlag, S - dS, X, T, Tm, re, v + 0.01) + EnergySwaptionApproximation(CallPutFlag, S - dS, X, T, Tm, re, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EEnergySwaptionApproximation = (EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v + 0.01) - EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EEnergySwaptionApproximation = (EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v + 0.01) - 2 * EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v) + EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EEnergySwaptionApproximation = v / 0.1 * (EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v + 0.01) - EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EEnergySwaptionApproximation = (EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v + 0.01) - 2 * EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v) + EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EEnergySwaptionApproximation = EnergySwaptionApproximation(CallPutFlag, S, X, 0.00001, Tm - 1 / 365, re, v) - EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v)
        Else
                EEnergySwaptionApproximation = EnergySwaptionApproximation(CallPutFlag, S, X, T - 1 / 365, Tm - 1 / 365, re, v) - EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v)
        End If
     ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EEnergySwaptionApproximation = (EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re + 0.01, v) - EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EEnergySwaptionApproximation = 1 / dS ^ 3 * (EnergySwaptionApproximation(CallPutFlag, S + 2 * dS, X, T, Tm, re, v) - 3 * EnergySwaptionApproximation(CallPutFlag, S + dS, X, T, Tm, re, v) _
                                + 3 * EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v) - EnergySwaptionApproximation(CallPutFlag, S - dS, X, T, Tm, re, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EEnergySwaptionApproximation = (EnergySwaptionApproximation(CallPutFlag, S, X + dS, T, Tm, re, v) - EnergySwaptionApproximation(CallPutFlag, S, X - dS, T, Tm, re, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EEnergySwaptionApproximation = (EnergySwaptionApproximation(CallPutFlag, S, X + dS, T, Tm, re, v) - 2 * EnergySwaptionApproximation(CallPutFlag, S, X, T, Tm, re, v) + EnergySwaptionApproximation(CallPutFlag, S, X - dS, T, Tm, re, v)) / dS ^ 2
    End If
    
End Function


'// Black-Scholes-Merton for options on forwards that expire after option expiry
Public Function BlackScholesForward(CallPutFlag As String, F As Double, X _
                As Double, T As Double, Tf As Double, r As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    '// T: years to option expiry
    '// Tf: years to forward expiry (Tf>=T)
    
    d1 = (Log(F / X) + v ^ 2 / 2 * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then
        BlackScholesForward = Exp(-r * Tf) * (F * CND(d1) - X * CND(d2))
    ElseIf CallPutFlag = "p" Then
        BlackScholesForward = Exp(-r * Tf) * (X * CND(-d2) - F * CND(-d1))
    End If
    
End Function


'// Black-Scholes-Merton for options on forwards with numerical Greeks
Public Function EBlackScholesForward(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, Tf As Double, _
                r As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EBlackScholesForward = BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EBlackScholesForward = (BlackScholesForward(CallPutFlag, S + dS, X, T, Tf, r, v) - BlackScholesForward(CallPutFlag, S - dS, X, T, Tf, r, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EBlackScholesForward = (BlackScholesForward(CallPutFlag, S + dS, X, T, Tf, r, v) - BlackScholesForward(CallPutFlag, S - dS, X, T, Tf, r, v)) / (2 * dS) * S / BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EBlackScholesForward = (BlackScholesForward(CallPutFlag, S + dS, X, T, Tf, r, v) - 2 * BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v) + BlackScholesForward(CallPutFlag, S - dS, X, T, Tf, r, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EBlackScholesForward = (BlackScholesForward(CallPutFlag, S + dS, X, T, Tf, r, v + 0.01) - 2 * BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v + 0.01) + BlackScholesForward(CallPutFlag, S - dS, X, T, Tf, r, v + 0.01) _
      - BlackScholesForward(CallPutFlag, S + dS, X, T, Tf, r, v - 0.01) + 2 * BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v - 0.01) - BlackScholesForward(CallPutFlag, S - dS, X, T, Tf, r, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "gp" Then 'GammaP
        EBlackScholesForward = S / 100 * (BlackScholesForward(CallPutFlag, S + dS, X, T, Tf, r, v) - 2 * BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v) + BlackScholesForward(CallPutFlag, S - dS, X, T, Tf, r, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EBlackScholesForward = 1 / (4 * dS * 0.01) * (BlackScholesForward(CallPutFlag, S + dS, X, T, Tf, r, v + 0.01) - BlackScholesForward(CallPutFlag, S + dS, X, T, Tf, r, v - 0.01) _
        - BlackScholesForward(CallPutFlag, S - dS, X, T, Tf, r, v + 0.01) + BlackScholesForward(CallPutFlag, S - dS, X, T, Tf, r, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EBlackScholesForward = (BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v + 0.01) - BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EBlackScholesForward = (BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v + 0.01) - 2 * BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v) + BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EBlackScholesForward = v / 0.1 * (BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v + 0.01) - BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EBlackScholesForward = (BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v + 0.01) - 2 * BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v) + BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EBlackScholesForward = BlackScholesForward(CallPutFlag, S, X, 0.00001, Tf - 1 / 365, r, v) - BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v)
        Else
                EBlackScholesForward = BlackScholesForward(CallPutFlag, S, X, T - 1 / 365, Tf - 1 / 365, r, v) - BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v)
        End If
     ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EBlackScholesForward = (BlackScholesForward(CallPutFlag, S, X, T, Tf, r + 0.01, v) - BlackScholesForward(CallPutFlag, S, X, T, Tf, r - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EBlackScholesForward = 1 / dS ^ 3 * (BlackScholesForward(CallPutFlag, S + 2 * dS, X, T, Tf, r, v) - 3 * BlackScholesForward(CallPutFlag, S + dS, X, T, Tf, r, v) _
                                + 3 * BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v) - BlackScholesForward(CallPutFlag, S - dS, X, T, Tf, r, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EBlackScholesForward = (BlackScholesForward(CallPutFlag, S, X + dS, T, Tf, r, v) - BlackScholesForward(CallPutFlag, S, X - dS, T, Tf, r, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EBlackScholesForward = (BlackScholesForward(CallPutFlag, S, X + dS, T, Tf, r, v) - 2 * BlackScholesForward(CallPutFlag, S, X, T, Tf, r, v) + BlackScholesForward(CallPutFlag, S, X - dS, T, Tf, r, v)) / dS ^ 2
    End If
End Function
