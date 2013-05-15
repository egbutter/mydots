Attribute VB_Name = "PlainVanilla"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.
Global Const Pi = 3.14159265358979


Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


' Programmer Espen Gaarder Haug, Copyright 2006

Public Function GBlackScholesNGreeks(OutPutFlag As String, CallPutFlag As String, S As Double, x As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    dS = 0.01
    If OutPutFlag = "p" Then ' Value
        GBlackScholesNGreeks = GBlackScholes(CallPutFlag, S, x, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S + dS, x, T, r, b, v) - GBlackScholes(CallPutFlag, S - dS, x, T, r, b, v)) / (2 * dS)
      ElseIf OutPutFlag = "dP" Then 'Delta
            dS = 0.25
            GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S * (1 + dS), x, T, r, b, v) - GBlackScholes(CallPutFlag, S * (1 - dS), x, T, r, b, v)) * 2 / S
    ElseIf OutPutFlag = "e" Then 'Elasticity
         GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S + dS, x, T, r, b, v) - GBlackScholes(CallPutFlag, S - dS, x, T, r, b, v)) / (2 * dS) * S / GBlackScholes(CallPutFlag, S, x, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S + dS, x, T, r, b, v) - 2 * GBlackScholes(CallPutFlag, S, x, T, r, b, v) + GBlackScholes(CallPutFlag, S - dS, x, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S + dS, x, T, r, b, v + 0.01) - 2 * GBlackScholes(CallPutFlag, S, x, T, r, b, v + 0.01) + GBlackScholes(CallPutFlag, S - dS, x, T, r, b, v + 0.01) _
      - GBlackScholes(CallPutFlag, S + dS, x, T, r, b, v - 0.01) + 2 * GBlackScholes(CallPutFlag, S, x, T, r, b, v - 0.01) - GBlackScholes(CallPutFlag, S - dS, x, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        GBlackScholesNGreeks = S / 100 * (GBlackScholes(CallPutFlag, S + dS, x, T, r, b, v) - 2 * GBlackScholes(CallPutFlag, S, x, T, r, b, v) + GBlackScholes(CallPutFlag, S - dS, x, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        GBlackScholesNGreeks = 1 / (4 * dS * 0.01) * (GBlackScholes(CallPutFlag, S + dS, x, T, r, b, v + 0.01) - GBlackScholes(CallPutFlag, S + dS, x, T, r, b, v - 0.01) _
        - GBlackScholes(CallPutFlag, S - dS, x, T, r, b, v + 0.01) + GBlackScholes(CallPutFlag, S - dS, x, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S, x, T, r, b, v + 0.01) - GBlackScholes(CallPutFlag, S, x, T, r, b, v - 0.01)) / 2
    ElseIf OutPutFlag = "vp" Then 'VegaP
         GBlackScholesNGreeks = v / 0.1 * (GBlackScholes(CallPutFlag, S, x, T, r, b, v + 0.01) - GBlackScholes(CallPutFlag, S, x, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S, x, T, r, b, v + 0.01) - 2 * GBlackScholes(CallPutFlag, S, x, T, r, b, v) + GBlackScholes(CallPutFlag, S, x, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                GBlackScholesNGreeks = GBlackScholes(CallPutFlag, S, x, 0.00001, r, b, v) - GBlackScholes(CallPutFlag, S, x, T, r, b, v)
        Else
                GBlackScholesNGreeks = GBlackScholes(CallPutFlag, S, x, T - 1 / 365, r, b, v) - GBlackScholes(CallPutFlag, S, x, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S, x, T, r + 0.01, b + 0.01, v) - GBlackScholes(CallPutFlag, S, x, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S, x, T, r + 0.01, 0, v) - GBlackScholes(CallPutFlag, S, x, T, r - 0.01, 0, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S, x, T, r, b - 0.01, v) - GBlackScholes(CallPutFlag, S, x, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S, x, T, r, b + 0.01, v) - GBlackScholes(CallPutFlag, S, x, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        GBlackScholesNGreeks = 1 / dS ^ 3 * (GBlackScholes(CallPutFlag, S + 2 * dS, x, T, r, b, v) - 3 * GBlackScholes(CallPutFlag, S + dS, x, T, r, b, v) _
                                + 3 * GBlackScholes(CallPutFlag, S, x, T, r, b, v) - GBlackScholes(CallPutFlag, S - dS, x, T, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S, x + dS, T, r, b, v) - GBlackScholes(CallPutFlag, S, x - dS, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        GBlackScholesNGreeks = (GBlackScholes(CallPutFlag, S, x + dS, T, r, b, v) - 2 * GBlackScholes(CallPutFlag, S, x, T, r, b, v) + GBlackScholes(CallPutFlag, S, x - dS, T, r, b, v)) / dS ^ 2
    End If
End Function

'// Black (1976) Options on futures/forwards
Public Function Black76(CallPutFlag As String, F As Double, x _
                As Double, T As Double, r As Double, v As Double) As Double
                
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(F / x) + (v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    If CallPutFlag = "c" Then
        Black76 = Exp(-r * T) * (F * CND(d1) - x * CND(d2))
    ElseIf CallPutFlag = "p" Then
        Black76 = Exp(-r * T) * (x * CND(-d2) - F * CND(-d1))
    End If
    
End Function


'// Garman and Kohlhagen (1983) Currency options
Public Function GarmanKolhagen(CallPutFlag As String, S As Double, x _
                As Double, T As Double, r As Double, rf As Double, v As Double) As Double
                
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (r - rf + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    If CallPutFlag = "c" Then
        GarmanKolhagen = S * Exp(-rf * T) * CND(d1) - x * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
        GarmanKolhagen = x * Exp(-r * T) * CND(-d2) - S * Exp(-rf * T) * CND(-d1)
    End If
    
End Function


'//  The generalized Black and Scholes formula
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


'// This is the generlaized Black-Scholes-Merton formula including all greeeks
'// This function is simply calling all the other functions
Public Function CGBlackScholes(OutPutFlag As String, Optional CallPutFlag As String, Optional S As Double, Optional x _
                As Double, Optional T As Double, Optional r As Double, Optional b As Double, Optional v As Double, Optional delta As Double, Optional InTheMoneyProb As Double, Optional ThetaDays As Double) As Double
                
                    Dim output As Double
                  
                    output = 0
                    
                If OutPutFlag = "p" Then 'Value
                    CGBlackScholes = GBlackScholes(CallPutFlag, S, x, T, r, b, v)
                    
                'DELTA GREEKS
                 ElseIf OutPutFlag = "d" Then 'Delta
                    CGBlackScholes = GDelta(CallPutFlag, S, x, T, r, b, v)
                ElseIf OutPutFlag = "df" Then 'Forward Delta
                    CGBlackScholes = GForwardDelta(CallPutFlag, S, x, T, r, b, v)
                ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
                    CGBlackScholes = GDdeltaDvol(S, x, T, r, b, v) / 100
                ElseIf OutPutFlag = "dvv" Then 'DDeltaDvolDvol
                    CGBlackScholes = GDdeltaDvolDvol(S, x, T, r, b, v) / 10000
                ElseIf OutPutFlag = "dt" Then 'DDeltaDtime/Charm
                    CGBlackScholes = GDdeltaDtime(CallPutFlag, S, x, T, r, b, v) / 365
                ElseIf OutPutFlag = "dmx" Then
                    CGBlackScholes = S ^ 2 / x * Exp((2 * b + v ^ 2) * T)
                    
                 ElseIf OutPutFlag = "e" Then ' Elasticity
                     CGBlackScholes = GElasticity(CallPutFlag, S, x, T, r, b, v)
         
                'GAMMA GREEKS
                ElseIf OutPutFlag = "sg" Then 'SaddleGamma
                CGBlackScholes = GSaddleGamma(x, T, r, b, v)
                 ElseIf OutPutFlag = "g" Then 'Gamma
                    CGBlackScholes = GGamma(S, x, T, r, b, v)
                 ElseIf OutPutFlag = "s" Then 'DgammaDspot/speed
                    CGBlackScholes = GDgammaDspot(S, x, T, r, b, v)
                 ElseIf OutPutFlag = "gv" Then 'DgammaDvol/Zomma
                    CGBlackScholes = GDgammaDvol(S, x, T, r, b, v) / 100
                 ElseIf OutPutFlag = "gt" Then 'DgammaDtime
                    CGBlackScholes = GDgammaDtime(S, x, T, r, b, v) / 365
                    
                ElseIf OutPutFlag = "gp" Then 'GammaP
                    CGBlackScholes = GGammaP(S, x, T, r, b, v)
                ElseIf OutPutFlag = "gps" Then 'DgammaPDspot
                    CGBlackScholes = GDgammaPDspot(S, x, T, r, b, v)
                ElseIf OutPutFlag = "gpv" Then 'DgammaDvol/Zomma
                    CGBlackScholes = GDgammaPDvol(S, x, T, r, b, v) / 100
                ElseIf OutPutFlag = "gpt" Then 'DgammaPDtime
                    CGBlackScholes = GDgammaPDtime(S, x, T, r, b, v) / 365
              
                    
                'VEGA GREEKS
                 ElseIf OutPutFlag = "v" Then 'Vega
                    CGBlackScholes = GVega(S, x, T, r, b, v) / 100
                 ElseIf OutPutFlag = "vt" Then 'DvegaDtime
                    CGBlackScholes = GDvegaDtime(S, x, T, r, b, v) / 365
                ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol/Vomma
                    CGBlackScholes = GDvegaDvol(S, x, T, r, b, v) / 10000
                ElseIf OutPutFlag = "vvv" Then 'DvommaDvol
                    CGBlackScholes = GDvommaDvol(S, x, T, r, b, v) / 1000000
                    
                ElseIf OutPutFlag = "vp" Then 'VegaP
                    CGBlackScholes = GVegaP(S, x, T, r, b, v)
                ElseIf OutPutFlag = "vpv" Then 'DvegaPDvol/VommaP
                    CGBlackScholes = GDvegaPDvol(S, x, T, r, b, v) / 100
                ElseIf OutPutFlag = "vl" Then 'Vega Leverage
                    CGBlackScholes = GVegaLeverage(CallPutFlag, S, x, T, r, b, v)
                
                'VARIANCE GREEKS
                ElseIf OutPutFlag = "varvega" Then 'Variance-Vega
                    CGBlackScholes = GVarianceVega(S, x, T, r, b, v) / 100
                 ElseIf OutPutFlag = "vardelta" Then 'Variance-delta
                    CGBlackScholes = GVarianceDelta(S, x, T, r, b, v) / 100
                 ElseIf OutPutFlag = "varvar" Then 'Variance-vomma
                    CGBlackScholes = GVarianceVomma(S, x, T, r, b, v) / 10000
                ElseIf OutPutFlag = "varult" Then 'Variance-ultima
                    CGBlackScholes = GVarianceUltima(S, x, T, r, b, v) / 1000000
                
                
                'THETA GREEKS
                ElseIf OutPutFlag = "t" Then 'Theta
                    CGBlackScholes = GTheta(CallPutFlag, S, x, T, r, b, v) / 365
                ElseIf OutPutFlag = "Dlt" Then 'Drift-less Theta
                    CGBlackScholes = GThetaDriftLess(S, x, T, r, b, v) / 365
                  
                'RATE/CARRY GREEKS
                ElseIf OutPutFlag = "r" Then 'Rho
                    CGBlackScholes = GRho(CallPutFlag, S, x, T, r, b, v) / 100
                 ElseIf OutPutFlag = "fr" Then 'Rho futures option
                    CGBlackScholes = GRhoFO(CallPutFlag, S, x, T, r, b, v) / 100
                ElseIf OutPutFlag = "b" Then 'Carry Rho
                    CGBlackScholes = GCarry(CallPutFlag, S, x, T, r, b, v) / 100
                ElseIf OutPutFlag = "f" Then 'Phi/Rho2
                    CGBlackScholes = GPhi(CallPutFlag, S, x, T, r, b, v) / 100
                
                'PROB GREEKS
                ElseIf OutPutFlag = "z" Then 'Zeta/In-the-money risk neutral probability
                    CGBlackScholes = GInTheMoneyProbability(CallPutFlag, S, x, T, b, v)
                ElseIf OutPutFlag = "zv" Then 'DzetaDvol
                    CGBlackScholes = GDzetaDvol(CallPutFlag, S, x, T, r, b, v) / 100
                ElseIf OutPutFlag = "zt" Then 'DzetaDtime
                    CGBlackScholes = GDzetaDtime(CallPutFlag, S, x, T, r, b, v) / 365
                ElseIf OutPutFlag = "bp" Then 'Brak even probability
                    CGBlackScholes = GBreakEvenProbability(CallPutFlag, S, x, T, r, b, v)
                 ElseIf OutPutFlag = "dx" Then 'StrikeDelta
                    CGBlackScholes = GStrikeDelta(CallPutFlag, S, x, T, r, b, v)
                ElseIf OutPutFlag = "dxdx" Then 'Risk Neutral Density
                    CGBlackScholes = GRiskNeutralDensity(S, x, T, r, b, v)
                    
                'FROM DELTA GREEKS
                ElseIf OutPutFlag = "gfd" Then 'Gamma from delta
                    CGBlackScholes = GGammaFromDelta(S, T, r, b, v, delta)
                  ElseIf OutPutFlag = "gpfd" Then 'GammaP from delta
                    CGBlackScholes = GGammaPFromDelta(S, T, r, b, v, delta)
                 ElseIf OutPutFlag = "vfd" Then 'Vega from delta
                    CGBlackScholes = GVegaFromDelta(S, T, r, b, delta) / 100
                ElseIf OutPutFlag = "vpfd" Then 'VegaP from delta
                    CGBlackScholes = GVegaPFromDelta(S, T, r, b, v, delta)
                 ElseIf OutPutFlag = "xfd" Then 'Strike from delta
                    CGBlackScholes = GStrikeFromDelta(CallPutFlag, S, T, r, b, v, delta)
                  ElseIf OutPutFlag = "ipfd" Then 'In-the-money probability from delta
                    CGBlackScholes = InTheMoneyProbFromDelta(CallPutFlag, S, T, r, b, v, delta)
                    
                    
                 'FROM IN-THE GREEKS
                 ElseIf OutPutFlag = "xfip" Then 'Strike from in-the-money probability
                    CGBlackScholes = GStrikeFromInTheMoneyProb(CallPutFlag, S, v, T, b, InTheMoneyProb)
                ElseIf OutPutFlag = "RNDfip" Then 'Risk Neutral Density from in-the-money probability
                    CGBlackScholes = GRNDFromInTheMoneyProb(x, T, r, v, InTheMoneyProb)
                ElseIf OutPutFlag = "dfip" Then 'Strike from in-the-money probability
                    CGBlackScholes = GDeltaFromInTheMoneyProb(CallPutFlag, S, T, r, b, v, InTheMoneyProb)
                    
                    
                'CALCULATIONS
                ElseIf OutPutFlag = "d1" Then 'd1
                    CGBlackScholes = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
                ElseIf OutPutFlag = "d2" Then 'd2
                    CGBlackScholes = (Log(S / x) + (b - v ^ 2 / 2) * T) / (v * Sqr(T))
                ElseIf OutPutFlag = "nd1" Then 'n(d1)
                    CGBlackScholes = ND((Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T)))
                ElseIf OutPutFlag = "nd2" Then 'n(d2)
                    CGBlackScholes = ND((Log(S / x) + (b - v ^ 2 / 2) * T) / (v * Sqr(T)))
                ElseIf OutPutFlag = "CNDd1" Then 'N(d1)
                    CGBlackScholes = CND((Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T)))
                ElseIf OutPutFlag = "CNDd2" Then 'N(d2)
                    CGBlackScholes = CND((Log(S / x) + (b - v ^ 2 / 2) * T) / (v * Sqr(T)))
                End If
End Function

'// DDeltaDvol also known as vanna
Public Function GDdeltaDvol(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double

Dim d1 As Double, d2 As Double

    d1 = (Log(S / x) + (b + v * v / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GDdeltaDvol = -Exp((b - r) * T) * d2 / v * ND(d1)
End Function

'// DDeltaDvolDvol also known as DVannaDvol
Public Function GDdeltaDvolDvol(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double

Dim d1 As Double, d2 As Double

    d1 = (Log(S / x) + (b + v * v / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GDdeltaDvolDvol = GDdeltaDvol(S, x, T, r, b, v) * 1 / v * (d1 * d2 - d1 / d2 - 1)

End Function

'// Vega from delta
Public Function GVegaFromDelta(S As Double, T As Double, r As Double, b As Double, delta As Double) As Double

    
    GVegaFromDelta = S * Exp((b - r) * T) * Sqr(T) * ND(CNDEV(Exp((r - b) * T) * Abs(delta)))
    
End Function

'// Gamma from delta
Public Function GGammaFromDelta(S As Double, T As Double, r As Double, b As Double, v As Double, delta As Double) As Double

    GGammaFromDelta = Exp((b - r) * T) * ND(CNDEV(Exp((r - b) * T) * Abs(delta))) / (S * v * Sqr(T))
    
End Function

'// Risk Neutral Density from in-the-money probability
Public Function GRNDFromInTheMoneyProb(x As Double, T As Double, r As Double, v As Double, Probability As Double) As Double

    GRNDFromInTheMoneyProb = Exp(-r * T) * ND(CNDEV(Probability)) / (x * v * Sqr(T))
    
End Function

'// GammaP from delta
Public Function GGammaPFromDelta(S As Double, T As Double, r As Double, b As Double, v As Double, delta As Double) As Double

    GGammaPFromDelta = S / 100 * GGammaFromDelta(S, T, r, b, v, delta)
    
End Function

'// VegaP from delta
Public Function GVegaPFromDelta(S As Double, T As Double, r As Double, b As Double, v As Double, delta As Double) As Double

    GVegaPFromDelta = v / 10 * GVegaFromDelta(S, T, r, b, delta)
    
End Function

'// What asset price that gives maximum DdeltaDvol
Public Function MaxDdeltaDvolAsset(UpperLowerFlag As String, x As Double, T As Double, b As Double, v As Double) As Double
    '// UpperLowerFlag"l" gives lower asset level that gives max DdeltaDvol
    '// UpperLowerFlag"l" gives upper asset level that gives max DdeltaDvol
    
    If UpperLowerFlag = "l" Then
        MaxDdeltaDvolAsset = x * Exp(-b * T - v * Sqr(T) * Sqr(4 + T * v ^ 2) / 2)
    ElseIf UpperLowerFlag = "u" Then
        MaxDdeltaDvolAsset = x * Exp(-b * T + v * Sqr(T) * Sqr(4 + T * v ^ 2) / 2)
    End If
    
End Function

'// What strike price that gives maximum DdeltaDvol
Public Function MaxDdeltaDvolStrike(UpperLowerFlag As String, S As Double, T As Double, b As Double, v As Double) As Double
    
    '// UpperLowerFlag"l" gives lower strike level that gives max DdeltaDvol
    '// UpperLowerFlag"l" gives upper strike level that gives max DdeltaDvol
    
    If UpperLowerFlag = "l" Then
        MaxDdeltaDvolStrike = S * Exp(b * T - v * Sqr(T) * Sqr(4 + T * v ^ 2) / 2)
    ElseIf UpperLowerFlag = "u" Then
        MaxDdeltaDvolStrike = S * Exp(b * T + v * Sqr(T) * Sqr(4 + T * v ^ 2) / 2)
    End If
    
End Function

'// What strike price that gives maximum gamma and vega
Public Function GMaxGammaVegaatX(S As Double, b As Double, T As Double, v As Double)
            
            GMaxGammaVegaatX = S * Exp((b + v * v / 2) * T)

End Function

'// What asset price that gives maximum gamma
Public Function GMaxGammaatS(x As Double, b As Double, T As Double, v As Double)

            GMaxGammaatS = x * Exp((-b - 3 * v * v / 2) * T)

End Function

'// What asset price that gives maximum vega
Public Function GMaxVegaatS(x As Double, b As Double, T As Double, v As Double)

            GMaxVegaatS = x * Exp((-b + v * v / 2) * T)
            
End Function



'// Forward delta for the generalized Black and Scholes formula
Public Function GForwardDelta(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, _
                b As Double, v As Double) As Double
                
    Dim d1 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    
    If CallPutFlag = "c" Then
        GForwardDelta = Exp(-r * T) * CND(d1)
    ElseIf CallPutFlag = "p" Then
        GForwardDelta = Exp(-r * T) * (CND(d1) - 1)
    End If
End Function

'// DZetaDvol for the generalized Black and Scholes formula
Public Function GDzetaDvol(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
                
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    If CallPutFlag = "c" Then
        GDzetaDvol = -ND(d2) * d1 / v
    Else
        GDzetaDvol = ND(d2) * d1 / v
    End If

End Function

'// DZetaDtime for the generalized Black and Scholes formula
Public Function GDzetaDtime(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
                
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    If CallPutFlag = "c" Then
        GDzetaDtime = ND(d2) * (b / (v * Sqr(T)) - d1 / (2 * T))
    Else
        GDzetaDtime = -ND(d2) * (b / (v * Sqr(T)) - d1 / (2 * T))
    End If

End Function

'// Delta for the generalized Black and Scholes formula
Public Function GInTheMoneyProbability(CallPutFlag As String, S As Double, x As Double, T As Double, _
                b As Double, v As Double) As Double
                
    Dim d2 As Double
    
    d2 = (Log(S / x) + (b - v ^ 2 / 2) * T) / (v * Sqr(T))
    
    If CallPutFlag = "c" Then
        GInTheMoneyProbability = CND(d2)
    ElseIf CallPutFlag = "p" Then
        GInTheMoneyProbability = CND(-d2)
    End If
    
End Function

'// Risk neutral break even probability for the generalized Black and Scholes formula
Public Function GBreakEvenProbability(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, _
                b As Double, v As Double) As Double
                
    Dim d2 As Double
    
    If CallPutFlag = "c" Then
        x = x + GBlackScholes("c", S, x, T, r, b, v) * Exp(r * T)
        d2 = (Log(S / x) + (b - v ^ 2 / 2) * T) / (v * Sqr(T))
        GBreakEvenProbability = CND(d2)
    ElseIf CallPutFlag = "p" Then
        x = x - GBlackScholes("p", S, x, T, r, b, v) * Exp(r * T)
        d2 = (Log(S / x) + (b - v ^ 2 / 2) * T) / (v * Sqr(T))
        GBreakEvenProbability = CND(-d2)
    End If
    
End Function



'// Closed form solution to find strike given the in-the-money risk neutral probability
Public Function GStrikeFromInTheMoneyProb(CallPutFlag As String, S As Double, v As Double, T As Double, b As Double, InTheMoneyProb As Double) As Double
        
    If CallPutFlag = "c" Then
          GStrikeFromInTheMoneyProb = S * Exp(-CNDEV(InTheMoneyProb) * v * Sqr(T) + (b - v * v / 2) * T)
        Else
           GStrikeFromInTheMoneyProb = S * Exp(CNDEV(InTheMoneyProb) * v * Sqr(T) + (b - v * v / 2) * T)
        End If
        
End Function
'// Closed form solution to find strike given the delta
Public Function GStrikeFromDelta(CallPutFlag As String, S As Double, T As Double, r As Double, b As Double, v As Double, delta As Double) As Double
        
       If CallPutFlag = "c" Then
          GStrikeFromDelta = S * Exp(-CNDEV(delta * Exp((r - b) * T)) * v * Sqr(T) + (b + v * v / 2) * T)
        Else
            GStrikeFromDelta = S * Exp(CNDEV(-delta * Exp((r - b) * T)) * v * Sqr(T) + (b + v * v / 2) * T)
        End If
        
End Function


'// Closed form solution to find in-the-money risk-neutral probaility given the delta
Public Function InTheMoneyProbFromDelta(CallPutFlag As String, S As Double, T As Double, r As Double, b As Double, v As Double, delta As Double) As Double
        
       If CallPutFlag = "c" Then
          InTheMoneyProbFromDelta = CND(CNDEV(delta / Exp((b - r) * T)) - v * Sqr(T))
        Else
            InTheMoneyProbFromDelta = CND(CNDEV(-delta / Exp((b - r) * T)) + v * Sqr(T))
        End If
        
End Function

'// Closed form solution to find in-the-money risk-neutral probaility given the delta
Public Function GDeltaFromInTheMoneyProb(CallPutFlag As String, S As Double, T As Double, r As Double, b As Double, v As Double, InTheMoneyProb As Double) As Double
        
       If CallPutFlag = "c" Then
          GDeltaFromInTheMoneyProb = CND(CNDEV(InTheMoneyProb * Exp((b - r) * T)) - v * Sqr(T))
        Else
            GDeltaFromInTheMoneyProb = -CND(CNDEV(InTheMoneyProb * Exp((b - r) * T)) + v * Sqr(T))
        End If
        
End Function

'// MirrorDeltaStrike, delta neutral straddle strike in the BSM formula
Public Function GDeltaMirrorStrike(S As Double, T As Double, _
                b As Double, v As Double) As Double
    
        GDeltaMirrorStrike = S * Exp((b + v ^ 2 / 2) * T)
    
End Function


'// MirrorProbabilityStrike, probability neutral straddle strike in the BSM formula
Public Function GProbabilityMirrorStrike(S As Double, T As Double, _
                b As Double, v As Double) As Double
    
        GProbabilityMirrorStrike = S * Exp((b - v ^ 2 / 2) * T)
    
End Function

'// MirrorDeltaStrike, general delta symmmetric strike in the BSM formula
Public Function GDeltaMirrorCallPutStrike(S As Double, x As Double, T As Double, _
                b As Double, v As Double) As Double
    
        GDeltaMirrorCallPutStrike = S ^ 2 / x * Exp((2 * b + v ^ 2) * T)
    
End Function


'// Gamma for the generalized Black and Scholes formula
Public Function GGamma(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    GGamma = Exp((b - r) * T) * ND(d1) / (S * v * Sqr(T))
End Function

'// SaddleGamma for the generalized Black and Scholes formula
Public Function GSaddleGamma(x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    GSaddleGamma = Sqr(Exp(1) / Application.Pi()) * Sqr((2 * b - r) / v ^ 2 + 1) / x
    
End Function


'// GammaP for the generalized Black and Scholes formula
Public Function GGammaP(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    GGammaP = S * GGamma(S, x, T, r, b, v) / 100
    
End Function

'// Delta for the generalized Black and Scholes formula
Public Function GDelta(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    If CallPutFlag = "c" Then
        GDelta = Exp((b - r) * T) * CND(d1)
    Else
        GDelta = -Exp((b - r) * T) * CND(-d1)
    End If
    
End Function

'// StrikeDelta for the generalized Black and Scholes formula
Public Function GStrikeDelta(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d2 As Double
    
    d2 = (Log(S / x) + (b - v ^ 2 / 2) * T) / (v * Sqr(T))
    If CallPutFlag = "c" Then
        GStrikeDelta = -Exp(-r * T) * CND(d2)
    Else
        GStrikeDelta = Exp(-r * T) * CND(-d2)
    End If
    
End Function


'// Elasticity for the generalized Black and Scholes formula
Public Function GElasticity(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
        GElasticity = GDelta(CallPutFlag, S, x, T, r, b, v) * S / GBlackScholes(CallPutFlag, S, x, T, r, b, v)
    
End Function


'// DgammaDvol/Zomma for the generalized Black and Scholes formula
Public Function GDgammaDvol(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GDgammaDvol = GGamma(S, x, T, r, b, v) * ((d1 * d2 - 1) / v)

End Function

'// DgammaPDvol for the generalized Black and Scholes formula
Public Function GDgammaPDvol(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GDgammaPDvol = S / 100 * GGamma(S, x, T, r, b, v) * ((d1 * d2 - 1) / v)

End Function

'// DgammaDspot/Speed for the generalized Black and Scholes formula
Public Function GDgammaDspot(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    
    GDgammaDspot = -GGamma(S, x, T, r, b, v) * (1 + d1 / (v * Sqr(T))) / S

End Function

'// DgammaPDspot/SpeedP for the generalized Black and Scholes formula
Public Function GDgammaPDspot(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    
    GDgammaPDspot = -GGamma(S, x, T, r, b, v) * (d1) / (100 * v * Sqr(T))

End Function

'// Risk Neutral Denisty for the generalized Black and Scholes formula
Public Function GRiskNeutralDensity(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d2 As Double
    
    d2 = (Log(S / x) + (b - v ^ 2 / 2) * T) / (v * Sqr(T))
    GRiskNeutralDensity = Exp(-r * T) * ND(d2) / (x * v * Sqr(T))

End Function

'// Volatility estimate confidence interval
Function GConfidenceIntervalVolatility(Alfa As Double, n As Integer, VolatilityEstimate As Double, UpperLower As String)
    'UpperLower     ="L" gives the lower cofidence interval
    '               ="U" gives the upper cofidence interval
    'n: number of observations
    If UpperLower = "L" Then
        GConfidenceIntervalVolatility = VolatilityEstimate * Sqr((n - 1) / (Application.ChiInv(Alfa / 2, n - 1)))
    ElseIf UpperLower = "U" Then
        GConfidenceIntervalVolatility = VolatilityEstimate * Sqr((n - 1) / (Application.ChiInv(1 - Alfa / 2, n - 1)))
    End If

End Function


'// Theta for the generalized Black and Scholes formula
Public Function GTheta(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then
        GTheta = -S * Exp((b - r) * T) * ND(d1) * v / (2 * Sqr(T)) - (b - r) * S * Exp((b - r) * T) * CND(d1) - r * x * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
        GTheta = -S * Exp((b - r) * T) * ND(d1) * v / (2 * Sqr(T)) + (b - r) * S * Exp((b - r) * T) * CND(-d1) + r * x * Exp(-r * T) * CND(-d2)
    End If

End Function


'// Drift-less Theta for the generalized Black and Scholes formula
Public Function GThetaDriftLess(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    GThetaDriftLess = -S * Exp((b - r) * T) * ND(d1) * v / (2 * Sqr(T))
    
End Function

'// Variance-vega for the generalized Black and Scholes formula
Public Function GVarianceVega(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    GVarianceVega = S * Exp((b - r) * T) * ND(d1) * Sqr(T) / (2 * v)

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

'// Variance-delta for the generalized Black and Scholes formula
Public Function GVarianceDelta(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GVarianceDelta = S * Exp((b - r) * T) * ND(d1) * (-d2) / (2 * v ^ 2)

End Function

'// Vega for the generalized Black and Scholes formula
Public Function GVega(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    GVega = S * Exp((b - r) * T) * ND(d1) * Sqr(T)

End Function

'// VegaP for the generalized Black and Scholes formula
Public Function GVegaP(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double

    GVegaP = v / 10 * GVega(S, x, T, r, b, v)

End Function


'// DdeltaDtime/Charm for the generalized Black and Scholes formula
Public Function GDdeltaDtime(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    
    If CallPutFlag = "c" Then
          GDdeltaDtime = -Exp((b - r) * T) * (ND(d1) * (b / (v * Sqr(T)) - d2 / (2 * T)) + (b - r) * CND(d1))
    ElseIf CallPutFlag = "p" Then
        GDdeltaDtime = -Exp((b - r) * T) * (ND(d1) * (b / (v * Sqr(T)) - d2 / (2 * T)) - (b - r) * CND(-d1))
   End If
    
End Function


'// Profitt/Loss STD for the generalized Black and Scholes formula
Public Function GProfitLossSTD(TypeFlag As String, CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double, NHedges As Integer) As Double
    
    If TypeFlag = "a" Then ' in dollars
        GProfitLossSTD = Sqr(Application.Pi() / 4) * GVega(S, x, T, r, b, v) * v / Sqr(NHedges)
    ElseIf TypeFlag = "p" Then ' in percent
        GProfitLossSTD = Sqr(Application.Pi() / 4) * GVega(S, x, T, r, b, v) * v / Sqr(NHedges) / GBlackScholes(CallPutFlag, S, x, T, r, b, v)
    End If

End Function

'// DvegaDvol/Vomma for the generalized Black and Scholes formula
Public Function GDvegaDvol(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GDvegaDvol = GVega(S, x, T, r, b, v) * d1 * d2 / v

End Function

'// DvegaPDvol/VommaP for the generalized Black and Scholes formula
Public Function GDvegaPDvol(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GDvegaPDvol = GVegaP(S, x, T, r, b, v) * d1 * d2 / v

End Function

'// DvegaDtime for the generalized Black and Scholes formula
Public Function GDvegaDtime(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GDvegaDtime = GVega(S, x, T, r, b, v) * (r - b + b * d1 / (v * Sqr(T)) - (1 + d1 * d2) / (2 * T))

End Function

'// DVommaDVol for the generalized Black and Scholes formula
Public Function GDvommaDvol(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GDvommaDvol = GDvegaDvol(S, x, T, r, b, v) * 1 / v * (d1 * d2 - d1 / d2 - d2 / d1 - 1)

End Function



'// GGammaDtime for the generalized Black and Scholes formula
Public Function GDgammaDtime(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GDgammaDtime = GGamma(S, x, T, r, b, v) * (r - b + b * d1 / (v * Sqr(T)) + (1 - d1 * d2) / (2 * T))

End Function

'// GGammaPDtime for the generalized Black and Scholes formula
Public Function GDgammaPDtime(S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    GDgammaPDtime = GGammaP(S, x, T, r, b, v) * (r - b + b * d1 / (v * Sqr(T)) + (1 - d1 * d2) / (2 * T))

End Function



'// Vega for the generalized Black and Scholes formula
Public Function GVegaLeverage(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    GVegaLeverage = GVega(S, x, T, r, b, v) * v / GBlackScholes(CallPutFlag, S, x, T, r, b, v)

End Function


'// Rho for the generalized Black and Scholes formula for all options except futures
Public Function GRho(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
   Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    If CallPutFlag = "c" Then
            GRho = T * x * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
            GRho = -T * x * Exp(-r * T) * CND(-d2)
    End If

End Function


'// Rho for the generalized Black and Scholes formula for Futures option
Public Function GRhoFO(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
            GRhoFO = -T * GBlackScholes(CallPutFlag, S, x, T, r, 0, v)
   
End Function

'// Rho2/Phi for the generalized Black and Scholes formula
Public Function GPhi(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double

    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    If CallPutFlag = "c" Then
        GPhi = -T * S * Exp((b - r) * T) * CND(d1)
    ElseIf CallPutFlag = "p" Then
        GPhi = T * S * Exp((b - r) * T) * CND(-d1)
    End If
    
End Function

'// Carry rho sensitivity for the generalized Black and Scholes formula
Public Function GCarry(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double

    d1 = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    If CallPutFlag = "c" Then
        GCarry = T * S * Exp((b - r) * T) * CND(d1)
    ElseIf CallPutFlag = "p" Then
        GCarry = -T * S * Exp((b - r) * T) * CND(-d1)
        End If

End Function




Function GBlackScholesVarianceNGreeks(OutPutFlag As String, CallPutFlag As String, S As Double, x As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        GBlackScholesVarianceNGreeks = GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S + dS, x, T, r, b, v) - GBlackScholesVariance(CallPutFlag, S - dS, x, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S + dS, x, T, r, b, v) - GBlackScholesVariance(CallPutFlag, S - dS, x, T, r, b, v)) / (2 * dS) * S / GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S + dS, x, T, r, b, v) - 2 * GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v) + GBlackScholesVariance(CallPutFlag, S - dS, x, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDvariance
        GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S + dS, x, T, r, b, v + 0.01) - 2 * GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v + 0.01) + GBlackScholesVariance(CallPutFlag, S - dS, x, T, r, b, v + 0.01) _
      - GBlackScholesVariance(CallPutFlag, S + dS, x, T, r, b, v - 0.01) + 2 * GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v - 0.01) - GBlackScholesVariance(CallPutFlag, S - dS, x, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        GBlackScholesVarianceNGreeks = S / 100 * (GBlackScholesVariance(CallPutFlag, S + dS, x, T, r, b, v) - 2 * GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v) + GBlackScholesVariance(CallPutFlag, S - dS, x, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvariance
        GBlackScholesVarianceNGreeks = 1 / (4 * dS * 0.01) * (GBlackScholesVariance(CallPutFlag, S + dS, x, T, r, b, v + 0.01) - GBlackScholesVariance(CallPutFlag, S + dS, x, T, r, b, v - 0.01) _
        - GBlackScholesVariance(CallPutFlag, S - dS, x, T, r, b, v + 0.01) + GBlackScholesVariance(CallPutFlag, S - dS, x, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Variance Vega
         GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v + 0.01) - GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v - 0.01)) / 2
    ElseIf OutPutFlag = "vp" Then 'Variance VegaP
         GBlackScholesVarianceNGreeks = v / 0.1 * (GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v + 0.01) - GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'Variance Dvegavariance
        GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v + 0.01) - 2 * GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v) + GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                GBlackScholesVarianceNGreeks = GBlackScholesVariance(CallPutFlag, S, x, 0.00001, r, b, v) - GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v)
        Else
                GBlackScholesVarianceNGreeks = GBlackScholesVariance(CallPutFlag, S, x, T - 1 / 365, r, b, v) - GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S, x, T, r + 0.01, b + 0.01, v) - GBlackScholesVariance(CallPutFlag, S, x, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S, x, T, r + 0.01, 0, v) - GBlackScholesVariance(CallPutFlag, S, x, T, r - 0.01, 0, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S, x, T, r, b - 0.01, v) - GBlackScholesVariance(CallPutFlag, S, x, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S, x, T, r, b + 0.01, v) - GBlackScholesVariance(CallPutFlag, S, x, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        GBlackScholesVarianceNGreeks = 1 / dS ^ 3 * (GBlackScholesVariance(CallPutFlag, S + 2 * dS, x, T, r, b, v) - 3 * GBlackScholesVariance(CallPutFlag, S + dS, x, T, r, b, v) _
                                + 3 * GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v) - GBlackScholesVariance(CallPutFlag, S - dS, x, T, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S, x + dS, T, r, b, v) - GBlackScholesVariance(CallPutFlag, S, x - dS, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        GBlackScholesVarianceNGreeks = (GBlackScholesVariance(CallPutFlag, S, x + dS, T, r, b, v) - 2 * GBlackScholesVariance(CallPutFlag, S, x, T, r, b, v) + GBlackScholesVariance(CallPutFlag, S, x - dS, T, r, b, v)) / dS ^ 2
    End If
End Function

'//  The generalized Black and Scholes formula on variance form
Public Function GBlackScholesVariance(CallPutFlag As String, S As Double, x _
                As Double, T As Double, r As Double, b As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    d1 = (Log(S / x) + (b + v / 2) * T) / Sqr(v * T)
    d2 = d1 - Sqr(v * T)

    If CallPutFlag = "c" Then
        GBlackScholesVariance = S * Exp((b - r) * T) * CND(d1) - x * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
        GBlackScholesVariance = x * Exp(-r * T) * CND(-d2) - S * Exp((b - r) * T) * CND(-d1)
    End If
    
End Function
