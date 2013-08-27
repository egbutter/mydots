Attribute VB_Name = "PowerDerivatives"
Option Explicit
' Programmer Espen Gaarder Haug
' Copyright 2006, Espen Gaarder Haug

  'Capped Power Option, based on paper by Esser
 Public Function CappedPowerOption(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, _
                            v As Double, i As Double, c As Double) As Double
    
    Dim e1 As Double, e2 As Double, e3 As Double, e4 As Double
    
    e1 = (Log(S / X ^ (1 / i)) + (b + (i - 1 / 2) * v ^ 2) * T) / (v * Sqr(T))
    e2 = e1 - i * v * Sqr(T)
    
    
    If CallPutFlag = "c" Then
            e3 = (Log(S / (X + c) ^ (1 / i)) + (b + (i - 1 / 2) * v ^ 2) * T) / (v * Sqr(T))
            e4 = e3 - i * v * Sqr(T)
            CappedPowerOption = S ^ i * Exp((i - 1) * (r + i * v ^ 2 / 2) * T - i * (r - b) * T) * (CND(e1) - CND(e3)) - Exp(-r * T) * (X * CND(e2) - (c + X) * CND(e4))
    ElseIf CallPutFlag = "p" Then 'Extended to also hold for put Espen G. Haug Oct 2004
            e3 = (Log(S / (X - c) ^ (1 / i)) + (b + (i - 1 / 2) * v ^ 2) * T) / (v * Sqr(T))
            e4 = e3 - i * v * Sqr(T)
            CappedPowerOption = Exp(-r * T) * (X * CND(-e2) - (X - c) * CND(-e4)) - S ^ i * Exp((i - 1) * (r + i * v ^ 2 / 2) * T - i * (r - b) * T) * (CND(-e1) - CND(-e3))
    End If
    
 End Function
 
 
 Public Function ECappedPowerOption(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, i As Double, c As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        ECappedPowerOption = CappedPowerOption(CallPutFlag, S, X, T, r, b, v, i, c)
    ElseIf OutPutFlag = "d" Then 'Delta
         ECappedPowerOption = (CappedPowerOption(CallPutFlag, S + dS, X, T, r, b, v, i, c) - CappedPowerOption(CallPutFlag, S - dS, X, T, r, b, v, i, c)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         ECappedPowerOption = (CappedPowerOption(CallPutFlag, S + dS, X, T, r, b, v, i, c) - CappedPowerOption(CallPutFlag, S - dS, X, T, r, b, v, i, c)) / (2 * dS) * S / CappedPowerOption(CallPutFlag, S, X, T, r, b, v, i, c)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ECappedPowerOption = (CappedPowerOption(CallPutFlag, S + dS, X, T, r, b, v, i, c) - 2 * CappedPowerOption(CallPutFlag, S, X, T, r, b, v, i, c) + CappedPowerOption(CallPutFlag, S - dS, X, T, r, b, v, i, c)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ECappedPowerOption = (CappedPowerOption(CallPutFlag, S + dS, X, T, r, b, v + 0.01, i, c) - 2 * CappedPowerOption(CallPutFlag, S, X, T, r, b, v + 0.01, i, c) + CappedPowerOption(CallPutFlag, S - dS, X, T, r, b, v + 0.01, i, c) _
      - CappedPowerOption(CallPutFlag, S + dS, X, T, r, b, v - 0.01, i, c) + 2 * CappedPowerOption(CallPutFlag, S, X, T, r, b, v - 0.01, i, c) - CappedPowerOption(CallPutFlag, S - dS, X, T, r, b, v - 0.01, i, c)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "gp" Then 'GammaP
        ECappedPowerOption = S / 100 * (CappedPowerOption(CallPutFlag, S + dS, X, T, r, b, v, i, c) - 2 * CappedPowerOption(CallPutFlag, S, X, T, r, b, v, i, c) + CappedPowerOption(CallPutFlag, S - dS, X, T, r, b, v, i, c)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        ECappedPowerOption = (CappedPowerOption(CallPutFlag, S, X, T + 1 / 365, r, b, v, i, c) - 2 * CappedPowerOption(CallPutFlag, S, X, T, r, b, v, i, c) + CappedPowerOption(CallPutFlag, S, X, T - 1 / 365, r, b, v, i, c)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ECappedPowerOption = 1 / (4 * dS * 0.01) * (CappedPowerOption(CallPutFlag, S + dS, X, T, r, b, v + 0.01, i, c) - CappedPowerOption(CallPutFlag, S + dS, X, T, r, b, v - 0.01, i, c) _
        - CappedPowerOption(CallPutFlag, S - dS, X, T, r, b, v + 0.01, i, c) + CappedPowerOption(CallPutFlag, S - dS, X, T, r, b, v - 0.01, i, c)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ECappedPowerOption = (CappedPowerOption(CallPutFlag, S, X, T, r, b, v + 0.01, i, c) - CappedPowerOption(CallPutFlag, S, X, T, r, b, v - 0.01, i, c)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        ECappedPowerOption = (CappedPowerOption(CallPutFlag, S, X, T, r, b, v + 0.01, i, c) - 2 * CappedPowerOption(CallPutFlag, S, X, T, r, b, v, i, c) + CappedPowerOption(CallPutFlag, S, X, T, r, b, v - 0.01, i, c)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ECappedPowerOption = v / 0.1 * (CappedPowerOption(CallPutFlag, S, X, T, r, b, v + 0.01, i, c) - CappedPowerOption(CallPutFlag, S, X, T, r, b, v - 0.01, i, c)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ECappedPowerOption = (CappedPowerOption(CallPutFlag, S, X, T, r, b, v + 0.01, i, c) - 2 * CappedPowerOption(CallPutFlag, S, X, T, r, b, v, i, c) + CappedPowerOption(CallPutFlag, S, X, T, r, b, v - 0.01, i, c))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ECappedPowerOption = CappedPowerOption(CallPutFlag, S, X, 0.00001, r, b, v, i, c) - CappedPowerOption(CallPutFlag, S, X, T, r, b, v, i, c)
        Else
                ECappedPowerOption = CappedPowerOption(CallPutFlag, S, X, T - 1 / 365, r, b, v, i, c) - CappedPowerOption(CallPutFlag, S, X, T, r, b, v, i, c)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         ECappedPowerOption = (CappedPowerOption(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v, i, c) - CappedPowerOption(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v, i, c)) / 2
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ECappedPowerOption = (CappedPowerOption(CallPutFlag, S, X, T, r + 0.01, b, v, i, c) - CappedPowerOption(CallPutFlag, S, X, T, r - 0.01, b, v, i, c)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         ECappedPowerOption = (CappedPowerOption(CallPutFlag, S, X, T, r, b - 0.01, v, i, c) - CappedPowerOption(CallPutFlag, S, X, T, r, b + 0.01, v, i, c)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        ECappedPowerOption = (CappedPowerOption(CallPutFlag, S, X, T, r, b + 0.01, v, i, c) - CappedPowerOption(CallPutFlag, S, X, T, r, b - 0.01, v, i, c)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ECappedPowerOption = 1 / dS ^ 3 * (CappedPowerOption(CallPutFlag, S + 2 * dS, X, T, r, b, v, i, c) - 3 * CappedPowerOption(CallPutFlag, S + dS, X, T, r, b, v, i, c) _
                                + 3 * CappedPowerOption(CallPutFlag, S, X, T, r, b, v, i, c) - CappedPowerOption(CallPutFlag, S - dS, X, T, r, b, v, i, c))
     ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ECappedPowerOption = (CappedPowerOption(CallPutFlag, S, X + dS, T, r, b, v, i, c) - CappedPowerOption(CallPutFlag, S, X - dS, T, r, b, v, i, c)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        ECappedPowerOption = (CappedPowerOption(CallPutFlag, S, X + dS, T, r, b, v, i, c) - 2 * CappedPowerOption(CallPutFlag, S, X, T, r, b, v, i, c) + CappedPowerOption(CallPutFlag, S, X - dS, T, r, b, v, i, c)) / dS ^ 2
    End If
End Function


  
'Powered Option, special case i=2  Heard On the Street
 Public Function PoweredOptioni2(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, _
                            v As Double) As Double
    
    Dim d0 As Double, d1 As Double, d2 As Double
    
    d0 = (Log(S / X) + (b + 3 * v ^ 2 / 2) * T) / (v * Sqr(T))
    d1 = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    
    If CallPutFlag = "c" Then
       PoweredOptioni2 = S ^ 2 * Exp((2 * b - r + v ^ 2) * T) * CND(d0) - 2 * X * S * Exp((b - r) * T) * CND(d1) + X ^ 2 * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
        PoweredOptioni2 = S ^ 2 * Exp((2 * b - r + v ^ 2) * T) * CND(-d0) - 2 * X * S * Exp((b - r) * T) * CND(-d1) + X ^ 2 * Exp(-r * T) * CND(-d2)

    End If
    
 End Function


Public Function EPoweredOptioni2(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EPoweredOptioni2 = PoweredOptioni2(CallPutFlag, S, X, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S + dS, X, T, r, b, v) - PoweredOptioni2(CallPutFlag, S - dS, X, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S + dS, X, T, r, b, v) - PoweredOptioni2(CallPutFlag, S - dS, X, T, r, b, v)) / (2 * dS) * S / PoweredOptioni2(CallPutFlag, S, X, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S + dS, X, T, r, b, v) - 2 * PoweredOptioni2(CallPutFlag, S, X, T, r, b, v) + PoweredOptioni2(CallPutFlag, S - dS, X, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S + dS, X, T, r, b, v + 0.01) - 2 * PoweredOptioni2(CallPutFlag, S, X, T, r, b, v + 0.01) + PoweredOptioni2(CallPutFlag, S - dS, X, T, r, b, v + 0.01) _
      - PoweredOptioni2(CallPutFlag, S + dS, X, T, r, b, v - 0.01) + 2 * PoweredOptioni2(CallPutFlag, S, X, T, r, b, v - 0.01) - PoweredOptioni2(CallPutFlag, S - dS, X, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EPoweredOptioni2 = S / 100 * (PoweredOptioni2(CallPutFlag, S + dS, X, T, r, b, v) - 2 * PoweredOptioni2(CallPutFlag, S, X, T, r, b, v) + PoweredOptioni2(CallPutFlag, S - dS, X, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S, X, T + 1 / 365, r, b, v) - 2 * PoweredOptioni2(CallPutFlag, S, X, T, r, b, v) + PoweredOptioni2(CallPutFlag, S, X, T - 1 / 365, r, b, v)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EPoweredOptioni2 = 1 / (4 * dS * 0.01) * (PoweredOptioni2(CallPutFlag, S + dS, X, T, r, b, v + 0.01) - PoweredOptioni2(CallPutFlag, S + dS, X, T, r, b, v - 0.01) _
        - PoweredOptioni2(CallPutFlag, S - dS, X, T, r, b, v + 0.01) + PoweredOptioni2(CallPutFlag, S - dS, X, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S, X, T, r, b, v + 0.01) - PoweredOptioni2(CallPutFlag, S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S, X, T, r, b, v + 0.01) - 2 * PoweredOptioni2(CallPutFlag, S, X, T, r, b, v) + PoweredOptioni2(CallPutFlag, S, X, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EPoweredOptioni2 = v / 0.1 * (PoweredOptioni2(CallPutFlag, S, X, T, r, b, v + 0.01) - PoweredOptioni2(CallPutFlag, S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S, X, T, r, b, v + 0.01) - 2 * PoweredOptioni2(CallPutFlag, S, X, T, r, b, v) + PoweredOptioni2(CallPutFlag, S, X, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EPoweredOptioni2 = PoweredOptioni2(CallPutFlag, S, X, 0.00001, r, b, v) - PoweredOptioni2(CallPutFlag, S, X, T, r, b, v)
        Else
                EPoweredOptioni2 = PoweredOptioni2(CallPutFlag, S, X, T - 1 / 365, r, b, v) - PoweredOptioni2(CallPutFlag, S, X, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v) - PoweredOptioni2(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S, X, T, r + 0.01, b, v) - PoweredOptioni2(CallPutFlag, S, X, T, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S, X, T, r, b - 0.01, v) - PoweredOptioni2(CallPutFlag, S, X, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S, X, T, r, b + 0.01, v) - PoweredOptioni2(CallPutFlag, S, X, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EPoweredOptioni2 = 1 / dS ^ 3 * (PoweredOptioni2(CallPutFlag, S + 2 * dS, X, T, r, b, v) - 3 * PoweredOptioni2(CallPutFlag, S + dS, X, T, r, b, v) _
                                + 3 * PoweredOptioni2(CallPutFlag, S, X, T, r, b, v) - PoweredOptioni2(CallPutFlag, S - dS, X, T, r, b, v))
     ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S, X + dS, T, r, b, v) - PoweredOptioni2(CallPutFlag, S, X - dS, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        EPoweredOptioni2 = (PoweredOptioni2(CallPutFlag, S, X + dS, T, r, b, v) - 2 * PoweredOptioni2(CallPutFlag, S, X, T, r, b, v) + PoweredOptioni2(CallPutFlag, S, X - dS, T, r, b, v)) / dS ^ 2
    End If
End Function







 Public Function PowerContract(S As Double, X As Double, T As Double, r As Double, b As Double, v As Double, i As Double)
        
        PowerContract = (S / X) ^ i * Exp(((b - v ^ 2 / 2) * i - r + i ^ 2 * v ^ 2 / 2) * T)
 
 End Function



Public Function EPowerContract(OutPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, i As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EPowerContract = PowerContract(S, X, T, r, b, v, i)
    ElseIf OutPutFlag = "d" Then 'Delta
         EPowerContract = (PowerContract(S + dS, X, T, r, b, v, i) - PowerContract(S - dS, X, T, r, b, v, i)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EPowerContract = (PowerContract(S + dS, X, T, r, b, v, i) - PowerContract(S - dS, X, T, r, b, v, i)) / (2 * dS) * S / PowerContract(S, X, T, r, b, v, i)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EPowerContract = (PowerContract(S + dS, X, T, r, b, v, i) - 2 * PowerContract(S, X, T, r, b, v, i) + PowerContract(S - dS, X, T, r, b, v, i)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EPowerContract = (PowerContract(S + dS, X, T, r, b, v + 0.01, i) - 2 * PowerContract(S, X, T, r, b, v + 0.01, i) + PowerContract(S - dS, X, T, r, b, v + 0.01, i) _
      - PowerContract(S + dS, X, T, r, b, v - 0.01, i) + 2 * PowerContract(S, X, T, r, b, v - 0.01, i) - PowerContract(S - dS, X, T, r, b, v - 0.01, i)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EPowerContract = S / 100 * (PowerContract(S + dS, X, T, r, b, v, i) - 2 * PowerContract(S, X, T, r, b, v, i) + PowerContract(S - dS, X, T, r, b, v, i)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        EPowerContract = (PowerContract(S, X, T + 1 / 365, r, b, v, i) - 2 * PowerContract(S, X, T, r, b, v, i) + PowerContract(S, X, T - 1 / 365, r, b, v, i)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EPowerContract = 1 / (4 * dS * 0.01) * (PowerContract(S + dS, X, T, r, b, v + 0.01, i) - PowerContract(S + dS, X, T, r, b, v - 0.01, i) _
        - PowerContract(S - dS, X, T, r, b, v + 0.01, i) + PowerContract(S - dS, X, T, r, b, v - 0.01, i)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EPowerContract = (PowerContract(S, X, T, r, b, v + 0.01, i) - PowerContract(S, X, T, r, b, v - 0.01, i)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EPowerContract = (PowerContract(S, X, T, r, b, v + 0.01, i) - 2 * PowerContract(S, X, T, r, b, v, i) + PowerContract(S, X, T, r, b, v - 0.01, i)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EPowerContract = v / 0.1 * (PowerContract(S, X, T, r, b, v + 0.01, i) - PowerContract(S, X, T, r, b, v - 0.01, i)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EPowerContract = (PowerContract(S, X, T, r, b, v + 0.01, i) - 2 * PowerContract(S, X, T, r, b, v, i) + PowerContract(S, X, T, r, b, v - 0.01, i))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EPowerContract = PowerContract(S, X, 0.00001, r, b, v, i) - PowerContract(S, X, T, r, b, v, i)
        Else
                EPowerContract = PowerContract(S, X, T - 1 / 365, r, b, v, i) - PowerContract(S, X, T, r, b, v, i)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EPowerContract = (PowerContract(S, X, T, r + 0.01, b + 0.01, v, i) - PowerContract(S, X, T, r - 0.01, b - 0.01, v, i)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EPowerContract = (PowerContract(S, X, T, r + 0.01, b, v, i) - PowerContract(S, X, T, r - 0.01, b, v, i)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EPowerContract = (PowerContract(S, X, T, r, b - 0.01, v, i) - PowerContract(S, X, T, r, b + 0.01, v, i)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EPowerContract = (PowerContract(S, X, T, r, b + 0.01, v, i) - PowerContract(S, X, T, r, b - 0.01, v, i)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EPowerContract = 1 / dS ^ 3 * (PowerContract(S + 2 * dS, X, T, r, b, v, i) - 3 * PowerContract(S + dS, X, T, r, b, v, i) _
                                + 3 * PowerContract(S, X, T, r, b, v, i) - PowerContract(S - dS, X, T, r, b, v, i))
    ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EPowerContract = (PowerContract(S, X + dS, T, r, b, v, i) - PowerContract(S, X - dS, T, r, b, v, i)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        EPowerContract = (PowerContract(S, X + dS, T, r, b, v, i) - 2 * PowerContract(S, X, T, r, b, v, i) + PowerContract(S, X - dS, T, r, b, v, i)) / dS ^ 2
    End If
End Function




Public Function EPoweredOption(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, i As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EPoweredOption = PoweredOption(CallPutFlag, S, X, T, r, b, v, i)
    ElseIf OutPutFlag = "d" Then 'Delta
         EPoweredOption = (PoweredOption(CallPutFlag, S + dS, X, T, r, b, v, i) - PoweredOption(CallPutFlag, S - dS, X, T, r, b, v, i)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EPoweredOption = (PoweredOption(CallPutFlag, S + dS, X, T, r, b, v, i) - PoweredOption(CallPutFlag, S - dS, X, T, r, b, v, i)) / (2 * dS) * S / PoweredOption(CallPutFlag, S, X, T, r, b, v, i)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EPoweredOption = (PoweredOption(CallPutFlag, S + dS, X, T, r, b, v, i) - 2 * PoweredOption(CallPutFlag, S, X, T, r, b, v, i) + PoweredOption(CallPutFlag, S - dS, X, T, r, b, v, i)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EPoweredOption = (PoweredOption(CallPutFlag, S + dS, X, T, r, b, v + 0.01, i) - 2 * PoweredOption(CallPutFlag, S, X, T, r, b, v + 0.01, i) + PoweredOption(CallPutFlag, S - dS, X, T, r, b, v + 0.01, i) _
      - PoweredOption(CallPutFlag, S + dS, X, T, r, b, v - 0.01, i) + 2 * PoweredOption(CallPutFlag, S, X, T, r, b, v - 0.01, i) - PoweredOption(CallPutFlag, S - dS, X, T, r, b, v - 0.01, i)) / (2 * 0.01 * dS ^ 2) / 100
 ElseIf OutPutFlag = "gp" Then 'GammaP
        EPoweredOption = S / 100 * (PoweredOption(CallPutFlag, S + dS, X, T, r, b, v, i) - 2 * PoweredOption(CallPutFlag, S, X, T, r, b, v, i) + PoweredOption(CallPutFlag, S - dS, X, T, r, b, v, i)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        EPoweredOption = (PoweredOption(CallPutFlag, S, X, T + 1 / 365, r, b, v, i) - 2 * PoweredOption(CallPutFlag, S, X, T, r, b, v, i) + PoweredOption(CallPutFlag, S, X, T - 1 / 365, r, b, v, i)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EPoweredOption = 1 / (4 * dS * 0.01) * (PoweredOption(CallPutFlag, S + dS, X, T, r, b, v + 0.01, i) - PoweredOption(CallPutFlag, S + dS, X, T, r, b, v - 0.01, i) _
        - PoweredOption(CallPutFlag, S - dS, X, T, r, b, v + 0.01, i) + PoweredOption(CallPutFlag, S - dS, X, T, r, b, v - 0.01, i)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EPoweredOption = (PoweredOption(CallPutFlag, S, X, T, r, b, v + 0.01, i) - PoweredOption(CallPutFlag, S, X, T, r, b, v - 0.01, i)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EPoweredOption = (PoweredOption(CallPutFlag, S, X, T, r, b, v + 0.01, i) - 2 * PoweredOption(CallPutFlag, S, X, T, r, b, v, i) + PoweredOption(CallPutFlag, S, X, T, r, b, v - 0.01, i)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EPoweredOption = v / 0.1 * (PoweredOption(CallPutFlag, S, X, T, r, b, v + 0.01, i) - PoweredOption(CallPutFlag, S, X, T, r, b, v - 0.01, i)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EPoweredOption = (PoweredOption(CallPutFlag, S, X, T, r, b, v + 0.01, i) - 2 * PoweredOption(CallPutFlag, S, X, T, r, b, v, i) + PoweredOption(CallPutFlag, S, X, T, r, b, v - 0.01, i))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EPoweredOption = PoweredOption(CallPutFlag, S, X, 0.00001, r, b, v, i) - PoweredOption(CallPutFlag, S, X, T, r, b, v, i)
        Else
                EPoweredOption = PoweredOption(CallPutFlag, S, X, T - 1 / 365, r, b, v, i) - PoweredOption(CallPutFlag, S, X, T, r, b, v, i)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EPoweredOption = (PoweredOption(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v, i) - PoweredOption(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v, i)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EPoweredOption = (PoweredOption(CallPutFlag, S, X, T, r + 0.01, b, v, i) - PoweredOption(CallPutFlag, S, X, T, r - 0.01, b, v, i)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EPoweredOption = (PoweredOption(CallPutFlag, S, X, T, r, b - 0.01, v, i) - PoweredOption(CallPutFlag, S, X, T, r, b + 0.01, v, i)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EPoweredOption = (PoweredOption(CallPutFlag, S, X, T, r, b + 0.01, v, i) - PoweredOption(CallPutFlag, S, X, T, r, b - 0.01, v, i)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EPoweredOption = 1 / dS ^ 3 * (PoweredOption(CallPutFlag, S + 2 * dS, X, T, r, b, v, i) - 3 * PoweredOption(CallPutFlag, S + dS, X, T, r, b, v, i) _
                                + 3 * PoweredOption(CallPutFlag, S, X, T, r, b, v, i) - PoweredOption(CallPutFlag, S - dS, X, T, r, b, v, i))
     ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EPoweredOption = (PoweredOption(CallPutFlag, S, X + dS, T, r, b, v, i) - PoweredOption(CallPutFlag, S, X - dS, T, r, b, v, i)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        EPoweredOption = (PoweredOption(CallPutFlag, S, X + dS, T, r, b, v, i) - 2 * PoweredOption(CallPutFlag, S, X, T, r, b, v, i) + PoweredOption(CallPutFlag, S, X - dS, T, r, b, v, i)) / dS ^ 2
    End If
End Function
  
 Public Function PoweredOption(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, _
                            v As Double, i As Double) As Double
    
    Dim d1 As Double, sum As Double
    Dim j As Integer
    
    If CallPutFlag = "c" Then
        sum = 0
        For j = 0 To i Step 1
            d1 = (Log(S / X) + (b + (i - j - 0.5) * v ^ 2) * T) / (v * Sqr(T))
            sum = sum + Application.Combin(i, j) * S ^ (i - j) * (-X) ^ j * Exp((i - j - 1) * (r + (i - j) * v ^ 2 / 2) * T - (i - j) * (r - b) * T) * CND(d1)
        Next
            PoweredOption = sum
    ElseIf CallPutFlag = "p" Then
        sum = 0
        For j = 0 To i Step 1
            d1 = (Log(S / X) + (b + (i - j - 0.5) * v ^ 2) * T) / (v * Sqr(T))
            sum = sum + Application.Combin(i, j) * (-S) ^ (i - j) * X ^ j * Exp((i - j - 1) * (r + (i - j) * v ^ 2 / 2) * T - (i - j) * (r - b) * T) * CND(-d1)
        Next
            PoweredOption = sum
    End If
    
 End Function



 Public Function PowerOptionAsymmetric(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, _
                            v As Double, i As Double) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / X ^ (1 / i)) + (b + (i - 1 / 2) * v ^ 2) * T) / (v * Sqr(T))
    d2 = d1 - i * v * Sqr(T)
    
    If CallPutFlag = "c" Then
            PowerOptionAsymmetric = S ^ i * Exp(((i - 1) * (r + i * v ^ 2 / 2) - i * (r - b)) * T) * CND(d1) - X * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
            PowerOptionAsymmetric = X * Exp(-r * T) * CND(-d2) - S ^ i * Exp(((i - 1) * (r + i * v ^ 2 / 2) - i * (r - b)) * T) * CND(-d1)
    End If
    
 End Function
 


Public Function EPowerOptionAsymmetric(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, i As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EPowerOptionAsymmetric = PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v, i)
    ElseIf OutPutFlag = "d" Then 'Delta
         EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S + dS, X, T, r, b, v, i) - PowerOptionAsymmetric(CallPutFlag, S - dS, X, T, r, b, v, i)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S + dS, X, T, r, b, v, i) - PowerOptionAsymmetric(CallPutFlag, S - dS, X, T, r, b, v, i)) / (2 * dS) * S / PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v, i)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S + dS, X, T, r, b, v, i) - 2 * PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v, i) + PowerOptionAsymmetric(CallPutFlag, S - dS, X, T, r, b, v, i)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S + dS, X, T, r, b, v + 0.01, i) - 2 * PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v + 0.01, i) + PowerOptionAsymmetric(CallPutFlag, S - dS, X, T, r, b, v + 0.01, i) _
      - PowerOptionAsymmetric(CallPutFlag, S + dS, X, T, r, b, v - 0.01, i) + 2 * PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v - 0.01, i) - PowerOptionAsymmetric(CallPutFlag, S - dS, X, T, r, b, v - 0.01, i)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "gp" Then 'GammaP
        EPowerOptionAsymmetric = S / 100 * (PowerOptionAsymmetric(CallPutFlag, S + dS, X, T, r, b, v, i) - 2 * PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v, i) + PowerOptionAsymmetric(CallPutFlag, S - dS, X, T, r, b, v, i)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S, X, T + 1 / 365, r, b, v, i) - 2 * PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v, i) + PowerOptionAsymmetric(CallPutFlag, S, X, T - 1 / 365, r, b, v, i)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EPowerOptionAsymmetric = 1 / (4 * dS * 0.01) * (PowerOptionAsymmetric(CallPutFlag, S + dS, X, T, r, b, v + 0.01, i) - PowerOptionAsymmetric(CallPutFlag, S + dS, X, T, r, b, v - 0.01, i) _
        - PowerOptionAsymmetric(CallPutFlag, S - dS, X, T, r, b, v + 0.01, i) + PowerOptionAsymmetric(CallPutFlag, S - dS, X, T, r, b, v - 0.01, i)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v + 0.01, i) - PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v - 0.01, i)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v + 0.01, i) - 2 * PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v, i) + PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v - 0.01, i)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EPowerOptionAsymmetric = v / 0.1 * (PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v + 0.01, i) - PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v - 0.01, i)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v + 0.01, i) - 2 * PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v, i) + PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v - 0.01, i))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EPowerOptionAsymmetric = PowerOptionAsymmetric(CallPutFlag, S, X, 0.00001, r, b, v, i) - PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v, i)
        Else
                EPowerOptionAsymmetric = PowerOptionAsymmetric(CallPutFlag, S, X, T - 1 / 365, r, b, v, i) - PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v, i)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v, i) - PowerOptionAsymmetric(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v, i)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S, X, T, r + 0.01, b, v, i) - PowerOptionAsymmetric(CallPutFlag, S, X, T, r - 0.01, b, v, i)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b - 0.01, v, i) - PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b + 0.01, v, i)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b + 0.01, v, i) - PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b - 0.01, v, i)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EPowerOptionAsymmetric = 1 / dS ^ 3 * (PowerOptionAsymmetric(CallPutFlag, S + 2 * dS, X, T, r, b, v, i) - 3 * PowerOptionAsymmetric(CallPutFlag, S + dS, X, T, r, b, v, i) _
                                + 3 * PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v, i) - PowerOptionAsymmetric(CallPutFlag, S - dS, X, T, r, b, v, i))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S, X + dS, T, r, b, v, i) - PowerOptionAsymmetric(CallPutFlag, S, X - dS, T, r, b, v, i)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        EPowerOptionAsymmetric = (PowerOptionAsymmetric(CallPutFlag, S, X + dS, T, r, b, v, i) - 2 * PowerOptionAsymmetric(CallPutFlag, S, X, T, r, b, v, i) + PowerOptionAsymmetric(CallPutFlag, S, X - dS, T, r, b, v, i)) / dS ^ 2
    End If
End Function

