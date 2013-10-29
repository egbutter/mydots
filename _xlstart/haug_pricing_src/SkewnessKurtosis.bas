Attribute VB_Name = "SkewnessKurtosis"
Option Explicit
Global Const Pi = 3.14159265358979

' Programmer Espen Gaarder Haug, Copyright 2006


Public Function SkewKurtCorradoSuModified(CallPutFlag As String, S As Double, X As Double, _
        T As Double, r As Double, b As Double, v As Double, Skew As Double, Kurt As Double) As Double

        Dim Q3 As Double, Q4 As Double
        Dim d As Double, w As Double
        Dim CallValue As Double
     

        w = Skew / 6 * v ^ 3 * T ^ 1.5 + Kurt / 24 * v ^ 4 * T ^ 2
        d = (Log(S / X) + (b + v ^ 2 / 2) * T - Log(1 + w)) / (v * Sqr(T))
        Q3 = 1 / (6 * (1 + w)) * S * v * Sqr(T) * (2 * v * Sqr(T) - d) * ND(d)
        Q4 = 1 / (24 * 1 + w) * S * v * Sqr(T) * (d ^ 2 - 3 * d * v * Sqr(T) + 3 * v ^ 2 * T - 1) * ND(d)
       
        CallValue = GBlackScholes("c", S, X, T, r, b, v) + Skew * Q3 + (Kurt - 3) * Q4
        If CallPutFlag = "c" Then
            SkewKurtCorradoSuModified = CallValue
        Else '// Use put-call parity to find put value
            SkewKurtCorradoSuModified = CallValue - S * Exp((b - r) * T) + X * Exp(-r * T)
        End If
End Function



Public Function ESkewKurtCorradoSuModified(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, Skew As Double, Kurt As Double, Optional dS)
            
  
    
    If IsMissing(dS) Then
        dS = 0.1
    End If
    
    
    
    If OutPutFlag = "p" Then ' Value
        ESkewKurtCorradoSuModified = SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
    ElseIf OutPutFlag = "d" Then 'Delta
         ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / (2 * dS) * S / SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - 2 * SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + SkewKurtCorradoSuModified(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S + dS, X, T, r, b, v + 0.01, Skew, Kurt) - 2 * SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) + SkewKurtCorradoSuModified(CallPutFlag, S - dS, X, T, r, b, v + 0.01, Skew, Kurt) _
      - SkewKurtCorradoSuModified(CallPutFlag, S + dS, X, T, r, b, v - 0.01, Skew, Kurt) + 2 * SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S - dS, X, T, r, b, v - 0.01, Skew, Kurt)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        ESkewKurtCorradoSuModified = S / 100 * (SkewKurtCorradoSuModified(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - 2 * SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + SkewKurtCorradoSuModified(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / dS ^ 2
    ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ESkewKurtCorradoSuModified = 1 / (4 * dS * 0.01) * (SkewKurtCorradoSuModified(CallPutFlag, S + dS, X, T, r, b, v + 0.01, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S + dS, X, T, r, b, v - 0.01, Skew, Kurt) _
        - SkewKurtCorradoSuModified(CallPutFlag, S - dS, X, T, r, b, v + 0.01, Skew, Kurt) + SkewKurtCorradoSuModified(CallPutFlag, S - dS, X, T, r, b, v - 0.01, Skew, Kurt)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - 2 * SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ESkewKurtCorradoSuModified = v / 0.1 * (SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - 2 * SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ESkewKurtCorradoSuModified = SkewKurtCorradoSuModified(CallPutFlag, S, X, 0.00001, r, b, v, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
        Else
                ESkewKurtCorradoSuModified = SkewKurtCorradoSuModified(CallPutFlag, S, X, T - 1 / 365, r, b, v, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v, Skew, Kurt)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r + 0.01, b, v, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r - 0.01, b, v, Skew, Kurt)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b - 0.01, v, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b + 0.01, v, Skew, Kurt)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b + 0.01, v, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b - 0.01, v, Skew, Kurt)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ESkewKurtCorradoSuModified = 1 / dS ^ 3 * (SkewKurtCorradoSuModified(CallPutFlag, S + 2 * dS, X, T, r, b, v, Skew, Kurt) - 3 * SkewKurtCorradoSuModified(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) _
            + 3 * SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S, X + dS, T, r, b, v, Skew, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S, X - dS, T, r, b, v, Skew, Kurt)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S, X + dS, T, r, b, v, Skew, Kurt) - 2 * SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + SkewKurtCorradoSuModified(CallPutFlag, S, X - dS, T, r, b, v, Skew, Kurt)) / dS ^ 2
       ElseIf OutPutFlag = "skew" Then 'Skewness delta
         ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew + 0.01, Kurt) - SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew - 0.01, Kurt)) / 2
    ElseIf OutPutFlag = "kurt" Then 'Kurt delta
         ESkewKurtCorradoSuModified = (SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt + 0.01) - SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt - 0.01)) / 2
    ElseIf OutPutFlag = "diff" Then 'SkewKurt Corrado-Su model minus Black-Scholes-Merton
         ESkewKurtCorradoSuModified = SkewKurtCorradoSuModified(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) - GBlackScholes(CallPutFlag, S, X, T, r, b, v)
   
    End If

End Function

'Returns Skewness for Log-normal distribution (price distribution, not-return)
Public Function LogNormalSkew(v As Double, T As Double) As Double
    
    Dim y As Double
    
    y = Sqr(Exp(v ^ 2 * T) - 1)
    LogNormalSkew = 3 * y + y ^ 3
    
End Function


'Returns Fisher kurtosis for Log-normal distribution (price distribution, not-return)
Public Function LogNormalKurt(v As Double, T As Double) As Double
    
    Dim y As Double
    
    y = Sqr(Exp(v ^ 2 * T) - 1)
    LogNormalKurt = 16 * y ^ 2 + 15 * y ^ 4 + 6 * y ^ 6 + y ^ 8
    
End Function


Public Function SkewKurtCorradoSu(CallPutFlag As String, S As Double, X _
                As Double, T As Double, r As Double, b As Double, v As Double, Skew As Double, Kurt As Double) As Double

        Dim Q3 As Double, Q4 As Double
        Dim d1 As Double, d2 As Double
        Dim CallValue As Double
      
        
        d1 = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
        d2 = d1 - v * Sqr(T)
        Q4 = 1 / 24 * S * v * Sqr(T) * ((d1 ^ 2 - 1 - 3 * v * Sqr(T) * d2) * ND(d1) + v ^ 3 * T ^ 1.5 * CND(d1))
        Q3 = 1 / 6 * S * v * Sqr(T) * ((2 * v * Sqr(T) - d1) * ND(d1) + v ^ 2 * T * CND(d1))

        CallValue = GBlackScholes("c", S, X, T, r, b, v) + Skew * Q3 + (Kurt - 3) * Q4
        
         If CallPutFlag = "c" Then
            SkewKurtCorradoSu = CallValue
        Else '// Use put-call parity to find put value
            SkewKurtCorradoSu = CallValue - S * Exp((b - r) * T) + X * Exp(-r * T)
        End If
        
End Function





Public Function ESkewKurtCorradoSu(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, Skew As Double, Kurt As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        ESkewKurtCorradoSu = SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
    ElseIf OutPutFlag = "d" Then 'Delta
         ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / (2 * dS) * S / SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - 2 * SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + SkewKurtCorradoSu(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S + dS, X, T, r, b, v + 0.01, Skew, Kurt) - 2 * SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) + SkewKurtCorradoSu(CallPutFlag, S - dS, X, T, r, b, v + 0.01, Skew, Kurt) _
      - SkewKurtCorradoSu(CallPutFlag, S + dS, X, T, r, b, v - 0.01, Skew, Kurt) + 2 * SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S - dS, X, T, r, b, v - 0.01, Skew, Kurt)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        ESkewKurtCorradoSu = S / 100 * (SkewKurtCorradoSu(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - 2 * SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + SkewKurtCorradoSu(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / dS ^ 2
    ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ESkewKurtCorradoSu = 1 / (4 * dS * 0.01) * (SkewKurtCorradoSu(CallPutFlag, S + dS, X, T, r, b, v + 0.01, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S + dS, X, T, r, b, v - 0.01, Skew, Kurt) _
        - SkewKurtCorradoSu(CallPutFlag, S - dS, X, T, r, b, v + 0.01, Skew, Kurt) + SkewKurtCorradoSu(CallPutFlag, S - dS, X, T, r, b, v - 0.01, Skew, Kurt)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - 2 * SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ESkewKurtCorradoSu = v / 0.1 * (SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - 2 * SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ESkewKurtCorradoSu = SkewKurtCorradoSu(CallPutFlag, S, X, 0.00001, r, b, v, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
        Else
                ESkewKurtCorradoSu = SkewKurtCorradoSu(CallPutFlag, S, X, T - 1 / 365, r, b, v, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v, Skew, Kurt)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S, X, T, r + 0.01, b, v, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S, X, T, r - 0.01, b, v, Skew, Kurt)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b - 0.01, v, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b + 0.01, v, Skew, Kurt)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b + 0.01, v, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b - 0.01, v, Skew, Kurt)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ESkewKurtCorradoSu = 1 / dS ^ 3 * (SkewKurtCorradoSu(CallPutFlag, S + 2 * dS, X, T, r, b, v, Skew, Kurt) - 3 * SkewKurtCorradoSu(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) _
                                + 3 * SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S, X + dS, T, r, b, v, Skew, Kurt) - SkewKurtCorradoSu(CallPutFlag, S, X - dS, T, r, b, v, Skew, Kurt)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S, X + dS, T, r, b, v, Skew, Kurt) - 2 * SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + SkewKurtCorradoSu(CallPutFlag, S, X - dS, T, r, b, v, Skew, Kurt)) / dS ^ 2
       ElseIf OutPutFlag = "skew" Then 'Skewness delta
         ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew + 0.01, Kurt) - SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew - 0.01, Kurt)) / 2
    ElseIf OutPutFlag = "kurt" Then 'Kurt delta
         ESkewKurtCorradoSu = (SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt + 0.01) - SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt - 0.01)) / 2
    ElseIf OutPutFlag = "diff" Then 'SkewKurt Corrado-Su model minus Black-Scholes-Merton
         ESkewKurtCorradoSu = SkewKurtCorradoSu(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) - GBlackScholes(CallPutFlag, S, X, T, r, b, v)
   
    End If

End Function


Public Function JarrowRuddSkewKurt(CallPutFlag As String, S As Double, X _
                As Double, T As Double, r As Double, b As Double, v As Double, Skew As Double, Kurt As Double) As Double


        Dim Q3 As Double, Q4 As Double
        Dim d1 As Double, d2 As Double
        Dim CallValue As Double
        
        Dim aX As Double, daX As Double, daXX As Double
        Dim q As Double, GA As Double, gAA As Double
        Dim Lambda1 As Double, Lambda2 As Double
        
        d1 = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
        d2 = d1 - v * Sqr(T)
        
        aX = (X * v * Sqr(T * 2 * Pi)) ^ (-1) * Exp(-d2 ^ 2 / 2)
        daX = aX * (d2 - v * Sqr(T)) / (X * v * Sqr(T))
        daXX = aX / (X ^ 2 * v * Sqr(T)) * ((d2 - v * Sqr(T)) ^ 2 - v * Sqr(T) * (d2 - v * Sqr(T)) - 1)
        
         q = Sqr(Exp(v ^ 2 * T) - 1)

         GA = 3 * q + q ^ 3
         gAA = 16 * q ^ 2 + 15 * q ^ 4 + 6 * q ^ 6 + q ^ 8 + 3
         
         Lambda1 = Skew - GA
         Lambda2 = Kurt - gAA
        
        Q3 = -(S * Exp(r * T)) ^ 3 * (Exp(v ^ 2 * T) - 1) ^ (3 / 2) * Exp(-r * T) / 6 * daX
        Q4 = (S * Exp(r * T)) ^ 4 * (Exp(v ^ 2 * T) - 1) ^ 2 * Exp(-r * T) / 24 * daXX
        
        CallValue = (GBlackScholes("c", S, X, T, r, b, v) + Lambda1 * Q3 + Lambda2 * Q4)
         If CallPutFlag = "c" Then
            JarrowRuddSkewKurt = CallValue
        Else '// Use put-call parity to find put value
            JarrowRuddSkewKurt = CallValue - S * Exp((b - r) * T) + X * Exp(-r * T)
        End If
End Function


Public Function EJarrowRuddSkewKurt(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, Skew As Double, Kurt As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EJarrowRuddSkewKurt = JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
    ElseIf OutPutFlag = "d" Then 'Delta
         EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / (2 * dS) * S / JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - 2 * JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + JarrowRuddSkewKurt(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S + dS, X, T, r, b, v + 0.01, Skew, Kurt) - 2 * JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) + JarrowRuddSkewKurt(CallPutFlag, S - dS, X, T, r, b, v + 0.01, Skew, Kurt) _
      - JarrowRuddSkewKurt(CallPutFlag, S + dS, X, T, r, b, v - 0.01, Skew, Kurt) + 2 * JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S - dS, X, T, r, b, v - 0.01, Skew, Kurt)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EJarrowRuddSkewKurt = S / 100 * (JarrowRuddSkewKurt(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) - 2 * JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + JarrowRuddSkewKurt(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt)) / dS ^ 2
    ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EJarrowRuddSkewKurt = 1 / (4 * dS * 0.01) * (JarrowRuddSkewKurt(CallPutFlag, S + dS, X, T, r, b, v + 0.01, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S + dS, X, T, r, b, v - 0.01, Skew, Kurt) _
        - JarrowRuddSkewKurt(CallPutFlag, S - dS, X, T, r, b, v + 0.01, Skew, Kurt) + JarrowRuddSkewKurt(CallPutFlag, S - dS, X, T, r, b, v - 0.01, Skew, Kurt)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - 2 * JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EJarrowRuddSkewKurt = v / 0.1 * (JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v + 0.01, Skew, Kurt) - 2 * JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v - 0.01, Skew, Kurt))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EJarrowRuddSkewKurt = JarrowRuddSkewKurt(CallPutFlag, S, X, 0.00001, r, b, v, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
        Else
                EJarrowRuddSkewKurt = JarrowRuddSkewKurt(CallPutFlag, S, X, T - 1 / 365, r, b, v, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v, Skew, Kurt)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S, X, T, r + 0.01, b, v, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S, X, T, r - 0.01, b, v, Skew, Kurt)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b - 0.01, v, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b + 0.01, v, Skew, Kurt)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b + 0.01, v, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b - 0.01, v, Skew, Kurt)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EJarrowRuddSkewKurt = 1 / dS ^ 3 * (JarrowRuddSkewKurt(CallPutFlag, S + 2 * dS, X, T, r, b, v, Skew, Kurt) - 3 * JarrowRuddSkewKurt(CallPutFlag, S + dS, X, T, r, b, v, Skew, Kurt) _
                                + 3 * JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S - dS, X, T, r, b, v, Skew, Kurt))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S, X + dS, T, r, b, v, Skew, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S, X - dS, T, r, b, v, Skew, Kurt)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S, X + dS, T, r, b, v, Skew, Kurt) - 2 * JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) + JarrowRuddSkewKurt(CallPutFlag, S, X - dS, T, r, b, v, Skew, Kurt)) / dS ^ 2
       ElseIf OutPutFlag = "skew" Then 'Skewness delta
         EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew + 0.01, Kurt) - JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew - 0.01, Kurt)) / 2
    ElseIf OutPutFlag = "kurt" Then 'Kurt delta
         EJarrowRuddSkewKurt = (JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt + 0.01) - JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt - 0.01)) / 2
    ElseIf OutPutFlag = "diff" Then 'SkewKurt Corrado-Su model minus Black-Scholes-Merton
         EJarrowRuddSkewKurt = JarrowRuddSkewKurt(CallPutFlag, S, X, T, r, b, v, Skew, Kurt) - GBlackScholes(CallPutFlag, S, X, T, r, b, v)
   
    End If

End Function



