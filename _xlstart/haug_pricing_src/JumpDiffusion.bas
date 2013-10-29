Attribute VB_Name = "JumpDiffusion"
Option Explicit



'// Bates (1991) generalized jump diffusion model
Public Function JumpDiffusionBates(CallPutFlag As String, S As Double, x As Double, T As Double, _
    r As Double, b As Double, v As Double, lambda As Double, avgk As Double, delta As Double) As Double

    Dim sum As Double, gam0 As Double
    Dim bi As Double, vi As Double
    Dim i As Long

    gam0 = Log(1 + avgk)
    sum = 0
    For i = 0 To 50
        bi = b - lambda * avgk + gam0 * (i / T)
        vi = Sqr(v ^ 2 + delta ^ 2 * (i / T))
        sum = sum + Exp(-lambda * T) * (lambda * T) ^ i / Application.Fact(i) * _
        GBlackScholes(CallPutFlag, S, x, T, r, bi, vi)
    Next
        JumpDiffusionBates = sum
        
End Function


Public Function EJumpDiffusionBates(OutPutFlag As String, CallPutFlag As String, S As Double, x As Double, T As Double, _
                r As Double, b As Double, v As Double, lambda As Double, avgk As Double, delta As Double, Optional dS)
                    
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        EJumpDiffusionBates = JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta)
    ElseIf OutPutFlag = "d" Then 'Delta
         EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S + dS, x, T, r, b, v, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S - dS, x, T, r, b, v, lambda, avgk, delta)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S + dS, x, T, r, b, v, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S - dS, x, T, r, b, v, lambda, avgk, delta)) / (2 * dS) * S / JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S + dS, x, T, r, b, v, lambda, avgk, delta) - 2 * JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta) + JumpDiffusionBates(CallPutFlag, S - dS, x, T, r, b, v, lambda, avgk, delta)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S + dS, x, T, r, b, v + 0.01, lambda, avgk, delta) - 2 * JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v + 0.01, lambda, avgk, delta) + JumpDiffusionBates(CallPutFlag, S - dS, x, T, r, b, v + 0.01, lambda, avgk, delta) _
      - JumpDiffusionBates(CallPutFlag, S + dS, x, T, r, b, v - 0.01, lambda, avgk, delta) + 2 * JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v - 0.01, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S - dS, x, T, r, b, v - 0.01, lambda, avgk, delta)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EJumpDiffusionBates = S / 100 * (JumpDiffusionBates(CallPutFlag, S + dS, x, T, r, b, v, lambda, avgk, delta) - 2 * JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta) + JumpDiffusionBates(CallPutFlag, S - dS, x, T, r, b, v, lambda, avgk, delta)) / dS ^ 2
    ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EJumpDiffusionBates = 1 / (4 * dS * 0.01) * (JumpDiffusionBates(CallPutFlag, S + dS, x, T, r, b, v + 0.01, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S + dS, x, T, r, b, v - 0.01, lambda, avgk, delta) _
        - JumpDiffusionBates(CallPutFlag, S - dS, x, T, r, b, v + 0.01, lambda, avgk, delta) + JumpDiffusionBates(CallPutFlag, S - dS, x, T, r, b, v - 0.01, lambda, avgk, delta)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v + 0.01, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v - 0.01, lambda, avgk, delta)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v + 0.01, lambda, avgk, delta) - 2 * JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta) + JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v - 0.01, lambda, avgk, delta)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EJumpDiffusionBates = v / 0.1 * (JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v + 0.01, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v - 0.01, lambda, avgk, delta)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v + 0.01, lambda, avgk, delta) - 2 * JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta) + JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v - 0.01, lambda, avgk, delta))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EJumpDiffusionBates = JumpDiffusionBates(CallPutFlag, S, x, 0.00001, r, b, v, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta)
        Else
                EJumpDiffusionBates = JumpDiffusionBates(CallPutFlag, S, x, T - 1 / 365, r, b, v, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S, x, T, r + 0.01, b + 0.01, v, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S, x, T, r - 0.01, b - 0.01, v, lambda, avgk, delta)) / (2)
     ElseIf OutPutFlag = "fr" Then 'Rho futures option
         EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S, x, T, r + 0.01, b, v, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S, x, T, r - 0.01, b, v, lambda, avgk, delta)) / (2)
    ElseIf OutPutFlag = "f" Then 'Phi/Rho2
         EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S, x, T, r, b - 0.01, v, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S, x, T, r, b + 0.01, v, lambda, avgk, delta)) / (2)
     ElseIf OutPutFlag = "b" Then 'Carry
         EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S, x, T, r, b + 0.01, v, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S, x, T, r, b - 0.01, v, lambda, avgk, delta)) / (2)
   ElseIf OutPutFlag = "s" Then 'Speed
        EJumpDiffusionBates = 1 / dS ^ 3 * (JumpDiffusionBates(CallPutFlag, S + 2 * dS, x, T, r, b, v, lambda, avgk, delta) - 3 * JumpDiffusionBates(CallPutFlag, S + dS, x, T, r, b, v, lambda, avgk, delta) _
                                + 3 * JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S - dS, x, T, r, b, v, lambda, avgk, delta))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S, x + dS, T, r, b, v, lambda, avgk, delta) - JumpDiffusionBates(CallPutFlag, S, x - dS, T, r, b, v, lambda, avgk, delta)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S, x + dS, T, r, b, v, lambda, avgk, delta) - 2 * JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta) + JumpDiffusionBates(CallPutFlag, S, x - dS, T, r, b, v, lambda, avgk, delta)) / dS ^ 2
       ElseIf OutPutFlag = "lambda" Then 'lambd delta
         EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda + 0.1, avgk, delta) - JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda - 0.1, avgk, delta)) / (0.1 * 2)
    ElseIf OutPutFlag = "JumpStd" Then 'gamma delta
         EJumpDiffusionBates = (JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta + 0.01) - JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta - 0.01)) / 2
    ElseIf OutPutFlag = "diff" Then 'Merton Jump-Diffusion minus Black-Scholes-Merton
        Dim gammabar As Double, vTotal As Double
        gammabar = avgk - delta ^ 2 / 2
        vTotal = Sqr(v ^ 2 + lambda * (gammabar ^ 2 + delta ^ 2))
    
         EJumpDiffusionBates = JumpDiffusionBates(CallPutFlag, S, x, T, r, b, v, lambda, avgk, delta) - GBlackScholes(CallPutFlag, S, x, T, r, b, vTotal)
   
    End If

End Function

'// Merton (1976) jump diffusion model
Public Function JumpDiffusionMerton(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, v As Double, _
                lambda As Double, gamma As Double) As Double

    Dim delta As Double, sum As Double
    Dim Z As Double, vi As Double
    Dim i As Integer

    delta = Sqr(gamma * v ^ 2 / lambda)
    Z = Sqr(v ^ 2 - lambda * delta ^ 2)
    sum = 0
    For i = 0 To 50
        vi = Sqr(Z ^ 2 + delta ^ 2 * (i / T))
        sum = sum + Exp(-lambda * T) * (lambda * T) ^ i / Application.Fact(i) * _
        GBlackScholes(CallPutFlag, S, x, T, r, r, vi)
    Next
        JumpDiffusionMerton = sum
End Function



Public Function EJumpDiffusionMerton(OutPutFlag As String, CallPutFlag As String, S As Double, x As Double, T As Double, _
                r As Double, v As Double, lambda As Double, gamma As Double, Optional dS)
            
        Dim b As Double
        b = 0
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EJumpDiffusionMerton = JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma)
    ElseIf OutPutFlag = "d" Then 'Delta
         EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S + dS, x, T, r, v, lambda, gamma) - JumpDiffusionMerton(CallPutFlag, S - dS, x, T, r, v, lambda, gamma)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S + dS, x, T, r, v, lambda, gamma) - JumpDiffusionMerton(CallPutFlag, S - dS, x, T, r, v, lambda, gamma)) / (2 * dS) * S / JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S + dS, x, T, r, v, lambda, gamma) - 2 * JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma) + JumpDiffusionMerton(CallPutFlag, S - dS, x, T, r, v, lambda, gamma)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S + dS, x, T, r, v + 0.01, lambda, gamma) - 2 * JumpDiffusionMerton(CallPutFlag, S, x, T, r, v + 0.01, lambda, gamma) + JumpDiffusionMerton(CallPutFlag, S - dS, x, T, r, v + 0.01, lambda, gamma) _
      - JumpDiffusionMerton(CallPutFlag, S + dS, x, T, r, v - 0.01, lambda, gamma) + 2 * JumpDiffusionMerton(CallPutFlag, S, x, T, r, v - 0.01, lambda, gamma) - JumpDiffusionMerton(CallPutFlag, S - dS, x, T, r, v - 0.01, lambda, gamma)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EJumpDiffusionMerton = S / 100 * (JumpDiffusionMerton(CallPutFlag, S + dS, x, T, r, v, lambda, gamma) - 2 * JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma) + JumpDiffusionMerton(CallPutFlag, S - dS, x, T, r, v, lambda, gamma)) / dS ^ 2
    ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EJumpDiffusionMerton = 1 / (4 * dS * 0.01) * (JumpDiffusionMerton(CallPutFlag, S + dS, x, T, r, v + 0.01, lambda, gamma) - JumpDiffusionMerton(CallPutFlag, S + dS, x, T, r, v - 0.01, lambda, gamma) _
        - JumpDiffusionMerton(CallPutFlag, S - dS, x, T, r, v + 0.01, lambda, gamma) + JumpDiffusionMerton(CallPutFlag, S - dS, x, T, r, v - 0.01, lambda, gamma)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S, x, T, r, v + 0.01, lambda, gamma) - JumpDiffusionMerton(CallPutFlag, S, x, T, r, v - 0.01, lambda, gamma)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S, x, T, r, v + 0.01, lambda, gamma) - 2 * JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma) + JumpDiffusionMerton(CallPutFlag, S, x, T, r, v - 0.01, lambda, gamma)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EJumpDiffusionMerton = v / 0.1 * (JumpDiffusionMerton(CallPutFlag, S, x, T, r, v + 0.01, lambda, gamma) - JumpDiffusionMerton(CallPutFlag, S, x, T, r, v - 0.01, lambda, gamma)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S, x, T, r, v + 0.01, lambda, gamma) - 2 * JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma) + JumpDiffusionMerton(CallPutFlag, S, x, T, r, v - 0.01, lambda, gamma))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EJumpDiffusionMerton = JumpDiffusionMerton(CallPutFlag, S, x, 0.00001, r, v, lambda, gamma) - JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma)
        Else
                EJumpDiffusionMerton = JumpDiffusionMerton(CallPutFlag, S, x, T - 1 / 365, r, v, lambda, gamma) - JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S, x, T, r + 0.01, v, lambda, gamma) - JumpDiffusionMerton(CallPutFlag, S, x, T, r - 0.01, v, lambda, gamma)) / (2)
   ElseIf OutPutFlag = "s" Then 'Speed
        EJumpDiffusionMerton = 1 / dS ^ 3 * (JumpDiffusionMerton(CallPutFlag, S + 2 * dS, x, T, r, v, lambda, gamma) - 3 * JumpDiffusionMerton(CallPutFlag, S + dS, x, T, r, v, lambda, gamma) _
                                + 3 * JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma) - JumpDiffusionMerton(CallPutFlag, S - dS, x, T, r, v, lambda, gamma))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S, x + dS, T, r, v, lambda, gamma) - JumpDiffusionMerton(CallPutFlag, S, x - dS, T, r, v, lambda, gamma)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S, x + dS, T, r, v, lambda, gamma) - 2 * JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma) + JumpDiffusionMerton(CallPutFlag, S, x - dS, T, r, v, lambda, gamma)) / dS ^ 2
       ElseIf OutPutFlag = "lambda" Then 'lambd delta
         EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda + 0.1, gamma) - JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda - 0.1, gamma)) / (0.1 * 2)
    ElseIf OutPutFlag = "gamma" Then 'gamma delta
         EJumpDiffusionMerton = (JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma + 0.01) - JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma - 0.01)) / 2
    ElseIf OutPutFlag = "diff" Then 'Merton Jump-Diffusion minus Black-Scholes-Merton
         EJumpDiffusionMerton = JumpDiffusionMerton(CallPutFlag, S, x, T, r, v, lambda, gamma) - GBlackScholes(CallPutFlag, S, x, T, r, r, v)
   
    End If

End Function

