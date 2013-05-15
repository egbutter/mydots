Attribute VB_Name = "SABR"
Option Explicit

Public Function SABRModel(CallPutFlag As String, F As Double, X As Double, T As Double, r As Double, ATMvol As Double, Beta As Double, VolVol As Double, rho As Double) As Double

    Dim v As Double, d1 As Double, d2 As Double
    
    v = SABRVolatility(F, X, T, ATMvol, Beta, VolVol, rho)
    
    d1 = (Log(F / X) + v ^ 2 / 2 * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    
    If CallPutFlag = "c" Then
        SABRModel = Exp(-r * T) * (F * CND(d1) - X * CND(d2))
    ElseIf CallPutFlag = "p" Then
        SABRModel = Exp(-r * T) * (X * CND(-d2) - F * CND(-d1))
    End If

End Function


Public Function FindAlpha(F As Double, X As Double, T As Double, ATMvol As Double, Beta As Double, VolVol As Double, rho As Double) As Double

  'alpha is a function of atmvol etc.

  FindAlpha = CRoot((1 - Beta) ^ 2 * T / (24 * F ^ (2 - 2 * Beta)), 0.25 * rho * VolVol * Beta * T / F ^ (1 - Beta), 1 + (2 - 3 * rho ^ 2) / 24 * VolVol ^ 2 * T, -ATMvol * F ^ (1 - Beta))

End Function

Public Function SABRVolatility(F As Double, X As Double, T As Double, ATMvol As Double, Beta As Double, VolVol As Double, rho As Double) As Double
    SABRVolatility = alphaSABR(F, X, T, FindAlpha(F, X, T, ATMvol, Beta, VolVol, rho), Beta, VolVol, rho)
End Function

Public Function alphaSABR(F As Double, X As Double, T As Double, Alpha As Double, Beta As Double, VolVol As Double, rho As Double) As Double

  'the SABR skew vol function
  
  Dim dSABR(1 To 3) As Double
  Dim sabrz As Double, y As Double

  dSABR(1) = Alpha / ((F * X) ^ ((1 - Beta) / 2) * (1 + (((1 - Beta) ^ 2) / 24) * (Log(F / X) ^ 2) + ((1 - Beta) ^ 4 / 1920) * (Log(F / X) ^ 4)))

  
  If Abs(F - X) > 10 ^ -8 Then
    sabrz = (VolVol / Alpha) * (F * X) ^ ((1 - Beta) / 2) * Log(F / X)
    y = (Sqr(1 - 2 * rho * sabrz + sabrz ^ 2) + sabrz - rho) / (1 - rho)
    If Abs(y - 1) < 10 ^ -8 Then
      dSABR(2) = 1
    ElseIf y > 0 Then
      dSABR(2) = sabrz / Log(y)
    Else
      dSABR(2) = 1
    End If
  Else
    dSABR(2) = 1
  End If

  dSABR(3) = 1 + ((((1 - Beta) ^ 2 / 24) * Alpha ^ 2 / ((F * X) ^ (1 - Beta))) + 0.25 * rho * Beta * VolVol * Alpha / ((F * X) ^ ((1 - Beta) / 2)) + (2 - 3 * rho ^ 2) * VolVol ^ 2 / 24) * T
  alphaSABR = dSABR(1) * dSABR(2) * dSABR(3)

End Function

Public Function CRoot(cubic As Double, quadratic As Double, linear As Double, constant As Double) As Double

  'finds the smallest postive root of the input cubic polynomial
  'algorithm from Numerical Recipes
  
  Dim roots(1 To 3) As Double
  Dim A As Double, b As Double, c As Double
  Dim r As Double, q As Double
  Dim capA As Double, capB As Double, theta As Double
  

  A = quadratic / cubic
  b = linear / cubic
  c = constant / cubic
  q = (A ^ 2 - 3 * b) / 9
  r = (2 * A ^ 3 - 9 * A * b + 27 * c) / 54
  
  If r ^ 2 - q ^ 3 >= 0 Then
    capA = -Sgn(r) * (Abs(r) + Sqr(r ^ 2 - q ^ 3)) ^ (1 / 3)
    If capA = 0 Then capB = 0 Else capB = q / capA
    CRoot = capA + capB - A / 3
  Else

    theta = ArcCos(r / q ^ 1.5)

    'the three roots

    roots(1) = -2 * Sqr(q) * Cos(theta / 3) - A / 3
    roots(2) = -2 * Sqr(q) * Cos(theta / 3 + 2.0943951023932) - A / 3
    roots(3) = -2 * Sqr(q) * Cos(theta / 3 - 2.0943951023932) - A / 3

    'locate that one which is the smallest positive root
    'assumes there is such a root (true for SABR model)
    'there is always a small positive root

    If roots(1) > 0 Then
      CRoot = roots(1)
    ElseIf roots(2) > 0 Then
      CRoot = roots(2)
    ElseIf roots(3) > 0 Then
      CRoot = roots(3)
    End If

    If roots(2) > 0 And roots(2) < CRoot Then
      CRoot = roots(2)
    End If
    If roots(3) > 0 And roots(3) < CRoot Then
      CRoot = roots(3)
    End If
  End If
End Function



Public Function ArcCos(y As Double) As Double
  ArcCos = Atn(-y / Sqr(-y * y + 1)) + 2 * Atn(1)
End Function


Public Function ESABRModel(OutPutFlag As String, CallPutFlag As String, F As Double, X As Double, T As Double, _
                r As Double, ATMvol As Double, Beta As Double, VolVol As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        ESABRModel = SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho)
    ElseIf OutPutFlag = "d" Then 'Delta
         ESABRModel = (SABRModel(CallPutFlag, F + dS, X, T, r, ATMvol, Beta, VolVol, rho) - SABRModel(CallPutFlag, F - dS, X, T, r, ATMvol, Beta, VolVol, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         ESABRModel = (SABRModel(CallPutFlag, F + dS, X, T, r, ATMvol, Beta, VolVol, rho) - SABRModel(CallPutFlag, F - dS, X, T, r, ATMvol, Beta, VolVol, rho)) / (2 * dS) * F / SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ESABRModel = (SABRModel(CallPutFlag, F + dS, X, T, r, ATMvol, Beta, VolVol, rho) - 2 * SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho) + SABRModel(CallPutFlag, F - dS, X, T, r, ATMvol, Beta, VolVol, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ESABRModel = (SABRModel(CallPutFlag, F + dS, X, T, r, ATMvol + 0.01, Beta, VolVol, rho) - 2 * SABRModel(CallPutFlag, F, X, T, r, ATMvol + 0.01, Beta, VolVol, rho) + SABRModel(CallPutFlag, F - dS, X, T, r, ATMvol + 0.01, Beta, VolVol, rho) _
      - SABRModel(CallPutFlag, F + dS, X, T, r, ATMvol - 0.01, Beta, VolVol, rho) + 2 * SABRModel(CallPutFlag, F, X, T, r, ATMvol - 0.01, Beta, VolVol, rho) - SABRModel(CallPutFlag, F - dS, X, T, r, ATMvol - 0.01, Beta, VolVol, rho)) / (2 * 0.01 * dS ^ 2) / 100
     ElseIf OutPutFlag = "grho" Then 'DGammaDCorr
        ESABRModel = (SABRModel(CallPutFlag, F + dS, X, T, r, ATMvol, Beta, VolVol, rho + 0.01) - 2 * SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho + 0.01) + SABRModel(CallPutFlag, F - dS, X, T, r, ATMvol, Beta, VolVol, rho + 0.01) _
      - SABRModel(CallPutFlag, F + dS, X, T, r, ATMvol, Beta, VolVol, rho - 0.01) + 2 * SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho - 0.01) - SABRModel(CallPutFlag, F - dS, X, T, r, ATMvol, Beta, VolVol, rho - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
 
    ElseIf OutPutFlag = "gp" Then 'GammaP
        ESABRModel = F / 100 * (SABRModel(CallPutFlag, F + dS, X, T, r, ATMvol, Beta, VolVol, rho) - 2 * SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho) + SABRModel(CallPutFlag, F - dS, X, T, r, ATMvol, Beta, VolVol, rho)) / dS ^ 2
    ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ESABRModel = 1 / (4 * dS * 0.01) * (SABRModel(CallPutFlag, F + dS, X, T, r, ATMvol + 0.01, Beta, VolVol, rho) - SABRModel(CallPutFlag, F + dS, X, T, r, ATMvol - 0.01, Beta, VolVol, rho) _
        - SABRModel(CallPutFlag, F - dS, X, T, r, ATMvol + 0.01, Beta, VolVol, rho) + SABRModel(CallPutFlag, F - dS, X, T, r, ATMvol - 0.01, Beta, VolVol, rho)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ESABRModel = (SABRModel(CallPutFlag, F, X, T, r, ATMvol + 0.01, Beta, VolVol, rho) - SABRModel(CallPutFlag, F, X, T, r, ATMvol - 0.01, Beta, VolVol, rho)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        ESABRModel = (SABRModel(CallPutFlag, F, X, T, r, ATMvol + 0.01, Beta, VolVol, rho) - 2 * SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho) + SABRModel(CallPutFlag, F, X, T, r, ATMvol - 0.01, Beta, VolVol, rho)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ESABRModel = ATMvol / 0.1 * (SABRModel(CallPutFlag, F, X, T, r, ATMvol + 0.01, Beta, VolVol, rho) - SABRModel(CallPutFlag, F, X, T, r, ATMvol - 0.01, Beta, VolVol, rho)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ESABRModel = (SABRModel(CallPutFlag, F, X, T, r, ATMvol + 0.01, Beta, VolVol, rho) - 2 * SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho) + SABRModel(CallPutFlag, F, X, T, r, ATMvol - 0.01, Beta, VolVol, rho))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ESABRModel = SABRModel(CallPutFlag, F, X, 0.00001, r, ATMvol, Beta, VolVol, rho) - SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho)
        Else
                ESABRModel = SABRModel(CallPutFlag, F, X, T - 1 / 365, r, ATMvol, Beta, VolVol, rho) - SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho)
        End If
      ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ESABRModel = (SABRModel(CallPutFlag, F, X, T, r + 0.01, ATMvol, Beta, VolVol, rho) - SABRModel(CallPutFlag, F, X, T, r - 0.01, ATMvol, Beta, VolVol, rho)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ESABRModel = 1 / dS ^ 3 * (SABRModel(CallPutFlag, F + 2 * dS, X, T, r, ATMvol, Beta, VolVol, rho) - 3 * SABRModel(CallPutFlag, F + dS, X, T, r, ATMvol, Beta, VolVol, rho) _
                                + 3 * SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho) - SABRModel(CallPutFlag, F - dS, X, T, r, ATMvol, Beta, VolVol, rho))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ESABRModel = (SABRModel(CallPutFlag, F, X + dS, T, r, ATMvol, Beta, VolVol, rho) - SABRModel(CallPutFlag, F, X - dS, T, r, ATMvol, Beta, VolVol, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        ESABRModel = (SABRModel(CallPutFlag, F, X + dS, T, r, ATMvol, Beta, VolVol, rho) - 2 * SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho) + SABRModel(CallPutFlag, F, X - dS, T, r, ATMvol, Beta, VolVol, rho)) / dS ^ 2
    ElseIf OutPutFlag = "vvVega" Then 'VolVol Vega
         ESABRModel = (SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol + 0.01, rho) - SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol - 0.01, rho)) / 2
      ElseIf OutPutFlag = "vvVomma" Then ' VolVol vomma
        ESABRModel = (SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol + 0.01, rho) - 2 * SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho) + SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol - 0.01, rho)) / 0.01 ^ 2 / 10000
     ElseIf OutPutFlag = "corr" Then 'Correlation sensitivity
         ESABRModel = (SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho + 0.01) - SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho - 0.01)) / 2
    ElseIf OutPutFlag = "corrcorr" Then ' correaltion vomma
        ESABRModel = (SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho + 0.01) - 2 * SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho) + SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho - 0.01)) / 0.01 ^ 2 / 10000
   ElseIf OutPutFlag = "diff" Then 'HullWhite87 minus Black-Scholes-Merton
         ESABRModel = SABRModel(CallPutFlag, F, X, T, r, ATMvol, Beta, VolVol, rho) - GBlackScholes(CallPutFlag, F, X, T, r, 0, ATMvol)
 
    End If

End Function


