Attribute VB_Name = "TwoAssetDigital"
Option Explicit


' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug

'// Two asset cash-or-nothing options
Public Function TwoAssetCashOrNothing(TypeFlag As Integer, S1 As Double, S2 As Double, X1 As Double, X2 As Double, k As Double, T As Double, r As Double, _
                b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double) As Double
    
    Dim d1 As Double, d2 As Double
                                   
    d1 = (Log(S1 / X1) + (b1 - v1 ^ 2 / 2) * T) / (v1 * Sqr(T))
    d2 = (Log(S2 / X2) + (b2 - v2 ^ 2 / 2) * T) / (v2 * Sqr(T))
                                
    If TypeFlag = 1 Then
        TwoAssetCashOrNothing = k * Exp(-r * T) * CBND(d1, d2, rho)
    ElseIf TypeFlag = 2 Then
        TwoAssetCashOrNothing = k * Exp(-r * T) * CBND(-d1, -d2, rho)
    ElseIf TypeFlag = 3 Then
        TwoAssetCashOrNothing = k * Exp(-r * T) * CBND(d1, -d2, -rho)
    ElseIf TypeFlag = 4 Then
        TwoAssetCashOrNothing = k * Exp(-r * T) * CBND(-d1, d2, -rho)
    End If
End Function

Public Function ETwoAssetCashOrNothing(OutPutFlag As String, TypeFlag As Integer, S1 As Double, S2 As Double, X1 As Double, X2 As Double, k As Double, T As Double, _
                                    r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        ETwoAssetCashOrNothing = TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "d1" Then 'Delta S1
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "d2" Then 'Delta S2
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e1" Then 'Elasticity S1
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S1 / TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "e2" Then 'Elasticity S2
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S2 / TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "g1" Then 'Gamma S1
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
      ElseIf OutPutFlag = "g2" Then 'Gamma S2
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv1" Then 'DGammaDVol S1
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) _
      - TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho) + 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "gv2" Then 'DGammaDVol S2
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) _
      - TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho) + 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgv1" Then 'Cross GammaDvol S1v2 Cross Zomma
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) _
      - TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho) + 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho) - TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "cgv2" Then 'Cross GammaDvol S2v1 Cross Zomma
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) _
      - TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho) + 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
 
    ElseIf OutPutFlag = "gp1" Then 'GammaP S1
        ETwoAssetCashOrNothing = S1 / 100 * (TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gp2" Then 'GammaP S2
        ETwoAssetCashOrNothing = S2 / 100 * (TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "cg" Then 'Cross gamma
        ETwoAssetCashOrNothing = 1 / (4 * dS * dS) * (TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) _
        - TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "dddv1" Then 'DDeltaDvol S1 v1
        ETwoAssetCashOrNothing = 1 / (4 * dS * dv) * (TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho) _
        - TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "dddv2" Then 'DDeltaDvol S2 v2
        ETwoAssetCashOrNothing = 1 / (4 * dS * dv) * (TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho) _
        - TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho)) / 100

   ElseIf OutPutFlag = "cv1" Then 'Cross vanna 1
        ETwoAssetCashOrNothing = 1 / (4 * dS * dv) * (TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho) _
        - TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho)) / 100
     ElseIf OutPutFlag = "cv2" Then 'Cross vanna 2
        ETwoAssetCashOrNothing = 1 / (4 * dS * dv) * (TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho) _
        - TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "v1" Then 'Vega v1
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho)) / 2
      ElseIf OutPutFlag = "v2" Then 'Vega v2
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "vv1" Then 'DvegaDvol/vomma v1
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho)) / dv ^ 2 / 10000
      ElseIf OutPutFlag = "vv2" Then 'DvegaDvol/vomma v2
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho)) / dv ^ 2 / 10000
   ElseIf OutPutFlag = "cv" Then 'Cross vomma
        ETwoAssetCashOrNothing = 1 / (4 * dv * dv) * (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2 + dv, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2 + dv, rho) _
        - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2 - dv, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2 - dv, rho)) / 10000
  
    ElseIf OutPutFlag = "vp1" Then 'VegaP v1
         ETwoAssetCashOrNothing = v1 / 0.1 * (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho)) / 2
        ElseIf OutPutFlag = "vp2" Then 'VegaP v2
         ETwoAssetCashOrNothing = v2 / 0.1 * (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "dvdv1" Then 'DvegaDvol v1
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1 - dv, v2, rho))
      ElseIf OutPutFlag = "dvdv2" Then 'DvegaDvol v2
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2 - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho + 0.01) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho - 0.01)) / 2
   ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho + 0.01) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ETwoAssetCashOrNothing = TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, 0.00001, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho)
        Else
                ETwoAssetCashOrNothing = TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T - 1 / 365, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r + 0.01, b1 + 0.01, b2 + 0.01, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r - 0.01, b1 - 0.01, b2 - 0.01, v1, v2, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r + 0.01, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r - 0.01, b1, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1 - 0.01, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1 - 0.01, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
  ElseIf OutPutFlag = "f2" Then 'Rho2
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2 - 0.01, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2 + 0.01, v1, v2, rho)) / (2)
  
    ElseIf OutPutFlag = "b1" Then 'Carry S1
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1 + 0.01, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1 - 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "b2" Then 'Carry S2
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2 + 0.01, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2 - 0.01, v1, v2, rho)) / (2)
   
    ElseIf OutPutFlag = "s1" Then 'Speed S1
        ETwoAssetCashOrNothing = 1 / dS ^ 3 * (TwoAssetCashOrNothing(TypeFlag, S1 + 2 * dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) - 3 * TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) _
                                + 3 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "s2" Then 'Speed S2
        ETwoAssetCashOrNothing = 1 / dS ^ 3 * (TwoAssetCashOrNothing(TypeFlag, S1, S2 + 2 * dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) - 3 * TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) _
                                + 3 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "cs1" Then 'Cross speed S1^2S2
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) _
        - TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) + 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "cs2" Then 'Cross speed S2^2S1
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1 + dS, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) _
        - TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2 + dS, X1, X2, k, T, r, b1, b2, v1, v2, rho) + 2 * TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1 - dS, S2 - dS, X1, X2, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx1" Then 'Strike Delta X1
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1 + dS, X2, k, T, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1 - dS, X2, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
        ElseIf OutPutFlag = "dx2" Then 'Strike Delta X2
         ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2 + dS, k, T, r, b1, b2, v1, v2, rho) - TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2 - dS, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
 
     ElseIf OutPutFlag = "dxdx1" Then 'GammaX1
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1 + dS, X2, k, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2, X1 - dS, X2, k, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
       ElseIf OutPutFlag = "dxdx2" Then 'GammaX2
        ETwoAssetCashOrNothing = (TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2 + dS, k, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2, k, T, r, b1, b2, v1, v2, rho) + TwoAssetCashOrNothing(TypeFlag, S1, S2, X1, X2 - dS, k, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     
    End If
End Function


'From Brockhaus et all  modelling and hedging equity derivatives
Public Function BestOfTwoAssetsDigital(OptionType As String, S1 As Double, S2 As Double, X As Double, k As Double, T As Double, r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double) As Double
   ' Best of call pay K at maturity if S1 or S2 is higher than X at maturity
   ' Best of put pay K at maturity if S1 or S2 is lower than X at maturit
    ' Worst of call pay K at maturity if S1 and S2 are higher than X at maturity
   ' Worst of put pay K at maturity if S1 and S2 are lower than X at maturit
   
    Dim y As Double, z1 As Double, z2 As Double, rho1 As Double, rho2 As Double, v As Double
 
    v = Sqr(v1 ^ 2 + v2 ^ 2 - 2 * rho * v1 * v2)
    y = (Log(S1 / S2) + (b1 - b2 + v ^ 2 / 2) * T) / (v * Sqr(T))
    z1 = (Log(S1 / X) + (b1 + v1 ^ 2 / 2) * T) / (v1 * Sqr(T))
    z2 = (Log(S2 / X) + (b2 + v2 ^ 2 / 2) * T) / (v2 * Sqr(T))
    
    rho1 = (v1 - rho * v2) / v
    rho2 = (v2 - rho * v1) / v
    
    If OptionType = "bc" Then 'Best of two assets digital call
        BestOfTwoAssetsDigital = k * Exp(-r * T) * (CBND(y, z1, -rho1) + CBND(-y, z2, -rho2))
    ElseIf OptionType = "bp" Then 'Best of two assets digital put
        BestOfTwoAssetsDigital = k * Exp(-r * T) * (1 - CBND(y, z1, -rho1) - CBND(-y, z2, -rho2))
    ElseIf OptionType = "wc" Then 'Worst of two assets digital call
        BestOfTwoAssetsDigital = k * Exp(-r * T) * (CBND(-y, z1, rho1) + CBND(y, z2, rho2))
    ElseIf OptionType = "wp" Then 'Worst of two assets digital put
        BestOfTwoAssetsDigital = k * Exp(-r * T) * (1 - CBND(-y, z1, rho1) - CBND(y, z2, rho2))
    End If
 End Function
 
 
 
Public Function EBestOfTwoAssetsDigital(OutPutFlag As String, TypeFlag As String, S1 As Double, S2 As Double, X As Double, k As Double, T As Double, _
                                    r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        EBestOfTwoAssetsDigital = BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "d1" Then 'Delta S1
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "d2" Then 'Delta S2
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e1" Then 'Elasticity S1
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S1 / BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "e2" Then 'Elasticity S2
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S2 / BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "g1" Then 'Gamma S1
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1, v2, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
      ElseIf OutPutFlag = "g2" Then 'Gamma S2
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv1" Then 'DGammaDVol S1
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1 + dv, v2, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 + dv, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1 + dv, v2, rho) _
      - BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1 - dv, v2, rho) + 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 - dv, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "gv2" Then 'DGammaDVol S2
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1, v2 + dv, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 + dv, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1, v2 + dv, rho) _
      - BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1, v2 - dv, rho) + 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 - dv, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgv1" Then 'Cross GammaDvol S1v2 Cross Zomma
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1, v2 + dv, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 + dv, rho) + BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1, v2 + dv, rho) _
      - BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1, v2 - dv, rho) + 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 - dv, rho) - BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "cgv2" Then 'Cross GammaDvol S2v1 Cross Zomma
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1 + dv, v2, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 + dv, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1 + dv, v2, rho) _
      - BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1 - dv, v2, rho) + 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 - dv, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
 
    ElseIf OutPutFlag = "gp1" Then 'GammaP S1
        EBestOfTwoAssetsDigital = S1 / 100 * (BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1, v2, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gp2" Then 'GammaP S2
        EBestOfTwoAssetsDigital = S2 / 100 * (BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "cg" Then 'Cross gamma
        EBestOfTwoAssetsDigital = 1 / (4 * dS * dS) * (BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho) _
        - BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "dddv1" Then 'DDeltaDvol S1 v1
        EBestOfTwoAssetsDigital = 1 / (4 * dS * dv) * (BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1 + dv, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1 - dv, v2, rho) _
        - BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1 + dv, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "dddv2" Then 'DDeltaDvol S2 v2
        EBestOfTwoAssetsDigital = 1 / (4 * dS * dv) * (BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1, v2 + dv, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1, v2 - dv, rho) _
        - BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1, v2 + dv, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1, v2 - dv, rho)) / 100

   ElseIf OutPutFlag = "cv1" Then 'Cross vanna 1
        EBestOfTwoAssetsDigital = 1 / (4 * dS * dv) * (BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1, v2 + dv, rho) - BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1, v2 - dv, rho) _
        - BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1, v2 + dv, rho) + BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1, v2 - dv, rho)) / 100
     ElseIf OutPutFlag = "cv2" Then 'Cross vanna 2
        EBestOfTwoAssetsDigital = 1 / (4 * dS * dv) * (BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1 + dv, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1 - dv, v2, rho) _
        - BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1 + dv, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "v1" Then 'Vega v1
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 + dv, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 - dv, v2, rho)) / 2
      ElseIf OutPutFlag = "v2" Then 'Vega v2
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 + dv, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "vv1" Then 'DvegaDvol/vomma v1
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 + dv, v2, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 - dv, v2, rho)) / dv ^ 2 / 10000
      ElseIf OutPutFlag = "vv2" Then 'DvegaDvol/vomma v2
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 + dv, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 - dv, rho)) / dv ^ 2 / 10000
   ElseIf OutPutFlag = "cv" Then 'Cross vomma
        EBestOfTwoAssetsDigital = 1 / (4 * dv * dv) * (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 + dv, v2 + dv, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 - dv, v2 + dv, rho) _
        - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 + dv, v2 - dv, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 - dv, v2 - dv, rho)) / 10000
  
    ElseIf OutPutFlag = "vp1" Then 'VegaP v1
         EBestOfTwoAssetsDigital = v1 / 0.1 * (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 + dv, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 - dv, v2, rho)) / 2
        ElseIf OutPutFlag = "vp2" Then 'VegaP v2
         EBestOfTwoAssetsDigital = v2 / 0.1 * (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 + dv, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "dvdv1" Then 'DvegaDvol v1
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 + dv, v2, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1 - dv, v2, rho))
      ElseIf OutPutFlag = "dvdv2" Then 'DvegaDvol v2
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 + dv, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2 - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho + 0.01) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho - 0.01)) / 2
   ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho + 0.01) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EBestOfTwoAssetsDigital = BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, 0.00001, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho)
        Else
                EBestOfTwoAssetsDigital = BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T - 1 / 365, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r + 0.01, b1 + 0.01, b2 + 0.01, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r - 0.01, b1 - 0.01, b2 - 0.01, v1, v2, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r + 0.01, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r - 0.01, b1, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1 - 0.01, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1 - 0.01, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
  ElseIf OutPutFlag = "f2" Then 'Rho2
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2 - 0.01, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2 + 0.01, v1, v2, rho)) / (2)
  
    ElseIf OutPutFlag = "b1" Then 'Carry S1
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1 + 0.01, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1 - 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "b2" Then 'Carry S2
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2 + 0.01, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2 - 0.01, v1, v2, rho)) / (2)
   
    ElseIf OutPutFlag = "s1" Then 'Speed S1
        EBestOfTwoAssetsDigital = 1 / dS ^ 3 * (BestOfTwoAssetsDigital(TypeFlag, S1 + 2 * dS, S2, X, k, T, r, b1, b2, v1, v2, rho) - 3 * BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1, v2, rho) _
                                + 3 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "s2" Then 'Speed S2
        EBestOfTwoAssetsDigital = 1 / dS ^ 3 * (BestOfTwoAssetsDigital(TypeFlag, S1, S2 + 2 * dS, X, k, T, r, b1, b2, v1, v2, rho) - 3 * BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) _
                                + 3 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "cs1" Then 'Cross speed S1^2S2
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) _
        - BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho) + 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "cs2" Then 'Cross speed S2^2S1
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1 + dS, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho) _
        - BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2 + dS, X, k, T, r, b1, b2, v1, v2, rho) + 2 * BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2, X, k, T, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1 - dS, S2 - dS, X, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X + dS, k, T, r, b1, b2, v1, v2, rho) - BestOfTwoAssetsDigital(TypeFlag, S1, S2, X - dS, k, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
   
     ElseIf OutPutFlag = "dxdx" Then 'GammaX
        EBestOfTwoAssetsDigital = (BestOfTwoAssetsDigital(TypeFlag, S1, S2, X + dS, k, T, r, b1, b2, v1, v2, rho) - 2 * BestOfTwoAssetsDigital(TypeFlag, S1, S2, X, k, T, r, b1, b2, v1, v2, rho) + BestOfTwoAssetsDigital(TypeFlag, S1, S2, X - dS, k, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
      
    End If
End Function

