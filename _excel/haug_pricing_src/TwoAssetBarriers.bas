Attribute VB_Name = "TwoAssetBarriers"
Option Explicit


' Programmer Espen Gaarder Haug
' Copyright 2006, Espen Gaarder Haug

Public Function EPartialTimeTwoAssetBarrier(OutPutFlag As String, TypeFlag As String, S1 As Double, S2 As Double, X As Double, H As Double, t1 As Double, T2 As Double, _
                                    r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    Dim CallPutFlag As String
    Dim OutInnFlag As String
    
    OutInnFlag = Right(TypeFlag, 2)
    CallPutFlag = Left(TypeFlag, 1)
    
     
    If (OutInnFlag = "do" And S2 <= H) Or (OutInnFlag = "uo" And S2 >= H) Then
        EPartialTimeTwoAssetBarrier = 0
        Exit Function
    ElseIf (OutInnFlag = "di" And S2 <= H) Or (OutInnFlag = "ui" And S2 >= H) Then
        If Right(OutPutFlag, 1) = "1" Then OutPutFlag = Left(OutPutFlag, 1)
            EPartialTimeTwoAssetBarrier = EGBlackScholes(OutPutFlag, CallPutFlag, S1, X, T2, r, b1, v1)
        Exit Function
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        EPartialTimeTwoAssetBarrier = PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "d1" Then 'Delta S1
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "d2" Then 'Delta S2
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e1" Then 'Elasticity S1
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS) * S1 / PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "e2" Then 'Elasticity S2
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS) * S2 / PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "g1" Then 'Gamma S1
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho)) / dS ^ 2
      ElseIf OutPutFlag = "g2" Then 'Gamma S2
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv1" Then 'DGammaDVol S1
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) _
      - PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho) + 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "gv2" Then 'DGammaDVol S2
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) _
      - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho) + 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgv1" Then 'Cross GammaDvol S1v2 Cross Zomma
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) _
      - PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho) + 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "cgv2" Then 'Cross GammaDvol S2v1 Cross Zomma
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) _
      - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho) + 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp1" Then 'GammaP S1
        EPartialTimeTwoAssetBarrier = S1 / 100 * (PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gp2" Then 'GammaP S2
        EPartialTimeTwoAssetBarrier = S2 / 100 * (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "cg" Then 'Cross gamma
        EPartialTimeTwoAssetBarrier = 1 / (4 * dS * dS) * (PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) _
        - PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "dddv1" Then 'DDeltaDvol S1 v1
        EPartialTimeTwoAssetBarrier = 1 / (4 * dS * dv) * (PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho) _
        - PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "dddv2" Then 'DDeltaDvol S2 v2
        EPartialTimeTwoAssetBarrier = 1 / (4 * dS * dv) * (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho) _
        - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / 100

   ElseIf OutPutFlag = "cv1" Then 'Cross vanna 1
        EPartialTimeTwoAssetBarrier = 1 / (4 * dS * dv) * (PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho) _
        - PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / 100
     ElseIf OutPutFlag = "cv2" Then 'Cross vanna 2
        EPartialTimeTwoAssetBarrier = 1 / (4 * dS * dv) * (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho) _
        - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "v1" Then 'Vega v1
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / 2
      ElseIf OutPutFlag = "v2" Then 'Vega v2
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "vv1" Then 'DvegaDvol/vomma v1
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / dv ^ 2 / 10000
      ElseIf OutPutFlag = "vv2" Then 'DvegaDvol/vomma v2
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / dv ^ 2 / 10000
   ElseIf OutPutFlag = "cv" Then 'Cross vomma
        EPartialTimeTwoAssetBarrier = 1 / (4 * dv * dv) * (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2 + dv, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2 + dv, rho) _
        - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2 - dv, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2 - dv, rho)) / 10000
  
    ElseIf OutPutFlag = "vp1" Then 'VegaP v1
         EPartialTimeTwoAssetBarrier = v1 / 0.1 * (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho)) / 2
        ElseIf OutPutFlag = "vp2" Then 'VegaP v2
         EPartialTimeTwoAssetBarrier = v2 / 0.1 * (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "dvdv1" Then 'DvegaDvol v1
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 + dv, v2, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1 - dv, v2, rho))
      ElseIf OutPutFlag = "dvdv2" Then 'DvegaDvol v2
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 + dv, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2 - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho + 0.01) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho - 0.01)) / 2
   ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho + 0.01) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If t1 <= 1 / 365 Then
                EPartialTimeTwoAssetBarrier = PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, 0.00001, T2 - 1 / 365, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho)
        Else
                EPartialTimeTwoAssetBarrier = PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1 - 1 / 365, T2 - 1 / 365, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r + 0.01, b1 + 0.01, b2 + 0.01, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r - 0.01, b1 - 0.01, b2 - 0.01, v1, v2, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r + 0.01, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r - 0.01, b1, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1 - 0.01, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1 - 0.01, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
  ElseIf OutPutFlag = "f2" Then 'Rho2
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2 - 0.01, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2 + 0.01, v1, v2, rho)) / (2)
  
    ElseIf OutPutFlag = "b1" Then 'Carry S1
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1 + 0.01, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1 - 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "b2" Then 'Carry S2
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2 + 0.01, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2 - 0.01, v1, v2, rho)) / (2)
   
    ElseIf OutPutFlag = "s1" Then 'Speed S1
        EPartialTimeTwoAssetBarrier = 1 / dS ^ 3 * (PartialTimeTwoAssetBarrier(TypeFlag, S1 + 2 * dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) - 3 * PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) _
                                + 3 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "s2" Then 'Speed S2
        EPartialTimeTwoAssetBarrier = 1 / dS ^ 3 * (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + 2 * dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) - 3 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) _
                                + 3 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "cs1" Then 'Cross speed S1^2S2
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) _
        - PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) + 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "cs2" Then 'Cross speed S2^2S1
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1 + dS, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) _
        - PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2 + dS, X, H, t1, T2, r, b1, b2, v1, v2, rho) + 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1 - dS, S2 - dS, X, H, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X + dS, H, t1, T2, r, b1, b2, v1, v2, rho) - PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X - dS, H, t1, T2, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EPartialTimeTwoAssetBarrier = (PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X + dS, H, t1, T2, r, b1, b2, v1, v2, rho) - 2 * PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X, H, t1, T2, r, b1, b2, v1, v2, rho) + PartialTimeTwoAssetBarrier(TypeFlag, S1, S2, X - dS, H, t1, T2, r, b1, b2, v1, v2, rho)) / dS ^ 2
    End If
End Function


'// Partial-time two asset barrier options
Public Function PartialTimeTwoAssetBarrier(TypeFlag As String, S1 As Double, S2 As Double, X As Double, H As Double, t1 As Double, T2 As Double, _
                r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double) As Double

    Dim d1 As Double, d2 As Double
    Dim d3 As Double, d4 As Double
    Dim e1 As Double, e2 As Double
    Dim e3 As Double, e4 As Double
    Dim mu1 As Double, mu2 As Double
    Dim OutBarrierValue As Double

    Dim eta As Integer
    Dim phi As Integer

    If TypeFlag = "cdo" Or TypeFlag = "pdo" Or TypeFlag = "cdi" Or TypeFlag = "pdi" Then
        phi = -1
    Else
        phi = 1
    End If
    
    If TypeFlag = "cdo" Or TypeFlag = "cuo" Or TypeFlag = "cdi" Or TypeFlag = "cui" Then
        eta = 1
    Else
        eta = -1
    End If
    mu1 = b1 - v1 ^ 2 / 2
    mu2 = b2 - v2 ^ 2 / 2
    d1 = (Log(S1 / X) + (mu1 + v1 ^ 2) * T2) / (v1 * Sqr(T2))
    d2 = d1 - v1 * Sqr(T2)
    d3 = d1 + 2 * rho * Log(H / S2) / (v2 * Sqr(T2))
    d4 = d2 + 2 * rho * Log(H / S2) / (v2 * Sqr(T2))
    e1 = (Log(H / S2) - (mu2 + rho * v1 * v2) * t1) / (v2 * Sqr(t1))
    e2 = e1 + rho * v1 * Sqr(t1)
    e3 = e1 - 2 * Log(H / S2) / (v2 * Sqr(t1))
    e4 = e2 - 2 * Log(H / S2) / (v2 * Sqr(t1))

    OutBarrierValue = eta * S1 * Exp((b1 - r) * T2) * (CBND(eta * d1, phi * e1, -eta * phi * rho * Sqr(t1 / T2)) - Exp(2 * Log(H / S2) * (mu2 + rho * v1 * v2) / (v2 ^ 2)) _
    * CBND(eta * d3, phi * e3, -eta * phi * rho * Sqr(t1 / T2))) _
    - eta * Exp(-r * T2) * X * (CBND(eta * d2, phi * e2, -eta * phi * rho * Sqr(t1 / T2)) - Exp(2 * Log(H / S2) * mu2 / (v2 ^ 2)) _
    * CBND(eta * d4, phi * e4, -eta * phi * rho * Sqr(t1 / T2)))

    If TypeFlag = "cdo" Or TypeFlag = "cuo" Or TypeFlag = "pdo" Or TypeFlag = "puo" Then
        PartialTimeTwoAssetBarrier = OutBarrierValue
     ElseIf TypeFlag = "cui" Or TypeFlag = "cdi" Then
        PartialTimeTwoAssetBarrier = GBlackScholes("c", S1, X, T2, r, b1, v1) - OutBarrierValue
    ElseIf TypeFlag = "pui" Or TypeFlag = "pdi" Then
        PartialTimeTwoAssetBarrier = GBlackScholes("p", S1, X, T2, r, b1, v1) - OutBarrierValue
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


'// Two asset barrier options
Public Function TwoAssetBarrier(TypeFlag As String, S1 As Double, S2 As Double, X As Double, H As Double, _
                T As Double, r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double) As Double
    
    Dim d1 As Double, d2 As Double, d3 As Double, d4 As Double
    Dim e1 As Double, e2 As Double, e3 As Double, e4 As Double
    Dim mu1 As Double, mu2 As Double
    Dim eta As Integer    'Binary variable: 1 for call options and -1 for put options
    Dim phi As Integer    'Binary variable: 1 for up options and -1 for down options
    Dim KnockOutValue As Double
    
    mu1 = b1 - v1 ^ 2 / 2
    mu2 = b2 - v2 ^ 2 / 2
    
    d1 = (Log(S1 / X) + (mu1 + v1 ^ 2) * T) / (v1 * Sqr(T))
    d2 = d1 - v1 * Sqr(T)
    d3 = d1 + 2 * rho * Log(H / S2) / (v2 * Sqr(T))
    d4 = d2 + 2 * rho * Log(H / S2) / (v2 * Sqr(T))
    e1 = (Log(H / S2) - (mu2 + rho * v1 * v2) * T) / (v2 * Sqr(T))
    e2 = e1 + rho * v1 * Sqr(T)
    e3 = e1 - 2 * Log(H / S2) / (v2 * Sqr(T))
    e4 = e2 - 2 * Log(H / S2) / (v2 * Sqr(T))
   
    If TypeFlag = "cuo" Or TypeFlag = "cui" Then
        eta = 1: phi = 1
    ElseIf TypeFlag = "cdo" Or TypeFlag = "cdi" Then
        eta = 1: phi = -1
    ElseIf TypeFlag = "puo" Or TypeFlag = "pui" Then
        eta = -1: phi = 1
    ElseIf TypeFlag = "pdo" Or TypeFlag = "pdi" Then
        eta = -1: phi = -1
    End If
    KnockOutValue = eta * S1 * Exp((b1 - r) * T) * (CBND(eta * d1, phi * e1, -eta * phi * rho) _
    - Exp(2 * (mu2 + rho * v1 * v2) * Log(H / S2) / v2 ^ 2) * CBND(eta * d3, phi * e3, -eta * phi * rho)) - eta * Exp(-r * T) * X * (CBND(eta * d2, phi * e2, -eta * phi * rho) _
    - Exp(2 * mu2 * Log(H / S2) / v2 ^ 2) * CBND(eta * d4, phi * e4, -eta * phi * rho))
    If TypeFlag = "cuo" Or TypeFlag = "cdo" Or TypeFlag = "puo" Or TypeFlag = "pdo" Then
        TwoAssetBarrier = KnockOutValue
    ElseIf TypeFlag = "cui" Or TypeFlag = "cdi" Then
        TwoAssetBarrier = GBlackScholes("c", S1, X, T, r, b1, v1) - KnockOutValue
    ElseIf TypeFlag = "pui" Or TypeFlag = "pdi" Then
        TwoAssetBarrier = GBlackScholes("p", S1, X, T, r, b1, v1) - KnockOutValue
    End If
    
End Function

Public Function ETwoAssetBarrier(OutPutFlag As String, TypeFlag As String, S1 As Double, S2 As Double, X As Double, H As Double, T As Double, _
                                    r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    Dim CallPutFlag As String
    Dim OutInnFlag As String
    
    OutInnFlag = Right(TypeFlag, 2)
    CallPutFlag = Left(TypeFlag, 1)
    
     
    If (OutInnFlag = "do" And S2 <= H) Or (OutInnFlag = "uo" And S2 >= H) Then
        ETwoAssetBarrier = 0
        Exit Function
    ElseIf (OutInnFlag = "di" And S2 <= H) Or (OutInnFlag = "ui" And S2 >= H) Then
        If Right(OutPutFlag, 1) = "1" Then OutPutFlag = Left(OutPutFlag, 1)
            ETwoAssetBarrier = EGBlackScholes(OutPutFlag, CallPutFlag, S1, X, T, r, b1, v1)
        Exit Function
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        ETwoAssetBarrier = TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "d1" Then 'Delta S1
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "d2" Then 'Delta S2
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
    ElseIf OutPutFlag = "e1" Then 'Elasticity S1
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S1 / TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "e2" Then 'Elasticity S2
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho)) / (2 * dS) * S2 / TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho)
    ElseIf OutPutFlag = "g1" Then 'Gamma S1
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
      ElseIf OutPutFlag = "g2" Then 'Gamma S2
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gv1" Then 'DGammaDVol S1
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1 + dv, v2, rho) _
      - TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1 - dv, v2, rho) + 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 - dv, v2, rho) - TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "gv2" Then 'DGammaDVol S2
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1, v2 + dv, rho) _
      - TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1, v2 - dv, rho) + 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 - dv, rho) - TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgv1" Then 'Cross GammaDvol S1v2 Cross Zomma
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1, v2 + dv, rho) _
      - TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1, v2 - dv, rho) + 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 - dv, rho) - TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1, v2 - dv, rho)) / (2 * dv * dS ^ 2) / 100
     ElseIf OutPutFlag = "cgv2" Then 'Cross GammaDvol S2v1 Cross Zomma
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1 + dv, v2, rho) _
      - TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1 - dv, v2, rho) + 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 - dv, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1 - dv, v2, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp1" Then 'GammaP S1
        ETwoAssetBarrier = S1 / 100 * (TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gp2" Then 'GammaP S2
        ETwoAssetBarrier = S2 / 100 * (TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
     ElseIf OutPutFlag = "cg" Then 'Cross gamma
        ETwoAssetBarrier = 1 / (4 * dS * dS) * (TwoAssetBarrier(TypeFlag, S1 + dS, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1 + dS, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho) _
        - TwoAssetBarrier(TypeFlag, S1 - dS, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1 - dS, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "dddv1" Then 'DDeltaDvol S1 v1
        ETwoAssetBarrier = 1 / (4 * dS * dv) * (TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1 - dv, v2, rho) _
        - TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "dddv2" Then 'DDeltaDvol S2 v2
        ETwoAssetBarrier = 1 / (4 * dS * dv) * (TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1, v2 - dv, rho) _
        - TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1, v2 - dv, rho)) / 100

   ElseIf OutPutFlag = "cv1" Then 'Cross vanna 1
        ETwoAssetBarrier = 1 / (4 * dS * dv) * (TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1, v2 - dv, rho) _
        - TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1, v2 + dv, rho) + TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1, v2 - dv, rho)) / 100
     ElseIf OutPutFlag = "cv2" Then 'Cross vanna 2
        ETwoAssetBarrier = 1 / (4 * dS * dv) * (TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1 - dv, v2, rho) _
        - TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1 + dv, v2, rho) + TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1 - dv, v2, rho)) / 100
    ElseIf OutPutFlag = "v1" Then 'Vega v1
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 - dv, v2, rho)) / 2
      ElseIf OutPutFlag = "v2" Then 'Vega v2
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "vv1" Then 'DvegaDvol/vomma v1
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 - dv, v2, rho)) / dv ^ 2 / 10000
      ElseIf OutPutFlag = "vv2" Then 'DvegaDvol/vomma v2
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 - dv, rho)) / dv ^ 2 / 10000
   ElseIf OutPutFlag = "cv" Then 'Cross vomma
        ETwoAssetBarrier = 1 / (4 * dv * dv) * (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 + dv, v2 + dv, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 - dv, v2 + dv, rho) _
        - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 + dv, v2 - dv, rho) + TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 - dv, v2 - dv, rho)) / 10000
  
    ElseIf OutPutFlag = "vp1" Then 'VegaP v1
         ETwoAssetBarrier = v1 / 0.1 * (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 + dv, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 - dv, v2, rho)) / 2
        ElseIf OutPutFlag = "vp2" Then 'VegaP v2
         ETwoAssetBarrier = v2 / 0.1 * (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 + dv, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 - dv, rho)) / 2
 
     ElseIf OutPutFlag = "dvdv1" Then 'DvegaDvol v1
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 + dv, v2, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1 - dv, v2, rho))
      ElseIf OutPutFlag = "dvdv2" Then 'DvegaDvol v2
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 + dv, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2 - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho + 0.01) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho - 0.01)) / 2
   ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho + 0.01) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ETwoAssetBarrier = TwoAssetBarrier(TypeFlag, S1, S2, X, H, 0.00001, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho)
        Else
                ETwoAssetBarrier = TwoAssetBarrier(TypeFlag, S1, S2, X, H, T - 1 / 365, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r + 0.01, b1 + 0.01, b2 + 0.01, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r - 0.01, b1 - 0.01, b2 - 0.01, v1, v2, rho)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r + 0.01, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r - 0.01, b1, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1 - 0.01, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "f1" Then 'Rho2
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1 - 0.01, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1 + 0.01, b2, v1, v2, rho)) / (2)
  ElseIf OutPutFlag = "f2" Then 'Rho2
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2 - 0.01, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2 + 0.01, v1, v2, rho)) / (2)
  
    ElseIf OutPutFlag = "b1" Then 'Carry S1
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1 + 0.01, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1 - 0.01, b2, v1, v2, rho)) / (2)
     ElseIf OutPutFlag = "b2" Then 'Carry S2
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2 + 0.01, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2 - 0.01, v1, v2, rho)) / (2)
   
    ElseIf OutPutFlag = "s1" Then 'Speed S1
        ETwoAssetBarrier = 1 / dS ^ 3 * (TwoAssetBarrier(TypeFlag, S1 + 2 * dS, S2, X, H, T, r, b1, b2, v1, v2, rho) - 3 * TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1, v2, rho) _
                                + 3 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "s2" Then 'Speed S2
        ETwoAssetBarrier = 1 / dS ^ 3 * (TwoAssetBarrier(TypeFlag, S1, S2 + 2 * dS, X, H, T, r, b1, b2, v1, v2, rho) - 3 * TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) _
                                + 3 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho))
    ElseIf OutPutFlag = "cs1" Then 'Cross speed S1^2S2
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1 + dS, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1 - dS, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) _
        - TwoAssetBarrier(TypeFlag, S1 + dS, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho) + 2 * TwoAssetBarrier(TypeFlag, S1, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1 - dS, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "cs2" Then 'Cross speed S2^2S1
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1 + dS, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetBarrier(TypeFlag, S1 + dS, S2, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1 + dS, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho) _
        - TwoAssetBarrier(TypeFlag, S1 - dS, S2 + dS, X, H, T, r, b1, b2, v1, v2, rho) + 2 * TwoAssetBarrier(TypeFlag, S1 - dS, S2, X, H, T, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1 - dS, S2 - dS, X, H, T, r, b1, b2, v1, v2, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X + dS, H, T, r, b1, b2, v1, v2, rho) - TwoAssetBarrier(TypeFlag, S1, S2, X - dS, H, T, r, b1, b2, v1, v2, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        ETwoAssetBarrier = (TwoAssetBarrier(TypeFlag, S1, S2, X + dS, H, T, r, b1, b2, v1, v2, rho) - 2 * TwoAssetBarrier(TypeFlag, S1, S2, X, H, T, r, b1, b2, v1, v2, rho) + TwoAssetBarrier(TypeFlag, S1, S2, X - dS, H, T, r, b1, b2, v1, v2, rho)) / dS ^ 2
    End If
End Function
