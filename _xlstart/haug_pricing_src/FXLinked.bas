Attribute VB_Name = "FXLinked"
Option Explicit

' Implementation By Espen Gaarder Haug
' Copyright Espen Gaarder Haug 2006


Public Function ETakeoverFXoption(OutPutFlag As String, VFirm As Double, N As Double, E As Double, X As Double, T As Double, r As Double, rf As Double, vV As Double, vE As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        ETakeoverFXoption = TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho)
    ElseIf OutPutFlag = "dE" Then 'Delta E
         ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE, rho)) / (2 * dS)
    ElseIf OutPutFlag = "dS" Then 'Delta S
         ETakeoverFXoption = (TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV, vE, rho)) / (2 * dS)
    ElseIf OutPutFlag = "eE" Then 'Elasticity E
         ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE, rho)) / (2 * dS) * E / TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho)
   ElseIf OutPutFlag = "eS" Then 'Elasticity S
         ETakeoverFXoption = (TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV, vE, rho)) / (2 * dS) * VFirm / TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho)
    ElseIf OutPutFlag = "gE" Then 'Gamma E
        ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV, vE, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gS" Then 'Gamma S
        ETakeoverFXoption = (TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV, vE, rho) - 2 * TakeoverFXoption(VFirm, N, X, E, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV, vE, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gvE" Then 'DGammaDVol E
        ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV, vE + dv, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE + dv, rho) + TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE + dv, rho) _
      - TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV, vE - dv, rho) + 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE - dv, rho) - TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "gvS" Then 'DGammaDVol S
        ETakeoverFXoption = (TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV + dv, vE, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV + dv, vE, rho) + TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV + dv, vE, rho) _
      - TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV - dv, vE, rho) + 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV - dv, vE, rho) - TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV - dv, vE, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgvE" Then 'Cross GammaDvol EvS (Cross Zomma)
        ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV + dv, vE, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV + dv, vE, rho) + TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV + dv, vE, rho) _
      - TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV - dv, vE, rho) + 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV - dv, vE, rho) - TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV - dv, vE, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgvS" Then 'Cross GammaDvol SvE (Cross Zomma)
        ETakeoverFXoption = (TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV, vE + dv, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE + dv, rho) + TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV, vE + dv, rho) _
      - TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV, vE - dv, rho) + 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE - dv, rho) - TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV, vE - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "gpE" Then 'GammaP E
        ETakeoverFXoption = E / 100 * (TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV, vE, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gpS" Then 'GammaP S
        ETakeoverFXoption = VFirm / 100 * (TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV, vE, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV, vE, rho)) / dS ^ 2
    ElseIf OutPutFlag = "cg" Then 'Cross gamma
        ETakeoverFXoption = 1 / (4 * dS * dS) * (TakeoverFXoption(VFirm + dS, N, E + dS, X, T, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm - dS, N, E + dS, X, T, r, rf, vV, vE, rho) _
        - TakeoverFXoption(VFirm + dS, N, E - dS, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm - dS, N, E - dS, X, T, r, rf, vV, vE, rho))
    ElseIf OutPutFlag = "dddvE" Then 'DDeltaDvol E vE
        ETakeoverFXoption = 1 / (4 * dS * dv) * (TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV, vE + dv, rho) - TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV, vE - dv, rho) _
        - TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE + dv, rho) + TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE - dv, rho)) / 100
    ElseIf OutPutFlag = "dddvS" Then 'DDeltaDvol S vS
        ETakeoverFXoption = 1 / (4 * dS * dv) * (TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV + dv, vE, rho) - TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV - dv, vE, rho) _
        - TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV + dv, vE, rho) + TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV - dv, vE, rho)) / 100
    ElseIf OutPutFlag = "cvE" Then 'Cross vanna E vS
        ETakeoverFXoption = 1 / (4 * dS * dv) * (TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV + dv, vE, rho) - TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV - dv, vE, rho) _
        - TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV + dv, vE, rho) + TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV - dv, vE, rho)) / 100
    ElseIf OutPutFlag = "cvS" Then 'Cross vanna S vE
        ETakeoverFXoption = 1 / (4 * dS * dv) * (TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV, vE + dv, rho) - TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV, vE - dv, rho) _
        - TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV, vE + dv, rho) + TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV, vE - dv, rho)) / 100
    ElseIf OutPutFlag = "vE" Then 'Vega vE
         ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE + dv, rho) - TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE - dv, rho)) / 2
    ElseIf OutPutFlag = "vS" Then 'Vega vS
          ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV + dv, vE, rho) - TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV - dv, vE, rho)) / 2
    ElseIf OutPutFlag = "dvdvE" Then 'DvegaDvol/vomma vE
            ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE + dv, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE - dv, rho)) / dv ^ 2 / 10000
    ElseIf OutPutFlag = "dvdvS" Then 'DvegaDvol/vomma vS
            ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV + dv, vE, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV - dv, vE, rho)) / dv ^ 2 / 10000
    ElseIf OutPutFlag = "cv" Then 'Cross vomma
        ETakeoverFXoption = 1 / (4 * dv * dv) * (TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV + dv, vE + dv, rho) - TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV - dv, vE + dv, rho) _
        - TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV + dv, vE - dv, rho) + TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV - dv, vE - dv, rho)) / 10000
    ElseIf OutPutFlag = "vpE" Then 'VegaP vE
         ETakeoverFXoption = vV / 0.1 * (TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE + dv, rho) - TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE - dv, rho)) / 2
    ElseIf OutPutFlag = "vpS" Then 'VegaP vS
         ETakeoverFXoption = vE / 0.1 * (TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV + dv, vE, rho) - TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV - dv, vE, rho)) / 2
    ElseIf OutPutFlag = "dvdvE" Then 'DvegaDvol vE
             ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE + dv, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE - dv, rho))
    ElseIf OutPutFlag = "dvdvS" Then 'DvegaDvol vS
            ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV + dv, vE, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV - dv, vE, rho))
     ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho + 0.01) - TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho - 0.01)) / 2
    ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho + 0.01) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ETakeoverFXoption = TakeoverFXoption(VFirm, N, E, X, 0.00001, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho)
        Else
                ETakeoverFXoption = TakeoverFXoption(VFirm, N, E, X, T - 1 / 365, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X, T, r + 0.01, rf, vV, vE, rho) - TakeoverFXoption(VFirm, N, E, X, T, r - 0.01, rf, vV, vE, rho)) / (2)
  ElseIf OutPutFlag = "f" Then 'Phi
         ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X, T, r, rf - 0.01, vV, vE, rho) - TakeoverFXoption(VFirm, N, E, X, T, r, rf + 0.01, vV, vE, rho)) / (2)
      ElseIf OutPutFlag = "sE" Then 'Speed E
        ETakeoverFXoption = 1 / dS ^ 3 * (TakeoverFXoption(VFirm, N, E + 2 * dS, X, T, r, rf, vV, vE, rho) - 3 * TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV, vE, rho) _
                                + 3 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE, rho))
    ElseIf OutPutFlag = "sS" Then 'Speed S
        ETakeoverFXoption = 1 / dS ^ 3 * (TakeoverFXoption(VFirm + 2 * dS, N, E, X, T, r, rf, vV, vE, rho) - 3 * TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV, vE, rho) _
                                + 3 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV, vE, rho))
    ElseIf OutPutFlag = "csE" Then 'Cross speed E^2S
        ETakeoverFXoption = (TakeoverFXoption(VFirm + dS, N, E + dS, X, T, r, rf, vV, vE, rho) - 2 * TakeoverFXoption(VFirm + dS, N, E, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE, rho) _
        - TakeoverFXoption(VFirm - dS, N, E + dS, X, T, r, rf, vV, vE, rho) + 2 * TakeoverFXoption(VFirm - dS, N, E, X, T, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "csS" Then 'Cross speed S^2E
        ETakeoverFXoption = (TakeoverFXoption(VFirm + dS, N, E + dS, X, T, r, rf, vV, vE, rho) - 2 * TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm, N, E + dS, X, T, r, rf, vV, vE, rho) _
        - TakeoverFXoption(VFirm + dS, N, E - dS, X, T, r, rf, vV, vE, rho) + 2 * TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm, N, E - dS, X, T, r, rf, vV, vE, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X + dS, T, r, rf, vV, vE, rho) - TakeoverFXoption(VFirm, N, E, X - dS, T, r, rf, vV, vE, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        ETakeoverFXoption = (TakeoverFXoption(VFirm, N, E, X + dS, T, r, rf, vV, vE, rho) - 2 * TakeoverFXoption(VFirm, N, E, X, T, r, rf, vV, vE, rho) + TakeoverFXoption(VFirm, N, E, X - dS, T, r, rf, vV, vE, rho)) / dS ^ 2
    End If
    
End Function


'// Takeover foreign exchange options
Public Function TakeoverFXoption(VFirm As Double, N As Double, E As Double, X As Double, T As Double, r As Double, rf As Double, vV As Double, vE As Double, rho As Double) As Double
    
    Dim a1 As Double, a2 As Double
    
    a1 = (Log(VFirm / N) + (rf - rho * vE * vV - vV ^ 2 / 2) * T) / (vV * Sqr(T))
    a2 = (Log(E / X) + (r - rf - vE ^ 2 / 2) * T) / (vE * Sqr(T))
    
    TakeoverFXoption = N * (E * Exp(-rf * T) * CBND(a2 + vE * Sqr(T), -a1 - rho * vE * Sqr(T), -rho) _
    - X * Exp(-r * T) * CBND(-a1, a2, -rho))

End Function


'// Foreign equity option struck in domestic currency
Public Function ForEquOptInDomCur(CallPutFlag As String, E As Double, S As Double, X As Double, T As Double, _
    r As Double, q As Double, vS As Double, vE As Double, rho As Double) As Double

    Dim v As Double, d1 As Double, d2 As Double

    v = Sqr(vE ^ 2 + vS ^ 2 + 2 * rho * vE * vS)
    d1 = (Log(E * S / X) + (r - q + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
   
    If CallPutFlag = "c" Then
        ForEquOptInDomCur = E * S * Exp(-q * T) * CND(d1) - X * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
        ForEquOptInDomCur = X * Exp(-r * T) * CND(-d2) - E * S * Exp(-q * T) * CND(-d1)
    End If
End Function




Public Function EForEquOptInDomCur(OutPutFlag As String, CallPutFlag As String, E As Double, S As Double, X As Double, T As Double, _
                                    r As Double, q As Double, vE As Double, vS As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        EForEquOptInDomCur = ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho)
    ElseIf OutPutFlag = "dE" Then 'Delta E
         EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE, vS, rho)) / (2 * dS)
    ElseIf OutPutFlag = "dS" Then 'Delta S
         EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE, vS, rho)) / (2 * dS)
    ElseIf OutPutFlag = "eE" Then 'Elasticity E
         EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE, vS, rho)) / (2 * dS) * E / ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho)
   ElseIf OutPutFlag = "eS" Then 'Elasticity S
         EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE, vS, rho)) / (2 * dS) * S / ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho)
    ElseIf OutPutFlag = "gE" Then 'Gamma E
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE, vS, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE, vS, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gS" Then 'Gamma S
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE, vS, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE, vS, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gvE" Then 'DGammaDVol E
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE + dv, vS, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE + dv, vS, rho) + ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE + dv, vS, rho) _
      - ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE - dv, vS, rho) + 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE - dv, vS, rho) - ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE - dv, vS, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "gvS" Then 'DGammaDVol S
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE, vS + dv, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS + dv, rho) + ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE, vS + dv, rho) _
      - ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE, vS - dv, rho) + 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS - dv, rho) - ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE, vS - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgvE" Then 'Cross GammaDvol EvS (Cross Zomma)
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE, vS + dv, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS + dv, rho) + ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE, vS + dv, rho) _
      - ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE, vS - dv, rho) + 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS - dv, rho) - ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE, vS - dv, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgvS" Then 'Cross GammaDvol SvE (Cross Zomma)
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE + dv, vS, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE + dv, vS, rho) + ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE + dv, vS, rho) _
      - ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE - dv, vS, rho) + 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE - dv, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE - dv, vS, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "gpE" Then 'GammaP E
        EForEquOptInDomCur = E / 100 * (ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE, vS, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE, vS, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gpS" Then 'GammaP S
        EForEquOptInDomCur = S / 100 * (ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE, vS, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE, vS, rho)) / dS ^ 2
    ElseIf OutPutFlag = "cg" Then 'Cross gamma
        EForEquOptInDomCur = 1 / (4 * dS * dS) * (ForEquOptInDomCur(CallPutFlag, E + dS, S + dS, X, T, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E + dS, S - dS, X, T, r, q, vE, vS, rho) _
        - ForEquOptInDomCur(CallPutFlag, E - dS, S + dS, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E - dS, S - dS, X, T, r, q, vE, vS, rho))
    ElseIf OutPutFlag = "dddvE" Then 'DDeltaDvol E vE
        EForEquOptInDomCur = 1 / (4 * dS * dv) * (ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE + dv, vS, rho) - ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE - dv, vS, rho) _
        - ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE + dv, vS, rho) + ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE - dv, vS, rho)) / 100
    ElseIf OutPutFlag = "dddvS" Then 'DDeltaDvol S vS
        EForEquOptInDomCur = 1 / (4 * dS * dv) * (ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE, vS + dv, rho) - ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE, vS - dv, rho) _
        - ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE, vS + dv, rho) + ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE, vS - dv, rho)) / 100
    ElseIf OutPutFlag = "cvE" Then 'Cross vanna E vS
        EForEquOptInDomCur = 1 / (4 * dS * dv) * (ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE, vS + dv, rho) - ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE, vS - dv, rho) _
        - ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE, vS + dv, rho) + ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE, vS - dv, rho)) / 100
    ElseIf OutPutFlag = "cvS" Then 'Cross vanna S vE
        EForEquOptInDomCur = 1 / (4 * dS * dv) * (ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE + dv, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE - dv, vS, rho) _
        - ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE + dv, vS, rho) + ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE - dv, vS, rho)) / 100
    ElseIf OutPutFlag = "vE" Then 'Vega vE
         EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE + dv, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE - dv, vS, rho)) / 2
    ElseIf OutPutFlag = "vS" Then 'Vega vS
         EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS + dv, rho) - ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS - dv, rho)) / 2
    ElseIf OutPutFlag = "dvdvE" Then 'DvegaDvol/vomma vE
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE + dv, vS, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE - dv, vS, rho)) / dv ^ 2 / 10000
    ElseIf OutPutFlag = "dvdvS" Then 'DvegaDvol/vomma vS
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS + dv, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS - dv, rho)) / dv ^ 2 / 10000
    ElseIf OutPutFlag = "cv" Then 'Cross vomma
        EForEquOptInDomCur = 1 / (4 * dv * dv) * (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE + dv, vS + dv, rho) - ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE - dv, vS + dv, rho) _
        - ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE + dv, vS - dv, rho) + ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE - dv, vS - dv, rho)) / 10000
    ElseIf OutPutFlag = "vpE" Then 'VegaP vE
         EForEquOptInDomCur = vE / 0.1 * (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE + dv, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE - dv, vS, rho)) / 2
    ElseIf OutPutFlag = "vpS" Then 'VegaP vS
         EForEquOptInDomCur = vS / 0.1 * (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS + dv, rho) - ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS - dv, rho)) / 2
    ElseIf OutPutFlag = "dvdvE" Then 'DvegaDvol vE
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE + dv, vS, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE - dv, vS, rho))
    ElseIf OutPutFlag = "dvdvS" Then 'DvegaDvol vS
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS + dv, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho + 0.01) - ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho - 0.01)) / 2
    ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho + 0.01) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EForEquOptInDomCur = ForEquOptInDomCur(CallPutFlag, E, S, X, 0.00001, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho)
        Else
                EForEquOptInDomCur = ForEquOptInDomCur(CallPutFlag, E, S, X, T - 1 / 365, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r + 0.01 + 0.01, q + 0.01, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S, X, T, r - 0.01 - 0.01, q - 0.01, vE, vS, rho)) / (2)
  ElseIf OutPutFlag = "dr" Then 'Dividend rho
         EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q - 0.01, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q + 0.01, vE, vS, rho)) / (2)
      ElseIf OutPutFlag = "sE" Then 'Speed E
        EForEquOptInDomCur = 1 / dS ^ 3 * (ForEquOptInDomCur(CallPutFlag, E + 2 * dS, S, X, T, r, q, vE, vS, rho) - 3 * ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE, vS, rho) _
                                + 3 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE, vS, rho))
    ElseIf OutPutFlag = "sS" Then 'Speed S
        EForEquOptInDomCur = 1 / dS ^ 3 * (ForEquOptInDomCur(CallPutFlag, E, S + 2 * dS, X, T, r, q, vE, vS, rho) - 3 * ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE, vS, rho) _
                                + 3 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE, vS, rho))
    ElseIf OutPutFlag = "csE" Then 'Cross speed E^2S
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E + dS, S + dS, X, T, r, q, vE, vS, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S + dS, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E - dS, S + dS, X, T, r, q, vE, vS, rho) _
        - ForEquOptInDomCur(CallPutFlag, E + dS, S - dS, X, T, r, q, vE, vS, rho) + 2 * ForEquOptInDomCur(CallPutFlag, E, S - dS, X, T, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E - dS, S - dS, X, T, r, q, vE, vS, rho)) / (2 * dS ^ 3)
   ElseIf OutPutFlag = "csS" Then 'Cross speed S^2E
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E + dS, S + dS, X, T, r, q, vE, vS, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E + dS, S, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E + dS, S - dS, X, T, r, q, vE, vS, rho) _
        - ForEquOptInDomCur(CallPutFlag, E - dS, S + dS, X, T, r, q, vE, vS, rho) + 2 * ForEquOptInDomCur(CallPutFlag, E - dS, S, X, T, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E - dS, S - dS, X, T, r, q, vE, vS, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X + dS, T, r, q, vE, vS, rho) - ForEquOptInDomCur(CallPutFlag, E, S, X - dS, T, r, q, vE, vS, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        EForEquOptInDomCur = (ForEquOptInDomCur(CallPutFlag, E, S, X + dS, T, r, q, vE, vS, rho) - 2 * ForEquOptInDomCur(CallPutFlag, E, S, X, T, r, q, vE, vS, rho) + ForEquOptInDomCur(CallPutFlag, E, S, X - dS, T, r, q, vE, vS, rho)) / dS ^ 2
    End If
    
End Function

'// Fixed exchange rate foreign equity options-- Quantos
Public Function Quanto(CallPutFlag As String, Ep As Double, S As Double, X As Double, T As Double, r As Double, _
                rf As Double, q As Double, vS As Double, vE As Double, rho As Double) As Double
    
    Dim d1 As Double, d2 As Double

    d1 = (Log(S / X) + (rf - q - rho * vS * vE + vS ^ 2 / 2) * T) / (vS * Sqr(T))
    d2 = d1 - vS * Sqr(T)
   
    If CallPutFlag = "c" Then
        Quanto = Ep * (S * Exp((rf - r - q - rho * vS * vE) * T) * CND(d1) - X * Exp(-r * T) * CND(d2))
    ElseIf CallPutFlag = "p" Then
        Quanto = Ep * (X * Exp(-r * T) * CND(-d2) - S * Exp((rf - r - q - rho * vS * vE) * T) * CND(-d1))
    End If
End Function


Public Function EQuanto(OutPutFlag As String, CallPutFlag As String, Ep As Double, S As Double, X As Double, T As Double, _
                                    r As Double, rf As Double, q As Double, vE As Double, vS As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        EQuanto = Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho)
    ElseIf OutPutFlag = "dE" Then 'Delta Ep
         EQuanto = (Quanto(CallPutFlag, Ep + dS, S, X, T, r, rf, q, vE, vS, rho) - Quanto(CallPutFlag, Ep - dS, S, X, T, r, rf, q, vE, vS, rho)) / (2 * dS)
    ElseIf OutPutFlag = "dS" Then 'Delta S
         EQuanto = (Quanto(CallPutFlag, Ep, S + dS, X, T, r, rf, q, vE, vS, rho) - Quanto(CallPutFlag, Ep, S - dS, X, T, r, rf, q, vE, vS, rho)) / (2 * dS)
   ElseIf OutPutFlag = "eE" Then 'Elasticity Ep
         EQuanto = (Quanto(CallPutFlag, Ep + dS, S, X, T, r, rf, q, vE, vS, rho) - Quanto(CallPutFlag, Ep - dS, S, X, T, r, rf, q, vE, vS, rho)) / (2 * dS) * S / Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho)
   ElseIf OutPutFlag = "eS" Then 'Elasticity S
         EQuanto = (Quanto(CallPutFlag, Ep, S + dS, X, T, r, rf, q, vE, vS, rho) - Quanto(CallPutFlag, Ep, S - dS, X, T, r, rf, q, vE, vS, rho)) / (2 * dS) * S / Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho)
    ElseIf OutPutFlag = "gS" Then 'Gamma S
        EQuanto = (Quanto(CallPutFlag, Ep, S + dS, X, T, r, rf, q, vE, vS, rho) - 2 * Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho) + Quanto(CallPutFlag, Ep, S - dS, X, T, r, rf, q, vE, vS, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gpS" Then 'GammaP S
        EQuanto = S / 100 * (Quanto(CallPutFlag, Ep, S + dS, X, T, r, rf, q, vE, vS, rho) - 2 * Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho) + Quanto(CallPutFlag, Ep, S - dS, X, T, r, rf, q, vE, vS, rho)) / dS ^ 2
    ElseIf OutPutFlag = "gvS" Then 'DGammaDVol S
        EQuanto = (Quanto(CallPutFlag, Ep, S + dS, X, T, r, rf, q, vE, vS + dv, rho) - 2 * Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS + dv, rho) + Quanto(CallPutFlag, Ep, S - dS, X, T, r, rf, q, vE, vS + dv, rho) _
      - Quanto(CallPutFlag, Ep, S + dS, X, T, r, rf, q, vE, vS - dv, rho) + 2 * Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS - dv, rho) - Quanto(CallPutFlag, Ep, S - dS, X, T, r, rf, q, vE, vS - dv, rho)) / (2 * dv * dS ^ 2) / 100
      ElseIf OutPutFlag = "dddvE" Then 'DDeltaDvol Ep vE
        EQuanto = 1 / (4 * dS * dv) * (Quanto(CallPutFlag, Ep + dS, S, X, T, r, rf, q, vE + dv, vS, rho) - Quanto(CallPutFlag, Ep + dS, S, X, T, r, rf, q, vE - dv, vS, rho) _
        - Quanto(CallPutFlag, Ep - dS, S, X, T, r, rf, q, vE + dv, vS, rho) + Quanto(CallPutFlag, Ep - dS, S, X, T, r, rf, q, vE - dv, vS, rho)) / 100
    ElseIf OutPutFlag = "dddvS" Then 'DDeltaDvol S vS
        EQuanto = 1 / (4 * dS * dv) * (Quanto(CallPutFlag, Ep, S + dS, X, T, r, rf, q, vE, vS + dv, rho) - Quanto(CallPutFlag, Ep, S + dS, X, T, r, rf, q, vE, vS - dv, rho) _
        - Quanto(CallPutFlag, Ep, S - dS, X, T, r, rf, q, vE, vS + dv, rho) + Quanto(CallPutFlag, Ep, S - dS, X, T, r, rf, q, vE, vS - dv, rho)) / 100
    ElseIf OutPutFlag = "cvE" Then 'Cross vanna E
        EQuanto = 1 / (4 * dS * dv) * (Quanto(CallPutFlag, Ep + dS, S, X, T, r, rf, q, vE, vS + dv, rho) - Quanto(CallPutFlag, Ep + dS, S, X, T, r, rf, q, vE, vS - dv, rho) _
        - Quanto(CallPutFlag, Ep - dS, S, X, T, r, rf, q, vE, vS + dv, rho) + Quanto(CallPutFlag, Ep - dS, S, X, T, r, rf, q, vE, vS - dv, rho)) / 100
    ElseIf OutPutFlag = "cvS" Then 'Cross vanna S
        EQuanto = 1 / (4 * dS * dv) * (Quanto(CallPutFlag, Ep, S + dS, X, T, r, rf, q, vE + dv, vS, rho) - Quanto(CallPutFlag, Ep, S + dS, X, T, r, rf, q, vE - dv, vS, rho) _
        - Quanto(CallPutFlag, Ep, S - dS, X, T, r, rf, q, vE + dv, vS, rho) + Quanto(CallPutFlag, Ep, S - dS, X, T, r, rf, q, vE - dv, vS, rho)) / 100
    ElseIf OutPutFlag = "vE" Then 'Vega vE
         EQuanto = (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE + dv, vS, rho) - Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE - dv, vS, rho)) / 2
    ElseIf OutPutFlag = "vS" Then 'Vega vS
         EQuanto = (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS + dv, rho) - Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS - dv, rho)) / 2
    ElseIf OutPutFlag = "vvE" Then 'DvegaDvol/vomma vE
        EQuanto = (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE + dv, vS, rho) - 2 * Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho) + Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE - dv, vS, rho)) / dv ^ 2 / 10000
    ElseIf OutPutFlag = "vvS" Then 'DvegaDvol/vomma vS
        EQuanto = (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS + dv, rho) - 2 * Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho) + Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS - dv, rho)) / dv ^ 2 / 10000
    ElseIf OutPutFlag = "cv" Then 'Cross vomma
        EQuanto = 1 / (4 * dv * dv) * (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE + dv, vS + dv, rho) - Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE - dv, vS + dv, rho) _
        - Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE + dv, vS - dv, rho) + Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE - dv, vS - dv, rho)) / 10000
    ElseIf OutPutFlag = "vpE" Then 'VegaP vE
         EQuanto = vE / 0.1 * (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE + dv, vS, rho) - Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE - dv, vS, rho)) / 2
    ElseIf OutPutFlag = "vpS" Then 'VegaP vS
         EQuanto = vS / 0.1 * (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS + dv, rho) - Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS - dv, rho)) / 2
    ElseIf OutPutFlag = "dvdvE" Then 'DvegaDvol vE
        EQuanto = (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE + dv, vS, rho) - 2 * Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho) + Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE - dv, vS, rho))
    ElseIf OutPutFlag = "dvdvS" Then 'DvegaDvol vS
        EQuanto = (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS + dv, rho) - 2 * Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho) + Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         EQuanto = (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho + 0.01) - Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho - 0.01)) / 2
    ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        EQuanto = (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho + 0.01) - 2 * Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho) + Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EQuanto = Quanto(CallPutFlag, Ep, S, X, 0.00001, r, rf, q, vE, vS, rho) - Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho)
        Else
                EQuanto = Quanto(CallPutFlag, Ep, S, X, T - 1 / 365, r, rf, q, vE, vS, rho) - Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         EQuanto = (Quanto(CallPutFlag, Ep, S, X, T, r + 0.01, rf + 0.01, q + 0.01, vE, vS, rho) - Quanto(CallPutFlag, Ep, S, X, T, r - 0.01, rf - 0.01, q - 0.01, vE, vS, rho)) / (2)
    ElseIf OutPutFlag = "f" Then 'Phi
         EQuanto = (Quanto(CallPutFlag, Ep, S, X, T, r, rf - 0.01, q, vE, vS, rho) - Quanto(CallPutFlag, Ep, S, X, T, r, rf + 0.01, q, vE, vS, rho)) / (2)
  ElseIf OutPutFlag = "dr" Then 'Rho2
         EQuanto = (Quanto(CallPutFlag, Ep, S, X, T, r, rf, q - 0.01, vE, vS, rho) - Quanto(CallPutFlag, Ep, S, X, T, r, rf, q + 0.01, vE, vS, rho)) / (2)
     ElseIf OutPutFlag = "sS" Then 'Speed S
        EQuanto = 1 / dS ^ 3 * (Quanto(CallPutFlag, Ep, S + 2 * dS, X, T, r, rf, q, vE, vS, rho) - 3 * Quanto(CallPutFlag, Ep, S + dS, X, T, r, rf, q, vE, vS, rho) _
                                + 3 * Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho) - Quanto(CallPutFlag, Ep, S - dS, X, T, r, rf, q, vE, vS, rho))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EQuanto = (Quanto(CallPutFlag, Ep, S, X + dS, T, r, rf, q, vE, vS, rho) - Quanto(CallPutFlag, Ep, S, X - dS, T, r, rf, q, vE, vS, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EQuanto = (Quanto(CallPutFlag, Ep, S, X + dS, T, r, rf, q, vE, vS, rho) - 2 * Quanto(CallPutFlag, Ep, S, X, T, r, rf, q, vE, vS, rho) + Quanto(CallPutFlag, Ep, S, X - dS, T, r, rf, q, vE, vS, rho)) / dS ^ 2
    End If
End Function




Public Function EEquityLinkedFXO(OutPutFlag As String, CallPutFlag As String, E As Double, S As Double, X As Double, T As Double, _
                                    r As Double, rf As Double, q As Double, vE As Double, vS As Double, rho As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    Dim dv As Double
    
    dv = 0.01
    
    If OutPutFlag = "p" Then ' Value
        EEquityLinkedFXO = EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho)
    ElseIf OutPutFlag = "dE" Then 'Delta E
         EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE, vS, rho)) / (2 * dS)
    ElseIf OutPutFlag = "dS" Then 'Delta S
         EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S + dS, X, T, r, rf, q, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E, S - dS, X, T, r, rf, q, vE, vS, rho)) / (2 * dS)
    ElseIf OutPutFlag = "eE" Then 'Elasticity E
         EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE, vS, rho)) / (2 * dS) * E / EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho)
   ElseIf OutPutFlag = "eS" Then 'Elasticity S
         EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S + dS, X, T, r, rf, q, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E, S - dS, X, T, r, rf, q, vE, vS, rho)) / (2 * dS) * S / EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho)
    ElseIf OutPutFlag = "gE" Then 'Gamma E
        EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE, vS, rho) - 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho) + EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE, vS, rho)) / dS ^ 2
     ElseIf OutPutFlag = "gvE" Then 'DGammaDVol E
        EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE + dv, vS, rho) - 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE + dv, vS, rho) + EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE + dv, vS, rho) _
      - EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE - dv, vS, rho) + 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE - dv, vS, rho) - EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE - dv, vS, rho)) / (2 * dv * dS ^ 2) / 100
    ElseIf OutPutFlag = "cgvE" Then 'Cross GammaDvol EvS (Cross Zomma)
        EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE, vS + dv, rho) - 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS + dv, rho) + EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE, vS + dv, rho) _
      - EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE, vS - dv, rho) + 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS - dv, rho) - EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE, vS - dv, rho)) / (2 * dv * dS ^ 2) / 100
   ElseIf OutPutFlag = "gpE" Then 'GammaP E
        EEquityLinkedFXO = E / 100 * (EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE, vS, rho) - 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho) + EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE, vS, rho)) / dS ^ 2
     ElseIf OutPutFlag = "dddvE" Then 'DDeltaDvol E vE
        EEquityLinkedFXO = 1 / (4 * dS * dv) * (EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE + dv, vS, rho) - EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE - dv, vS, rho) _
        - EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE + dv, vS, rho) + EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE - dv, vS, rho)) / 100
    ElseIf OutPutFlag = "dddvS" Then 'DDeltaDvol S vS
        EEquityLinkedFXO = 1 / (4 * dS * dv) * (EquityLinkedFXO(CallPutFlag, E, S + dS, X, T, r, rf, q, vE, vS + dv, rho) - EquityLinkedFXO(CallPutFlag, E, S + dS, X, T, r, rf, q, vE, vS - dv, rho) _
        - EquityLinkedFXO(CallPutFlag, E, S - dS, X, T, r, rf, q, vE, vS + dv, rho) + EquityLinkedFXO(CallPutFlag, E, S - dS, X, T, r, rf, q, vE, vS - dv, rho)) / 100
    ElseIf OutPutFlag = "cvE" Then 'Cross vanna E vS
        EEquityLinkedFXO = 1 / (4 * dS * dv) * (EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE, vS + dv, rho) - EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE, vS - dv, rho) _
        - EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE, vS + dv, rho) + EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE, vS - dv, rho)) / 100
    ElseIf OutPutFlag = "cvS" Then 'Cross vanna S vE
        EEquityLinkedFXO = 1 / (4 * dS * dv) * (EquityLinkedFXO(CallPutFlag, E, S + dS, X, T, r, rf, q, vE + dv, vS, rho) - EquityLinkedFXO(CallPutFlag, E, S + dS, X, T, r, rf, q, vE - dv, vS, rho) _
        - EquityLinkedFXO(CallPutFlag, E, S - dS, X, T, r, rf, q, vE + dv, vS, rho) + EquityLinkedFXO(CallPutFlag, E, S - dS, X, T, r, rf, q, vE - dv, vS, rho)) / 100
    ElseIf OutPutFlag = "vE" Then 'Vega vE
         EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE + dv, vS, rho) - EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE - dv, vS, rho)) / 2
    ElseIf OutPutFlag = "vS" Then 'Vega vS
         EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS + dv, rho) - EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS - dv, rho)) / 2
    ElseIf OutPutFlag = "dvdvE" Then 'DvegaDvol/vomma vE
        EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE + dv, vS, rho) - 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho) + EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE - dv, vS, rho)) / dv ^ 2 / 10000
    ElseIf OutPutFlag = "dvdvS" Then 'DvegaDvol/vomma vS
        EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS + dv, rho) - 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho) + EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS - dv, rho)) / dv ^ 2 / 10000
    ElseIf OutPutFlag = "cv" Then 'Cross vomma
        EEquityLinkedFXO = 1 / (4 * dv * dv) * (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE + dv, vS + dv, rho) - EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE - dv, vS + dv, rho) _
        - EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE + dv, vS - dv, rho) + EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE - dv, vS - dv, rho)) / 10000
    ElseIf OutPutFlag = "vpE" Then 'VegaP vE
         EEquityLinkedFXO = vE / 0.1 * (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE + dv, vS, rho) - EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE - dv, vS, rho)) / 2
    ElseIf OutPutFlag = "vpS" Then 'VegaP vS
         EEquityLinkedFXO = vS / 0.1 * (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS + dv, rho) - EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS - dv, rho)) / 2
    ElseIf OutPutFlag = "dvdvE" Then 'DvegaDvol vE
        EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE + dv, vS, rho) - 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho) + EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE - dv, vS, rho))
    ElseIf OutPutFlag = "dvdvS" Then 'DvegaDvol vS
        EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS + dv, rho) - 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho) + EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS - dv, rho))
    ElseIf OutPutFlag = "corr" Then 'Corr sensitivty
         EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho + 0.01) - EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho - 0.01)) / 2
    ElseIf OutPutFlag = "corrcorr" Then 'DrhoDrho
        EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho + 0.01) - 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho) + EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho - 0.01)) / 0.01 ^ 2 / 10000
   
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EEquityLinkedFXO = EquityLinkedFXO(CallPutFlag, E, S, X, 0.00001, r, rf, q, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho)
        Else
                EEquityLinkedFXO = EquityLinkedFXO(CallPutFlag, E, S, X, T - 1 / 365, r, rf, q, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho)
        End If
    ElseIf OutPutFlag = "r" Then 'Rho
         EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X, T, r + 0.01, rf, q, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E, S, X, T, r - 0.01, rf, q, vE, vS, rho)) / (2)
   ElseIf OutPutFlag = "f" Then 'Phi
         EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf + 0.01, q, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf - 0.01, q, vE, vS, rho)) / (2)
  
  ElseIf OutPutFlag = "dr" Then 'Dividend rho
         EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q - 0.01, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q + 0.01, vE, vS, rho)) / (2)
      ElseIf OutPutFlag = "sE" Then 'Speed E
        EEquityLinkedFXO = 1 / dS ^ 3 * (EquityLinkedFXO(CallPutFlag, E + 2 * dS, S, X, T, r, rf, q, vE, vS, rho) - 3 * EquityLinkedFXO(CallPutFlag, E + dS, S, X, T, r, rf, q, vE, vS, rho) _
                                + 3 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E - dS, S, X, T, r, rf, q, vE, vS, rho))
      ElseIf OutPutFlag = "csE" Then 'Cross speed E^2S
        EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E + dS, S + dS, X, T, r, rf, q, vE, vS, rho) - 2 * EquityLinkedFXO(CallPutFlag, E, S + dS, X, T, r, rf, q, vE, vS, rho) + EquityLinkedFXO(CallPutFlag, E - dS, S + dS, X, T, r, rf, q, vE, vS, rho) _
        - EquityLinkedFXO(CallPutFlag, E + dS, S - dS, X, T, r, rf, q, vE, vS, rho) + 2 * EquityLinkedFXO(CallPutFlag, E, S - dS, X, T, r, rf, q, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E - dS, S - dS, X, T, r, rf, q, vE, vS, rho)) / (2 * dS ^ 3)
   
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X + dS, T, r, rf, q, vE, vS, rho) - EquityLinkedFXO(CallPutFlag, E, S, X - dS, T, r, rf, q, vE, vS, rho)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        EEquityLinkedFXO = (EquityLinkedFXO(CallPutFlag, E, S, X + dS, T, r, rf, q, vE, vS, rho) - 2 * EquityLinkedFXO(CallPutFlag, E, S, X, T, r, rf, q, vE, vS, rho) + EquityLinkedFXO(CallPutFlag, E, S, X - dS, T, r, rf, q, vE, vS, rho)) / dS ^ 2
    End If
    
End Function

'// Equity linked foreign exchange option
Public Function EquityLinkedFXO(CallPutFlag As String, E As Double, S As Double, X As Double, T As Double, r As Double, _
                rf As Double, q As Double, vS As Double, vE As Double, rho As Double) As Double

    Dim d1 As Double, d2 As Double
    
    d1 = (Log(E / X) + (r - rf + rho * vS * vE + vE ^ 2 / 2) * T) / (vE * Sqr(T))
    d2 = d1 - vE * Sqr(T)
     
    If CallPutFlag = "c" Then
        EquityLinkedFXO = E * S * Exp(-q * T) * CND(d1) - X * S * Exp((rf - r - q - rho * vS * vE) * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
        EquityLinkedFXO = X * S * Exp((rf - r - q - rho * vS * vE) * T) * CND(-d2) - E * S * Exp(-q * T) * CND(-d1)
    End If
End Function



