Attribute VB_Name = "MultiTime"
Option Explicit


' Programmer Espen Gaarder Haug
' Copyright Espen Gaarder Haug 2006

'// Forward start options
Public Function ForwardStartOption(CallPutFlag As String, S As Double, Alpha As Double, t1 As Double, _
                T As Double, r As Double, b As Double, v As Double) As Double

    ForwardStartOption = S * Exp((b - r) * t1) * GBlackScholes(CallPutFlag, 1, Alpha, T - t1, r, b, v)

End Function

Public Function EForwardStartOption(OutPutFlag As String, CallPutFlag As String, S As Double, Alpha As Double, t1 As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        EForwardStartOption = ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EForwardStartOption = (ForwardStartOption(CallPutFlag, S + dS, Alpha, t1, T, r, b, v) - ForwardStartOption(CallPutFlag, S - dS, Alpha, t1, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EForwardStartOption = (ForwardStartOption(CallPutFlag, S + dS, Alpha, t1, T, r, b, v) - 2 * ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v) + ForwardStartOption(CallPutFlag, S - dS, Alpha, t1, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EForwardStartOption = (ForwardStartOption(CallPutFlag, S + dS, Alpha, t1, T, r, b, v + 0.01) - 2 * ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v + 0.01) + ForwardStartOption(CallPutFlag, S - dS, Alpha, t1, T, r, b, v + 0.01) _
      - ForwardStartOption(CallPutFlag, S + dS, Alpha, t1, T, r, b, v - 0.01) + 2 * ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v - 0.01) - ForwardStartOption(CallPutFlag, S - dS, Alpha, t1, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EForwardStartOption = 1 / (4 * dS * 0.01) * (ForwardStartOption(CallPutFlag, S + dS, Alpha, t1, T, r, b, v + 0.01) - ForwardStartOption(CallPutFlag, S + dS, Alpha, t1, T, r, b, v - 0.01) _
        - ForwardStartOption(CallPutFlag, S - dS, Alpha, t1, T, r, b, v + 0.01) + ForwardStartOption(CallPutFlag, S - dS, Alpha, t1, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EForwardStartOption = (ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v + 0.01) - ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v - 0.01)) / 2
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EForwardStartOption = v / 0.1 * (ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v + 0.01) - ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EForwardStartOption = (ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v + 0.01) - 2 * ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v) + ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
            EForwardStartOption = ForwardStartOption(CallPutFlag, S, Alpha, t1, 0.00001, r, b, v) - ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v)
        Else
                EForwardStartOption = ForwardStartOption(CallPutFlag, S, Alpha, t1, T - 1 / 365, r, b, v) - ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EForwardStartOption = (ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r + 0.01, b + 0.01, v) - ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures Rho
         EForwardStartOption = (ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r + 0.01, b, v) - ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r - 0.01, b, v)) / (2)
    ElseIf OutPutFlag = "f" Then 'Rho2
         EForwardStartOption = (ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b - 0.01, v) - ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EForwardStartOption = (ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b + 0.01, v) - ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EForwardStartOption = 1 / dS ^ 3 * (ForwardStartOption(CallPutFlag, S + 2 * dS, Alpha, t1, T, r, b, v) - 3 * ForwardStartOption(CallPutFlag, S + dS, Alpha, t1, T, r, b, v) _
                                + 3 * ForwardStartOption(CallPutFlag, S, Alpha, t1, T, r, b, v) - ForwardStartOption(CallPutFlag, S - dS, Alpha, t1, T, r, b, v))
    End If
    
End Function

'// Writer extendible options
Public Function ExtendibleWriter(CallPutFlag As String, S As Double, X1 As Double, X2 As Double, t1 As Double, _
                T2 As Double, r As Double, b As Double, v As Double) As Double

    Dim rho As Double, z1 As Double, z2 As Double
    rho = Sqr(t1 / T2)
    z1 = (Log(S / X2) + (b + v ^ 2 / 2) * T2) / (v * Sqr(T2))
    z2 = (Log(S / X1) + (b + v ^ 2 / 2) * t1) / (v * Sqr(t1))

    If CallPutFlag = "c" Then
        ExtendibleWriter = GBlackScholes(CallPutFlag, S, X1, t1, r, b, v) + S * Exp((b - r) * T2) * CBND(z1, -z2, -rho) - X2 * Exp(-r * T2) * CBND(z1 - Sqr(v ^ 2 * T2), -z2 + Sqr(v ^ 2 * t1), -rho)
    ElseIf CallPutFlag = "p" Then
        ExtendibleWriter = GBlackScholes(CallPutFlag, S, X1, t1, r, b, v) + X2 * Exp(-r * T2) * CBND(-z1 + Sqr(v ^ 2 * T2), z2 - Sqr(v ^ 2 * t1), -rho) - S * Exp((b - r) * T2) * CBND(-z1, z2, -rho)
    End If
End Function


Public Function EExtendibleWriter(OutPutFlag As String, CallPutFlag As String, S As Double, X1 As Double, X2 As Double, t1 As Double, T2 As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EExtendibleWriter = ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EExtendibleWriter = (ExtendibleWriter(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v) - ExtendibleWriter(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EExtendibleWriter = (ExtendibleWriter(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v) - ExtendibleWriter(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v)) / (2 * dS) * S / ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EExtendibleWriter = (ExtendibleWriter(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v) - 2 * ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v) + ExtendibleWriter(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EExtendibleWriter = (ExtendibleWriter(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v + 0.01) - 2 * ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v + 0.01) + ExtendibleWriter(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v + 0.01) _
      - ExtendibleWriter(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v - 0.01) + 2 * ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v - 0.01) - ExtendibleWriter(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EExtendibleWriter = S / 100 * (ExtendibleWriter(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v) - 2 * ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v) + ExtendibleWriter(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EExtendibleWriter = 1 / (4 * dS * 0.01) * (ExtendibleWriter(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v + 0.01) - ExtendibleWriter(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v - 0.01) _
        - ExtendibleWriter(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v + 0.01) + ExtendibleWriter(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EExtendibleWriter = (ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v + 0.01) - ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EExtendibleWriter = (ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v + 0.01) - 2 * ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v) + ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EExtendibleWriter = v / 0.1 * (ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v + 0.01) - ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EExtendibleWriter = (ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v + 0.01) - 2 * ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v) + ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If t1 <= 1 / 365 Then
                EExtendibleWriter = ExtendibleWriter(CallPutFlag, S, X1, X2, 0.00001, T2 - 1 / 365, r, b, v) - ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v)
        Else
                EExtendibleWriter = ExtendibleWriter(CallPutFlag, S, X1, X2, t1 - 1 / 365, T2 - 1 / 365, r, b, v) - ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EExtendibleWriter = (ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r + 0.01, b + 0.01, v) - ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EExtendibleWriter = (ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r + 0.01, b, v) - ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EExtendibleWriter = (ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b - 0.01, v) - ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EExtendibleWriter = (ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b + 0.01, v) - ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EExtendibleWriter = 1 / dS ^ 3 * (ExtendibleWriter(CallPutFlag, S + 2 * dS, X1, X2, t1, T2, r, b, v) - 3 * ExtendibleWriter(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v) _
                                + 3 * ExtendibleWriter(CallPutFlag, S, X1, X2, t1, T2, r, b, v) - ExtendibleWriter(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v))
      End If
End Function


' Reset Strike Option Type 1
Public Function ResetOptionGrayWhaleyT1(CallPutFlag As String, S As Double, X As Double, tau As Double, T As Double, _
                r As Double, b As Double, v As Double) As Double

    Dim a1 As Double, a2 As Double, z1 As Double, z2 As Double, y1 As Double, y2 As Double
    Dim rho As Double
    
        a1 = (Log(S / X) + (b + v ^ 2 / 2) * tau) / (v * Sqr(tau))
        a2 = a1 - v * Sqr(tau)
        y1 = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
        y2 = y1 - v * Sqr(T)
        z1 = (b + v ^ 2 / 2) * (T - tau) / (v * Sqr(T - tau))
        z2 = z1 - v * Sqr(T - tau)
        rho = Sqr(tau / T)
        
    If CallPutFlag = "c" Then
        ResetOptionGrayWhaleyT1 = Exp((b - r) * (T - tau)) * CND(-a2) * CND(z1) * Exp(-r * tau) - Exp(-r * T) * CND(-a2) * CND(z2) _
        - Exp(-r * T) * CBND(a2, y2, rho) + (S / X) * Exp((b - r) * T) * CBND(a1, y1, rho)
  
    ElseIf CallPutFlag = "p" Then
        ResetOptionGrayWhaleyT1 = Exp(-r * T) * CND(a2) * CND(-z2) - Exp((b - r) * (T - tau)) * CND(a2) * CND(-z1) * Exp(-r * tau) _
        + Exp(-r * T) * CBND(-a2, -y2, rho) - (S / X) * Exp((b - r) * T) * CBND(-a1, -y1, rho)
  
            End If
End Function

Public Function EResetOptionGrayWhaleyT1(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, tau As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EResetOptionGrayWhaleyT1 = ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S + dS, X, tau, T, r, b, v) - ResetOptionGrayWhaleyT1(CallPutFlag, S - dS, X, tau, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S + dS, X, tau, T, r, b, v) - ResetOptionGrayWhaleyT1(CallPutFlag, S - dS, X, tau, T, r, b, v)) / (2 * dS) * S / ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S + dS, X, tau, T, r, b, v) - 2 * ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v) + ResetOptionGrayWhaleyT1(CallPutFlag, S - dS, X, tau, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S + dS, X, tau, T, r, b, v + 0.01) - 2 * ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v + 0.01) + ResetOptionGrayWhaleyT1(CallPutFlag, S - dS, X, tau, T, r, b, v + 0.01) _
      - ResetOptionGrayWhaleyT1(CallPutFlag, S + dS, X, tau, T, r, b, v - 0.01) + 2 * ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v - 0.01) - ResetOptionGrayWhaleyT1(CallPutFlag, S - dS, X, tau, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EResetOptionGrayWhaleyT1 = S / 100 * (ResetOptionGrayWhaleyT1(CallPutFlag, S + dS, X, tau, T, r, b, v) - 2 * ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v) + ResetOptionGrayWhaleyT1(CallPutFlag, S - dS, X, tau, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EResetOptionGrayWhaleyT1 = 1 / (4 * dS * 0.01) * (ResetOptionGrayWhaleyT1(CallPutFlag, S + dS, X, tau, T, r, b, v + 0.01) - ResetOptionGrayWhaleyT1(CallPutFlag, S + dS, X, tau, T, r, b, v - 0.01) _
        - ResetOptionGrayWhaleyT1(CallPutFlag, S - dS, X, tau, T, r, b, v + 0.01) + ResetOptionGrayWhaleyT1(CallPutFlag, S - dS, X, tau, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v + 0.01) - ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v + 0.01) - 2 * ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v) + ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EResetOptionGrayWhaleyT1 = v / 0.1 * (ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v + 0.01) - ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v + 0.01) - 2 * ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v) + ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If tau <= 1 / 365 Then
                EResetOptionGrayWhaleyT1 = ResetOptionGrayWhaleyT1(CallPutFlag, S, X, 0, T - 1 / 365, r, b, v) - ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v)
        Else
                EResetOptionGrayWhaleyT1 = ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau - 1 / 365, T - 1 / 365, r, b, v) - ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r + 0.01, b + 0.01, v) - ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r + 0.01, b, v) - ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b - 0.01, v) - ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b + 0.01, v) - ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EResetOptionGrayWhaleyT1 = 1 / dS ^ 3 * (ResetOptionGrayWhaleyT1(CallPutFlag, S + 2 * dS, X, tau, T, r, b, v) - 3 * ResetOptionGrayWhaleyT1(CallPutFlag, S + dS, X, tau, T, r, b, v) _
                                + 3 * ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v) - ResetOptionGrayWhaleyT1(CallPutFlag, S - dS, X, tau, T, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S, X + dS, tau, T, r, b, v) - ResetOptionGrayWhaleyT1(CallPutFlag, S, X - dS, tau, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EResetOptionGrayWhaleyT1 = (ResetOptionGrayWhaleyT1(CallPutFlag, S, X + dS, tau, T, r, b, v) - 2 * ResetOptionGrayWhaleyT1(CallPutFlag, S, X, tau, T, r, b, v) + ResetOptionGrayWhaleyT1(CallPutFlag, S, X - dS, tau, T, r, b, v)) / dS ^ 2
    End If
End Function



Public Function ResetOptionGrayWhaley(CallPutFlag As String, S As Double, X As Double, tau As Double, T As Double, _
                r As Double, b As Double, v As Double) As Double

    Dim a1 As Double, a2 As Double, z1 As Double, z2 As Double, y1 As Double, y2 As Double
    Dim rho As Double
    
        a1 = (Log(S / X) + (b + v ^ 2 / 2) * tau) / (v * Sqr(tau))
        a2 = a1 - v * Sqr(tau)
        y1 = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
        y2 = y1 - v * Sqr(T)
        z1 = (b + v ^ 2 / 2) * (T - tau) / (v * Sqr(T - tau))
        z2 = z1 - v * Sqr(T - tau)
        rho = Sqr(tau / T)
        
    If CallPutFlag = "c" Then
        ResetOptionGrayWhaley = S * Exp((b - r) * T) * CBND(a1, y1, rho) - X * Exp(-r * T) * CBND(a2, y2, rho) _
                                - S * Exp((b - r) * tau) * CND(-a1) * CND(z2) * Exp(-r * (T - tau)) + _
                    S * Exp((b - r) * T) * CND(-a1) * CND(z1)
    ElseIf CallPutFlag = "p" Then
            ResetOptionGrayWhaley = S * Exp((b - r) * tau) * CND(a1) * CND(-z2) * Exp(-r * (T - tau)) - _
                    S * Exp((b - r) * T) * CND(a1) * CND(-z1) _
                    + X * Exp(-r * T) * CBND(-a2, -y2, rho) - S * Exp((b - r) * T) * CBND(-a1, -y1, rho)
    End If
End Function


Public Function EResetOptionGrayWhaley(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, tau As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EResetOptionGrayWhaley = ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S + dS, X, tau, T, r, b, v) - ResetOptionGrayWhaley(CallPutFlag, S - dS, X, tau, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S + dS, X, tau, T, r, b, v) - ResetOptionGrayWhaley(CallPutFlag, S - dS, X, tau, T, r, b, v)) / (2 * dS) * S / ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S + dS, X, tau, T, r, b, v) - 2 * ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v) + ResetOptionGrayWhaley(CallPutFlag, S - dS, X, tau, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S + dS, X, tau, T, r, b, v + 0.01) - 2 * ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v + 0.01) + ResetOptionGrayWhaley(CallPutFlag, S - dS, X, tau, T, r, b, v + 0.01) _
      - ResetOptionGrayWhaley(CallPutFlag, S + dS, X, tau, T, r, b, v - 0.01) + 2 * ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v - 0.01) - ResetOptionGrayWhaley(CallPutFlag, S - dS, X, tau, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EResetOptionGrayWhaley = S / 100 * (ResetOptionGrayWhaley(CallPutFlag, S + dS, X, tau, T, r, b, v) - 2 * ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v) + ResetOptionGrayWhaley(CallPutFlag, S - dS, X, tau, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EResetOptionGrayWhaley = 1 / (4 * dS * 0.01) * (ResetOptionGrayWhaley(CallPutFlag, S + dS, X, tau, T, r, b, v + 0.01) - ResetOptionGrayWhaley(CallPutFlag, S + dS, X, tau, T, r, b, v - 0.01) _
        - ResetOptionGrayWhaley(CallPutFlag, S - dS, X, tau, T, r, b, v + 0.01) + ResetOptionGrayWhaley(CallPutFlag, S - dS, X, tau, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v + 0.01) - ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v + 0.01) - 2 * ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v) + ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EResetOptionGrayWhaley = v / 0.1 * (ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v + 0.01) - ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v + 0.01) - 2 * ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v) + ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If tau <= 1 / 365 Then
                EResetOptionGrayWhaley = ResetOptionGrayWhaley(CallPutFlag, S, X, 0, T - 1 / 365, r, b, v) - ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v)
        Else
                EResetOptionGrayWhaley = ResetOptionGrayWhaley(CallPutFlag, S, X, tau - 1 / 365, T - 1 / 365, r, b, v) - ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r + 0.01, b + 0.01, v) - ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r + 0.01, b, v) - ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b - 0.01, v) - ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b + 0.01, v) - ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EResetOptionGrayWhaley = 1 / dS ^ 3 * (ResetOptionGrayWhaley(CallPutFlag, S + 2 * dS, X, tau, T, r, b, v) - 3 * ResetOptionGrayWhaley(CallPutFlag, S + dS, X, tau, T, r, b, v) _
                                + 3 * ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v) - ResetOptionGrayWhaley(CallPutFlag, S - dS, X, tau, T, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S, X + dS, tau, T, r, b, v) - ResetOptionGrayWhaley(CallPutFlag, S, X - dS, tau, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EResetOptionGrayWhaley = (ResetOptionGrayWhaley(CallPutFlag, S, X + dS, tau, T, r, b, v) - 2 * ResetOptionGrayWhaley(CallPutFlag, S, X, tau, T, r, b, v) + ResetOptionGrayWhaley(CallPutFlag, S, X - dS, tau, T, r, b, v)) / dS ^ 2
    End If
End Function



'// Options on options approximation
Public Function OptionsOnOptionsApprox(TypeFlag As String, S As Double, X1 As Double, X2 As Double, t1 As Double, _
    T2 As Double, r As Double, b As Double, v As Double) As Double
    
    Dim OptionVol As Double, CGBS As Double
    Dim d1 As Double, d2 As Double
    Dim CallPutFlag As String
    
    If TypeFlag = "cc" Or TypeFlag = "pc" Then
        CallPutFlag = "c"
    Else
        CallPutFlag = "p"
    End If
    
    CGBS = GBlackScholes(CallPutFlag, S, X1, T2, r, b, v)
    
    OptionVol = v * Abs(GDelta(CallPutFlag, S, X1, T2, r, b, v)) * S / CGBS
    
    d1 = (Log(CGBS / X2) + (b + OptionVol ^ 2 / 2) * t1) / (OptionVol * Sqr(t1))
    d2 = d1 - OptionVol * Sqr(t1)
     If TypeFlag = "cc" Or TypeFlag = "cp" Then
        OptionsOnOptionsApprox = CGBS * CND(d1) - X2 * Exp(-r * t1) * CND(d2)
    ElseIf TypeFlag = "pc" Or TypeFlag = "pp" Then
        OptionsOnOptionsApprox = X2 * Exp(-r * t1) * CND(-d2) - CGBS * CND(-d1)
    End If
End Function
                


'// Options on options
Public Function OptionsOnOptions(TypeFlag As String, S As Double, X1 As Double, X2 As Double, t1 As Double, _
                T2 As Double, r As Double, b As Double, v As Double) As Double

    Dim y1 As Double, y2 As Double, z1 As Double, z2 As Double
    Dim i As Double, rho As Double, CallPutFlag As String
    
    If TypeFlag = "cc" Or TypeFlag = "pc" Then
        CallPutFlag = "c"
    Else
        CallPutFlag = "p"
    End If
    
    i = CriticalValueOptionsOnOptions(CallPutFlag, X1, X2, T2 - t1, r, b, v)
    
    rho = Sqr(t1 / T2)
    y1 = (Log(S / i) + (b + v ^ 2 / 2) * t1) / (v * Sqr(t1))
    y2 = y1 - v * Sqr(t1)
    z1 = (Log(S / X1) + (b + v ^ 2 / 2) * T2) / (v * Sqr(T2))
    z2 = z1 - v * Sqr(T2)

    If TypeFlag = "cc" Then
        OptionsOnOptions = S * Exp((b - r) * T2) * CBND(z1, y1, rho) - X1 * Exp(-r * T2) * CBND(z2, y2, rho) - X2 * Exp(-r * t1) * CND(y2)
    ElseIf TypeFlag = "pc" Then
        OptionsOnOptions = X1 * Exp(-r * T2) * CBND(z2, -y2, -rho) - S * Exp((b - r) * T2) * CBND(z1, -y1, -rho) + X2 * Exp(-r * t1) * CND(-y2)
    ElseIf TypeFlag = "cp" Then
        OptionsOnOptions = X1 * Exp(-r * T2) * CBND(-z2, -y2, rho) - S * Exp((b - r) * T2) * CBND(-z1, -y1, rho) - X2 * Exp(-r * t1) * CND(-y2)
    ElseIf TypeFlag = "pp" Then
        OptionsOnOptions = S * Exp((b - r) * T2) * CBND(-z1, y1, -rho) - X1 * Exp(-r * T2) * CBND(-z2, y2, -rho) + Exp(-r * t1) * X2 * CND(y2)
    End If
End Function

'// Calculation of critical price options on options
Private Function CriticalValueOptionsOnOptions(CallPutFlag As String, X1 As Double, X2 As Double, T As Double, _
                r As Double, b As Double, v As Double) As Double

    Dim Si As Double, ci As Double, di As Double, epsilon As Double
    
    Si = X1
    ci = GBlackScholes(CallPutFlag, Si, X1, T, r, b, v)
    di = GDelta(CallPutFlag, Si, X1, T, r, b, v)
    epsilon = 0.000001
    '// Newton-Raphson algorithm
    While Abs(ci - X2) > epsilon
        Si = Si - (ci - X2) / di
        ci = GBlackScholes(CallPutFlag, Si, X1, T, r, b, v)
        di = GDelta(CallPutFlag, Si, X1, T, r, b, v)
    Wend
    CriticalValueOptionsOnOptions = Si
End Function



Public Function EOptionsOnOptions(OutPutFlag As String, CallPutFlag As String, S As Double, X1 As Double, X2 As Double, t1 As Double, T2 As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EOptionsOnOptions = OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v) - OptionsOnOptions(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v) - OptionsOnOptions(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v)) / (2 * dS) * S / OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v) - 2 * OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v) + OptionsOnOptions(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v + 0.01) - 2 * OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v + 0.01) + OptionsOnOptions(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v + 0.01) _
      - OptionsOnOptions(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v - 0.01) + 2 * OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v - 0.01) - OptionsOnOptions(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EOptionsOnOptions = S / 100 * (OptionsOnOptions(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v) - 2 * OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v) + OptionsOnOptions(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EOptionsOnOptions = 1 / (4 * dS * 0.01) * (OptionsOnOptions(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v + 0.01) - OptionsOnOptions(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v - 0.01) _
        - OptionsOnOptions(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v + 0.01) + OptionsOnOptions(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v + 0.01) - OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v + 0.01) - 2 * OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v) + OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EOptionsOnOptions = v / 0.1 * (OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v + 0.01) - OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v + 0.01) - 2 * OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v) + OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If t1 <= 1 / 365 Then
                EOptionsOnOptions = OptionsOnOptions(CallPutFlag, S, X1, X2, 0.0000001, T2 - 1 / 365, r, b, v) - OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v)
        Else
                EOptionsOnOptions = OptionsOnOptions(CallPutFlag, S, X1, X2, t1 - 1 / 365, T2 - 1 / 365, r, b, v) - OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r + 0.01, b + 0.01, v) - OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r + 0.01, b, v) - OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b - 0.01, v) - OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b + 0.01, v) - OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EOptionsOnOptions = 1 / dS ^ 3 * (OptionsOnOptions(CallPutFlag, S + 2 * dS, X1, X2, t1, T2, r, b, v) - 3 * OptionsOnOptions(CallPutFlag, S + dS, X1, X2, t1, T2, r, b, v) _
                                + 3 * OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v) - OptionsOnOptions(CallPutFlag, S - dS, X1, X2, t1, T2, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S, X1 + dS, X2, t1, T2, r, b, v) - OptionsOnOptions(CallPutFlag, S, X1 - dS, X2, t1, T2, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EOptionsOnOptions = (OptionsOnOptions(CallPutFlag, S, X1 + dS, X2, t1, T2, r, b, v) - 2 * OptionsOnOptions(CallPutFlag, S, X1, X2, t1, T2, r, b, v) + OptionsOnOptions(CallPutFlag, S, X1 - dS, X2, t1, T2, r, b, v)) / dS ^ 2
    End If
End Function


'// Simple chooser options
Public Function SimpleChooser(S As Double, X As Double, t1 As Double, T2 As Double, _
                r As Double, b As Double, v As Double) As Double

    Dim d As Double, y As Double

    d = (Log(S / X) + (b + v ^ 2 / 2) * T2) / (v * Sqr(T2))
    y = (Log(S / X) + b * T2 + v ^ 2 * t1 / 2) / (v * Sqr(t1))
  
    SimpleChooser = S * Exp((b - r) * T2) * CND(d) - X * Exp(-r * T2) * CND(d - v * Sqr(T2)) _
    - S * Exp((b - r) * T2) * CND(-y) + X * Exp(-r * T2) * CND(-y + v * Sqr(t1))
End Function


Public Function ESimpleChooser(OutPutFlag As String, S As Double, X As Double, t1 As Double, T2 As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        ESimpleChooser = SimpleChooser(S, X, t1, T2, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         ESimpleChooser = (SimpleChooser(S + dS, X, t1, T2, r, b, v) - SimpleChooser(S - dS, X, t1, T2, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         ESimpleChooser = (SimpleChooser(S + dS, X, t1, T2, r, b, v) - SimpleChooser(S - dS, X, t1, T2, r, b, v)) / (2 * dS) * S / SimpleChooser(S, X, t1, T2, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ESimpleChooser = (SimpleChooser(S + dS, X, t1, T2, r, b, v) - 2 * SimpleChooser(S, X, t1, T2, r, b, v) + SimpleChooser(S - dS, X, t1, T2, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ESimpleChooser = (SimpleChooser(S + dS, X, t1, T2, r, b, v + 0.01) - 2 * SimpleChooser(S, X, t1, T2, r, b, v + 0.01) + SimpleChooser(S - dS, X, t1, T2, r, b, v + 0.01) _
      - SimpleChooser(S + dS, X, t1, T2, r, b, v - 0.01) + 2 * SimpleChooser(S, X, t1, T2, r, b, v - 0.01) - SimpleChooser(S - dS, X, t1, T2, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        ESimpleChooser = S / 100 * (SimpleChooser(S + dS, X, t1, T2, r, b, v) - 2 * SimpleChooser(S, X, t1, T2, r, b, v) + SimpleChooser(S - dS, X, t1, T2, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ESimpleChooser = 1 / (4 * dS * 0.01) * (SimpleChooser(S + dS, X, t1, T2, r, b, v + 0.01) - SimpleChooser(S + dS, X, t1, T2, r, b, v - 0.01) _
        - SimpleChooser(S - dS, X, t1, T2, r, b, v + 0.01) + SimpleChooser(S - dS, X, t1, T2, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ESimpleChooser = (SimpleChooser(S, X, t1, T2, r, b, v + 0.01) - SimpleChooser(S, X, t1, T2, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        ESimpleChooser = (SimpleChooser(S, X, t1, T2, r, b, v + 0.01) - 2 * SimpleChooser(S, X, t1, T2, r, b, v) + SimpleChooser(S, X, t1, T2, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ESimpleChooser = v / 0.1 * (SimpleChooser(S, X, t1, T2, r, b, v + 0.01) - SimpleChooser(S, X, t1, T2, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ESimpleChooser = (SimpleChooser(S, X, t1, T2, r, b, v + 0.01) - 2 * SimpleChooser(S, X, t1, T2, r, b, v) + SimpleChooser(S, X, t1, T2, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If t1 <= 1 / 365 Then
                ESimpleChooser = SimpleChooser(S, X, 0.000001, T2 - 1 / 365, r, b, v) - SimpleChooser(S, X, t1, T2, r, b, v)
        Else
                ESimpleChooser = SimpleChooser(S, X, t1 - 1 / 365, T2 - 1 / 365, r, b, v) - SimpleChooser(S, X, t1, T2, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         ESimpleChooser = (SimpleChooser(S, X, t1, T2, r + 0.01, b + 0.01, v) - SimpleChooser(S, X, t1, T2, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ESimpleChooser = (SimpleChooser(S, X, t1, T2, r + 0.01, b, v) - SimpleChooser(S, X, t1, T2, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         ESimpleChooser = (SimpleChooser(S, X, t1, T2, r, b - 0.01, v) - SimpleChooser(S, X, t1, T2, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        ESimpleChooser = (SimpleChooser(S, X, t1, T2, r, b + 0.01, v) - SimpleChooser(S, X, t1, T2, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ESimpleChooser = 1 / dS ^ 3 * (SimpleChooser(S + 2 * dS, X, t1, T2, r, b, v) - 3 * SimpleChooser(S + dS, X, t1, T2, r, b, v) _
                                + 3 * SimpleChooser(S, X, t1, T2, r, b, v) - SimpleChooser(S - dS, X, t1, T2, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ESimpleChooser = (SimpleChooser(S, X + dS, t1, T2, r, b, v) - SimpleChooser(S, X - dS, t1, T2, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        ESimpleChooser = (SimpleChooser(S, X + dS, t1, T2, r, b, v) - 2 * SimpleChooser(S, X, t1, T2, r, b, v) + SimpleChooser(S, X - dS, t1, T2, r, b, v)) / dS ^ 2
    End If
End Function




'// Complex chooser options
Public Function ComplexChooser(S As Double, Xc As Double, Xp As Double, T As Double, Tc As Double, _
                Tp As Double, r As Double, b As Double, v As Double) As Double
    
    Dim d1 As Double, d2 As Double, y1 As Double, y2 As Double
    Dim rho1 As Double, rho2 As Double, i As Double

    i = CriticalValueChooser(S, Xc, Xp, T, Tc, Tp, r, b, v)
    d1 = (Log(S / i) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    y1 = (Log(S / Xc) + (b + v ^ 2 / 2) * Tc) / (v * Sqr(Tc))
    y2 = (Log(S / Xp) + (b + v ^ 2 / 2) * Tp) / (v * Sqr(Tp))
    rho1 = Sqr(T / Tc)
    rho2 = Sqr(T / Tp)
    
    ComplexChooser = S * Exp((b - r) * Tc) * CBND(d1, y1, rho1) - Xc * Exp(-r * Tc) * CBND(d2, y1 - v * Sqr(Tc), rho1) - S * Exp((b - r) * Tp) * CBND(-d1, -y2, rho2) + Xp * Exp(-r * Tp) * CBND(-d2, -y2 + v * Sqr(Tp), rho2)
End Function


'// Critical value complex chooser option
Private Function CriticalValueChooser(S As Double, Xc As Double, Xp As Double, T As Double, _
                Tc As Double, Tp As Double, r As Double, b As Double, v As Double) As Double

    Dim Sv As Double, ci As Double, Pi As Double, epsilon As Double
    Dim dc As Double, dp As Double, yi As Double, di As Double

    Sv = S
    
    ci = GBlackScholes("c", Sv, Xc, Tc - T, r, b, v)
    Pi = GBlackScholes("p", Sv, Xp, Tp - T, r, b, v)
    dc = GDelta("c", Sv, Xc, Tc - T, r, b, v)
    dp = GDelta("p", Sv, Xp, Tp - T, r, b, v)
    yi = ci - Pi
    di = dc - dp
    epsilon = 0.001
    'Newton-Raphson søkeprosess
    While Abs(yi) > epsilon
        Sv = Sv - yi / di
        ci = GBlackScholes("c", Sv, Xc, Tc - T, r, b, v)
        Pi = GBlackScholes("p", Sv, Xp, Tp - T, r, b, v)
        dc = GDelta("c", Sv, Xc, Tc - T, r, b, v)
        dp = GDelta("p", Sv, Xp, Tp - T, r, b, v)
        yi = ci - Pi
        di = dc - dp
    Wend
    CriticalValueChooser = Sv
End Function



Public Function EComplexChooser(OutPutFlag As String, S As Double, Xc As Double, Xp As Double, T As Double, Tc As Double, _
                Tp As Double, r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EComplexChooser = ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EComplexChooser = (ComplexChooser(S + dS, Xc, Xp, T, Tc, Tp, r, b, v) - ComplexChooser(S - dS, Xc, Xp, T, Tc, Tp, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EComplexChooser = (ComplexChooser(S + dS, Xc, Xp, T, Tc, Tp, r, b, v) - ComplexChooser(S - dS, Xc, Xp, T, Tc, Tp, r, b, v)) / (2 * dS) * S / ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EComplexChooser = (ComplexChooser(S + dS, Xc, Xp, T, Tc, Tp, r, b, v) - 2 * ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v) + ComplexChooser(S - dS, Xc, Xp, T, Tc, Tp, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EComplexChooser = (ComplexChooser(S + dS, Xc, Xp, T, Tc, Tp, r, b, v + 0.01) - 2 * ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v + 0.01) + ComplexChooser(S - dS, Xc, Xp, T, Tc, Tp, r, b, v + 0.01) _
      - ComplexChooser(S + dS, Xc, Xp, T, Tc, Tp, r, b, v - 0.01) + 2 * ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v - 0.01) - ComplexChooser(S - dS, Xc, Xp, T, Tc, Tp, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EComplexChooser = S / 100 * (ComplexChooser(S + dS, Xc, Xp, T, Tc, Tp, r, b, v) - 2 * ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v) + ComplexChooser(S - dS, Xc, Xp, T, Tc, Tp, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EComplexChooser = 1 / (4 * dS * 0.01) * (ComplexChooser(S + dS, Xc, Xp, T, Tc, Tp, r, b, v + 0.01) - ComplexChooser(S + dS, Xc, Xp, T, Tc, Tp, r, b, v - 0.01) _
        - ComplexChooser(S - dS, Xc, Xp, T, Tc, Tp, r, b, v + 0.01) + ComplexChooser(S - dS, Xc, Xp, T, Tc, Tp, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EComplexChooser = (ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v + 0.01) - ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EComplexChooser = (ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v + 0.01) - 2 * ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v) + ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EComplexChooser = v / 0.1 * (ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v + 0.01) - ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EComplexChooser = (ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v + 0.01) - 2 * ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v) + ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EComplexChooser = ComplexChooser(S, Xc, Xp, 0.00001, Tc - 1 / 365, Tp - 1 / 365, r, b, v) - ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v)
        Else
                EComplexChooser = ComplexChooser(S, Xc, Xp, T - 1 / 365, Tc - 1 / 365, Tp - 1 / 365, r, b, v) - ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EComplexChooser = (ComplexChooser(S, Xc, Xp, T, Tc, Tp, r + 0.01, b + 0.01, v) - ComplexChooser(S, Xc, Xp, T, Tc, Tp, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EComplexChooser = (ComplexChooser(S, Xc, Xp, T, Tc, Tp, r + 0.01, b, v) - ComplexChooser(S, Xc, Xp, T, Tc, Tp, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EComplexChooser = (ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b - 0.01, v) - ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EComplexChooser = (ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b + 0.01, v) - ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EComplexChooser = 1 / dS ^ 3 * (ComplexChooser(S + 2 * dS, Xc, Xp, T, Tc, Tp, r, b, v) - 3 * ComplexChooser(S + dS, Xc, Xp, T, Tc, Tp, r, b, v) _
                                + 3 * ComplexChooser(S, Xc, Xp, T, Tc, Tp, r, b, v) - ComplexChooser(S - dS, Xc, Xp, T, Tc, Tp, r, b, v))
    End If
End Function



