Attribute VB_Name = "LogDerivatives"
Option Explicit

' Programmer Espen Gaarder Haug
' Copyright 2006, Espen Gaarder Haug

Public Function LogContract_LnS(S As Double, T As Double, r As Double, b As Double, v As Double)
        
        LogContract_LnS = Exp(-r * T) * (Log(S) + (b - v ^ 2 / 2) * T)
        
 End Function


Public Function ELogContract_LnS(OutPutFlag As String, S As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        ELogContract_LnS = LogContract_LnS(S, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         ELogContract_LnS = (LogContract_LnS(S + dS, T, r, b, v) - LogContract_LnS(S - dS, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         ELogContract_LnS = (LogContract_LnS(S + dS, T, r, b, v) - LogContract_LnS(S - dS, T, r, b, v)) / (2 * dS) * S / LogContract_LnS(S, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ELogContract_LnS = (LogContract_LnS(S + dS, T, r, b, v) - 2 * LogContract_LnS(S, T, r, b, v) + LogContract_LnS(S - dS, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ELogContract_LnS = (LogContract_LnS(S + dS, T, r, b, v + 0.01) - 2 * LogContract_LnS(S, T, r, b, v + 0.01) + LogContract_LnS(S - dS, T, r, b, v + 0.01) _
      - LogContract_LnS(S + dS, T, r, b, v - 0.01) + 2 * LogContract_LnS(S, T, r, b, v - 0.01) - LogContract_LnS(S - dS, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        ELogContract_LnS = S / 100 * (LogContract_LnS(S + dS, T, r, b, v) - 2 * LogContract_LnS(S, T, r, b, v) + LogContract_LnS(S - dS, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        ELogContract_LnS = (LogContract_LnS(S, T + 1 / 365, r, b, v) - 2 * LogContract_LnS(S, T, r, b, v) + LogContract_LnS(S, T - 1 / 365, r, b, v)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ELogContract_LnS = 1 / (4 * dS * 0.01) * (LogContract_LnS(S + dS, T, r, b, v + 0.01) - LogContract_LnS(S + dS, T, r, b, v - 0.01) _
        - LogContract_LnS(S - dS, T, r, b, v + 0.01) + LogContract_LnS(S - dS, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ELogContract_LnS = (LogContract_LnS(S, T, r, b, v + 0.01) - LogContract_LnS(S, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        ELogContract_LnS = (LogContract_LnS(S, T, r, b, v + 0.01) - 2 * LogContract_LnS(S, T, r, b, v) + LogContract_LnS(S, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ELogContract_LnS = v / 0.1 * (LogContract_LnS(S, T, r, b, v + 0.01) - LogContract_LnS(S, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ELogContract_LnS = (LogContract_LnS(S, T, r, b, v + 0.01) - 2 * LogContract_LnS(S, T, r, b, v) + LogContract_LnS(S, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ELogContract_LnS = LogContract_LnS(S, 0.00001, r, b, v) - LogContract_LnS(S, T, r, b, v)
        Else
                ELogContract_LnS = LogContract_LnS(S, T - 1 / 365, r, b, v) - LogContract_LnS(S, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         ELogContract_LnS = (LogContract_LnS(S, T, r + 0.01, b + 0.01, v) - LogContract_LnS(S, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ELogContract_LnS = (LogContract_LnS(S, T, r + 0.01, b, v) - LogContract_LnS(S, T, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         ELogContract_LnS = (LogContract_LnS(S, T, r, b - 0.01, v) - LogContract_LnS(S, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        ELogContract_LnS = (LogContract_LnS(S, T, r, b + 0.01, v) - LogContract_LnS(S, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ELogContract_LnS = 1 / dS ^ 3 * (LogContract_LnS(S + 2 * dS, T, r, b, v) - 3 * LogContract_LnS(S + dS, T, r, b, v) _
                                + 3 * LogContract_LnS(S, T, r, b, v) - LogContract_LnS(S - dS, T, r, b, v))
     End If
End Function





Public Function LogContract_lnSX(S As Double, X As Double, T As Double, r As Double, b As Double, v As Double)
        
        LogContract_lnSX = Exp(-r * T) * (Log(S / X) + (b - v ^ 2 / 2) * T)
        
 End Function


Public Function ELogContract_lnSX(OutPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        ELogContract_lnSX = LogContract_lnSX(S, X, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         ELogContract_lnSX = (LogContract_lnSX(S + dS, X, T, r, b, v) - LogContract_lnSX(S - dS, X, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         ELogContract_lnSX = (LogContract_lnSX(S + dS, X, T, r, b, v) - LogContract_lnSX(S - dS, X, T, r, b, v)) / (2 * dS) * S / LogContract_lnSX(S, X, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ELogContract_lnSX = (LogContract_lnSX(S + dS, X, T, r, b, v) - 2 * LogContract_lnSX(S, X, T, r, b, v) + LogContract_lnSX(S - dS, X, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ELogContract_lnSX = (LogContract_lnSX(S + dS, X, T, r, b, v + 0.01) - 2 * LogContract_lnSX(S, X, T, r, b, v + 0.01) + LogContract_lnSX(S - dS, X, T, r, b, v + 0.01) _
      - LogContract_lnSX(S + dS, X, T, r, b, v - 0.01) + 2 * LogContract_lnSX(S, X, T, r, b, v - 0.01) - LogContract_lnSX(S - dS, X, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        ELogContract_lnSX = S / 100 * (LogContract_lnSX(S + dS, X, T, r, b, v) - 2 * LogContract_lnSX(S, X, T, r, b, v) + LogContract_lnSX(S - dS, X, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        ELogContract_lnSX = (LogContract_lnSX(S, X, T + 1 / 365, r, b, v) - 2 * LogContract_lnSX(S, X, T, r, b, v) + LogContract_lnSX(S, X, T - 1 / 365, r, b, v)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ELogContract_lnSX = 1 / (4 * dS * 0.01) * (LogContract_lnSX(S + dS, X, T, r, b, v + 0.01) - LogContract_lnSX(S + dS, X, T, r, b, v - 0.01) _
        - LogContract_lnSX(S - dS, X, T, r, b, v + 0.01) + LogContract_lnSX(S - dS, X, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ELogContract_lnSX = (LogContract_lnSX(S, X, T, r, b, v + 0.01) - LogContract_lnSX(S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        ELogContract_lnSX = (LogContract_lnSX(S, X, T, r, b, v + 0.01) - 2 * LogContract_lnSX(S, X, T, r, b, v) + LogContract_lnSX(S, X, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ELogContract_lnSX = v / 0.1 * (LogContract_lnSX(S, X, T, r, b, v + 0.01) - LogContract_lnSX(S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ELogContract_lnSX = (LogContract_lnSX(S, X, T, r, b, v + 0.01) - 2 * LogContract_lnSX(S, X, T, r, b, v) + LogContract_lnSX(S, X, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ELogContract_lnSX = LogContract_lnSX(S, X, 0.00001, r, b, v) - LogContract_lnSX(S, X, T, r, b, v)
        Else
                ELogContract_lnSX = LogContract_lnSX(S, X, T - 1 / 365, r, b, v) - LogContract_lnSX(S, X, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         ELogContract_lnSX = (LogContract_lnSX(S, X, T, r + 0.01, b + 0.01, v) - LogContract_lnSX(S, X, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ELogContract_lnSX = (LogContract_lnSX(S, X, T, r + 0.01, b, v) - LogContract_lnSX(S, X, T, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         ELogContract_lnSX = (LogContract_lnSX(S, X, T, r, b - 0.01, v) - LogContract_lnSX(S, X, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        ELogContract_lnSX = (LogContract_lnSX(S, X, T, r, b + 0.01, v) - LogContract_lnSX(S, X, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ELogContract_lnSX = 1 / dS ^ 3 * (LogContract_lnSX(S + 2 * dS, X, T, r, b, v) - 3 * LogContract_lnSX(S + dS, X, T, r, b, v) _
                                + 3 * LogContract_lnSX(S, X, T, r, b, v) - LogContract_lnSX(S - dS, X, T, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ELogContract_lnSX = (LogContract_lnSX(S, X + dS, T, r, b, v) - LogContract_lnSX(S, X - dS, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        ELogContract_lnSX = (LogContract_lnSX(S, X + dS, T, r, b, v) - 2 * LogContract_lnSX(S, X, T, r, b, v) + LogContract_lnSX(S, X - dS, T, r, b, v)) / dS ^ 2
    End If
End Function

  Public Function LogOption(S As Double, X _
                As Double, T As Double, r As Double, b As Double, v As Double) As Double

    Dim d2 As Double
    
    d2 = (Log(S / X) + (b - v ^ 2 / 2) * T) / (v * Sqr(T))
  
    LogOption = Exp(-r * T) * ND(d2) * v * Sqr(T) + Exp(-r * T) * (Log(S / X) + (b - v ^ 2 / 2) * T) * CND(d2)
   
End Function

Public Function ELogOption(OutPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        ELogOption = LogOption(S, X, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         ELogOption = (LogOption(S + dS, X, T, r, b, v) - LogOption(S - dS, X, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         ELogOption = (LogOption(S + dS, X, T, r, b, v) - LogOption(S - dS, X, T, r, b, v)) / (2 * dS) * S / LogOption(S, X, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ELogOption = (LogOption(S + dS, X, T, r, b, v) - 2 * LogOption(S, X, T, r, b, v) + LogOption(S - dS, X, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ELogOption = (LogOption(S + dS, X, T, r, b, v + 0.01) - 2 * LogOption(S, X, T, r, b, v + 0.01) + LogOption(S - dS, X, T, r, b, v + 0.01) _
      - LogOption(S + dS, X, T, r, b, v - 0.01) + 2 * LogOption(S, X, T, r, b, v - 0.01) - LogOption(S - dS, X, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "gp" Then 'GammaP
        ELogOption = S / 100 * (LogOption(S + dS, X, T, r, b, v) - 2 * LogOption(S, X, T, r, b, v) + LogOption(S - dS, X, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        ELogOption = (LogOption(S, X, T + 1 / 365, r, b, v) - 2 * LogOption(S, X, T, r, b, v) + LogOption(S, X, T - 1 / 365, r, b, v)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ELogOption = 1 / (4 * dS * 0.01) * (LogOption(S + dS, X, T, r, b, v + 0.01) - LogOption(S + dS, X, T, r, b, v - 0.01) _
        - LogOption(S - dS, X, T, r, b, v + 0.01) + LogOption(S - dS, X, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ELogOption = (LogOption(S, X, T, r, b, v + 0.01) - LogOption(S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        ELogOption = (LogOption(S, X, T, r, b, v + 0.01) - 2 * LogOption(S, X, T, r, b, v) + LogOption(S, X, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ELogOption = v / 0.1 * (LogOption(S, X, T, r, b, v + 0.01) - LogOption(S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ELogOption = (LogOption(S, X, T, r, b, v + 0.01) - 2 * LogOption(S, X, T, r, b, v) + LogOption(S, X, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ELogOption = LogOption(S, X, 0.00001, r, b, v) - LogOption(S, X, T, r, b, v)
        Else
                ELogOption = LogOption(S, X, T - 1 / 365, r, b, v) - LogOption(S, X, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         ELogOption = (LogOption(S, X, T, r + 0.01, b + 0.01, v) - LogOption(S, X, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ELogOption = (LogOption(S, X, T, r + 0.01, b, v) - LogOption(S, X, T, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         ELogOption = (LogOption(S, X, T, r, b - 0.01, v) - LogOption(S, X, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        ELogOption = (LogOption(S, X, T, r, b + 0.01, v) - LogOption(S, X, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ELogOption = 1 / dS ^ 3 * (LogOption(S + 2 * dS, X, T, r, b, v) - 3 * LogOption(S + dS, X, T, r, b, v) _
                                + 3 * LogOption(S, X, T, r, b, v) - LogOption(S - dS, X, T, r, b, v))
    ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ELogOption = (LogOption(S, X + dS, T, r, b, v) - LogOption(S, X - dS, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        ELogOption = (LogOption(S, X + dS, T, r, b, v) - 2 * LogOption(S, X, T, r, b, v) + LogOption(S, X - dS, T, r, b, v)) / dS ^ 2
    End If
End Function
