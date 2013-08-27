Attribute VB_Name = "BSDAdjustments"
Option Explicit


' Programmer Espen Gaarder Haug, Copyright 2006

Public Function LoWangOUVolatility(v As Double, rho As Double, dt As Double) As Double
                    
    If rho = 0 Then
        LoWangOUVolatility = v
    Else
        LoWangOUVolatility = Sqr(v ^ 2 / dt * Log(1 + 2 * rho) / ((1 + 2 * rho) ^ (1 / dt) - 1))
    End If
    
End Function

Public Function ECEVmodel(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, v As Double, Beta As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        ECEVmodel = CEVmodel(CallPutFlag, S, X, T, r, v, Beta)
    ElseIf OutPutFlag = "d" Then 'Delta
         ECEVmodel = (CEVmodel(CallPutFlag, S + dS, X, T, r, v, Beta) - CEVmodel(CallPutFlag, S - dS, X, T, r, v, Beta)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         ECEVmodel = (CEVmodel(CallPutFlag, S + dS, X, T, r, v, Beta) - CEVmodel(CallPutFlag, S - dS, X, T, r, v, Beta)) / (2 * dS) * S / CEVmodel(CallPutFlag, S, X, T, r, v, Beta)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ECEVmodel = (CEVmodel(CallPutFlag, S + dS, X, T, r, v, Beta) - 2 * CEVmodel(CallPutFlag, S, X, T, r, v, Beta) + CEVmodel(CallPutFlag, S - dS, X, T, r, v, Beta)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ECEVmodel = (CEVmodel(CallPutFlag, S + dS, X, T, r, v + 0.01, Beta) - 2 * CEVmodel(CallPutFlag, S, X, T, r, v + 0.01, Beta) + CEVmodel(CallPutFlag, S - dS, X, T, r, v + 0.01, Beta) _
      - CEVmodel(CallPutFlag, S + dS, X, T, r, v - 0.01, Beta) + 2 * CEVmodel(CallPutFlag, S, X, T, r, v - 0.01, Beta) - CEVmodel(CallPutFlag, S - dS, X, T, r, v - 0.01, Beta)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        ECEVmodel = S / 100 * (CEVmodel(CallPutFlag, S + dS, X, T, r, v, Beta) - 2 * CEVmodel(CallPutFlag, S, X, T, r, v, Beta) + CEVmodel(CallPutFlag, S - dS, X, T, r, v, Beta)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ECEVmodel = 1 / (4 * dS * 0.01) * (CEVmodel(CallPutFlag, S + dS, X, T, r, v + 0.01, Beta) - CEVmodel(CallPutFlag, S + dS, X, T, r, v - 0.01, Beta) _
        - CEVmodel(CallPutFlag, S - dS, X, T, r, v + 0.01, Beta) + CEVmodel(CallPutFlag, S - dS, X, T, r, v - 0.01, Beta)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ECEVmodel = (CEVmodel(CallPutFlag, S, X, T, r, v + 0.01, Beta) - CEVmodel(CallPutFlag, S, X, T, r, v - 0.01, Beta)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        ECEVmodel = (CEVmodel(CallPutFlag, S, X, T, r, v + 0.01, Beta) - 2 * CEVmodel(CallPutFlag, S, X, T, r, v, Beta) + CEVmodel(CallPutFlag, S, X, T, r, v - 0.01, Beta)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ECEVmodel = v / 0.1 * (CEVmodel(CallPutFlag, S, X, T, r, v + 0.01, Beta) - CEVmodel(CallPutFlag, S, X, T, r, v - 0.01, Beta)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ECEVmodel = (CEVmodel(CallPutFlag, S, X, T, r, v + 0.01, Beta) - 2 * CEVmodel(CallPutFlag, S, X, T, r, v, Beta) + CEVmodel(CallPutFlag, S, X, T, r, v - 0.01, Beta))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ECEVmodel = CEVmodel(CallPutFlag, S, X, 0.000001, r, v, Beta) - CEVmodel(CallPutFlag, S, X, T, r, v, Beta)
        Else
                ECEVmodel = CEVmodel(CallPutFlag, S, X, T - 1 / 365, r, v, Beta) - CEVmodel(CallPutFlag, S, X, T, r, v, Beta)
        End If
      ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ECEVmodel = (CEVmodel(CallPutFlag, S, X, T, r + 0.01, v, Beta) - CEVmodel(CallPutFlag, S, X, T, r - 0.01, v, Beta)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ECEVmodel = 1 / dS ^ 3 * (CEVmodel(CallPutFlag, S + 2 * dS, X, T, r, v, Beta) - 3 * CEVmodel(CallPutFlag, S + dS, X, T, r, v, Beta) _
                                + 3 * CEVmodel(CallPutFlag, S, X, T, r, v, Beta) - CEVmodel(CallPutFlag, S - dS, X, T, r, v, Beta))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ECEVmodel = (CEVmodel(CallPutFlag, S, X + dS, T, r, v, Beta) - CEVmodel(CallPutFlag, S, X - dS, T, r, v, Beta)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        ECEVmodel = (CEVmodel(CallPutFlag, S, X + dS, T, r, v, Beta) - 2 * CEVmodel(CallPutFlag, S, X, T, r, v, Beta) + CEVmodel(CallPutFlag, S, X - dS, T, r, v, Beta)) / dS ^ 2
    ElseIf OutPutFlag = "beta" Then 'Beta Delta
         ECEVmodel = (CEVmodel(CallPutFlag, S, X, T, r, v, Beta + 0.01) - CEVmodel(CallPutFlag, S, X, T, r, v, Beta - 0.01)) / (2 * 0.01) / 100
    End If
End Function


'// CEV Constant Elasticity of Variance model
Public Function CEVmodel(CallPutFlag As String, F As Double, X _
                As Double, T As Double, r As Double, v As Double, Beta As Double) As Double
                
    Dim d1 As Double, d2 As Double, vm As Double, ff As Double
    
    ff = 0.5 * (F + X)
    
    vm = v / ff ^ (1 - Beta) * (1 + (1 - Beta) * (2 + Beta) / 24 * ((F - X) / ff) ^ 2 _
            + (1 - Beta) ^ 2 / 24 * v ^ 2 * T / (ff ^ (2 - 2 * Beta)))
    
    d1 = (Log(F / X) + vm ^ 2 / 2 * T) / (vm * Sqr(T))
    d2 = d1 - vm * Sqr(T)
    If CallPutFlag = "c" Then
        CEVmodel = Exp(-r * T) * (F * CND(d1) - X * CND(d2))
    ElseIf CallPutFlag = "p" Then
        CEVmodel = Exp(-r * T) * (X * CND(-d2) - F * CND(-d1))
    End If
    
End Function



'// French (1984) adjusted Black and Scholes model for trading day volatility
Public Function French(CallPutFlag As String, S As Double, X As Double, t1 As Double, T As Double, _
                r As Double, b As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / X) + b * T + v ^ 2 / 2 * t1) / (v * Sqr(t1))
    d2 = d1 - v * Sqr(t1)
  
    If CallPutFlag = "c" Then
        French = S * Exp((b - r) * T) * CND(d1) - X * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
        French = X * Exp(-r * T) * CND(-d2) - S * Exp((b - r) * T) * CND(-d1)
    End If
End Function




Public Function EFrench(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, t1 As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EFrench = French(CallPutFlag, S, X, t1, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EFrench = (French(CallPutFlag, S + dS, X, t1, T, r, b, v) - French(CallPutFlag, S - dS, X, t1, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EFrench = (French(CallPutFlag, S + dS, X, t1, T, r, b, v) - French(CallPutFlag, S - dS, X, t1, T, r, b, v)) / (2 * dS) * S / French(CallPutFlag, S, X, t1, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EFrench = (French(CallPutFlag, S + dS, X, t1, T, r, b, v) - 2 * French(CallPutFlag, S, X, t1, T, r, b, v) + French(CallPutFlag, S - dS, X, t1, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EFrench = (French(CallPutFlag, S + dS, X, t1, T, r, b, v + 0.01) - 2 * French(CallPutFlag, S, X, t1, T, r, b, v + 0.01) + French(CallPutFlag, S - dS, X, t1, T, r, b, v + 0.01) _
      - French(CallPutFlag, S + dS, X, t1, T, r, b, v - 0.01) + 2 * French(CallPutFlag, S, X, t1, T, r, b, v - 0.01) - French(CallPutFlag, S - dS, X, t1, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EFrench = S / 100 * (French(CallPutFlag, S + dS, X, t1, T, r, b, v) - 2 * French(CallPutFlag, S, X, t1, T, r, b, v) + French(CallPutFlag, S - dS, X, t1, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EFrench = 1 / (4 * dS * 0.01) * (French(CallPutFlag, S + dS, X, t1, T, r, b, v + 0.01) - French(CallPutFlag, S + dS, X, t1, T, r, b, v - 0.01) _
        - French(CallPutFlag, S - dS, X, t1, T, r, b, v + 0.01) + French(CallPutFlag, S - dS, X, t1, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EFrench = (French(CallPutFlag, S, X, t1, T, r, b, v + 0.01) - French(CallPutFlag, S, X, t1, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EFrench = (French(CallPutFlag, S, X, t1, T, r, b, v + 0.01) - 2 * French(CallPutFlag, S, X, t1, T, r, b, v) + French(CallPutFlag, S, X, t1, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EFrench = v / 0.1 * (French(CallPutFlag, S, X, t1, T, r, b, v + 0.01) - French(CallPutFlag, S, X, t1, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EFrench = (French(CallPutFlag, S, X, t1, T, r, b, v + 0.01) - 2 * French(CallPutFlag, S, X, t1, T, r, b, v) + French(CallPutFlag, S, X, t1, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If t1 <= 1 / 365 Then
                EFrench = French(CallPutFlag, S, X, 0.000001, T - 1 / 365, r, b, v) - French(CallPutFlag, S, X, t1, T, r, b, v)
        Else
                EFrench = French(CallPutFlag, S, X, t1 - 1 / 365, T - 1 / 365, r, b, v) - French(CallPutFlag, S, X, t1, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EFrench = (French(CallPutFlag, S, X, t1, T, r + 0.01, b + 0.01, v) - French(CallPutFlag, S, X, t1, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EFrench = (French(CallPutFlag, S, X, t1, T, r + 0.01, b, v) - French(CallPutFlag, S, X, t1, T, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EFrench = (French(CallPutFlag, S, X, t1, T, r, b - 0.01, v) - French(CallPutFlag, S, X, t1, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EFrench = (French(CallPutFlag, S, X, t1, T, r, b + 0.01, v) - French(CallPutFlag, S, X, t1, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EFrench = 1 / dS ^ 3 * (French(CallPutFlag, S + 2 * dS, X, t1, T, r, b, v) - 3 * French(CallPutFlag, S + dS, X, t1, T, r, b, v) _
                                + 3 * French(CallPutFlag, S, X, t1, T, r, b, v) - French(CallPutFlag, S - dS, X, t1, T, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EFrench = (French(CallPutFlag, S, X + dS, t1, T, r, b, v) - French(CallPutFlag, S, X - dS, t1, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EFrench = (French(CallPutFlag, S, X + dS, t1, T, r, b, v) - 2 * French(CallPutFlag, S, X, t1, T, r, b, v) + French(CallPutFlag, S, X - dS, t1, T, r, b, v)) / dS ^ 2
    End If
End Function


'//  The generalized Black and Scholes formula
Public Function GBlackScholesSettlementAdj(CallPutFlag As String, S As Double, X As Double, _
                 t1 As Double, T2 As Double, T As Double, r1 As Double, r2 As Double, b As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    d1 = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then
        GBlackScholesSettlementAdj = Exp(r1 * t1) * Exp(-r2 * T2) * (S * Exp(b * T) * CND(d1) - X * CND(d2))
    ElseIf CallPutFlag = "p" Then
        GBlackScholesSettlementAdj = Exp(r1 * t1) * Exp(-r2 * T2) * (X * CND(-d2) - S * Exp(b * T) * CND(-d1))
    End If
End Function


Public Function EGBlackScholesSettlementAdj(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, t1 As Double, T2 As Double, _
              T As Double, r1 As Double, r2 As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EGBlackScholesSettlementAdj = GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S + dS, X, t1, T2, T, r1, r2, b, v) - GBlackScholesSettlementAdj(CallPutFlag, S - dS, X, t1, T2, T, r1, r2, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S + dS, X, t1, T2, T, r1, r2, b, v) - GBlackScholesSettlementAdj(CallPutFlag, S - dS, X, t1, T2, T, r1, r2, b, v)) / (2 * dS) * S / GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S + dS, X, t1, T2, T, r1, r2, b, v) - 2 * GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v) + GBlackScholesSettlementAdj(CallPutFlag, S - dS, X, t1, T2, T, r1, r2, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S + dS, X, t1, T2, T, r1, r2, b, v + 0.01) - 2 * GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v + 0.01) + GBlackScholesSettlementAdj(CallPutFlag, S - dS, X, t1, T2, T, r1, r2, b, v + 0.01) _
      - GBlackScholesSettlementAdj(CallPutFlag, S + dS, X, t1, T2, T, r1, r2, b, v - 0.01) + 2 * GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v - 0.01) - GBlackScholesSettlementAdj(CallPutFlag, S - dS, X, t1, T2, T, r1, r2, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EGBlackScholesSettlementAdj = S / 100 * (GBlackScholesSettlementAdj(CallPutFlag, S + dS, X, t1, T2, T, r1, r2, b, v) - 2 * GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v) + GBlackScholesSettlementAdj(CallPutFlag, S - dS, X, t1, T2, T, r1, r2, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EGBlackScholesSettlementAdj = 1 / (4 * dS * 0.01) * (GBlackScholesSettlementAdj(CallPutFlag, S + dS, X, t1, T2, T, r1, r2, b, v + 0.01) - GBlackScholesSettlementAdj(CallPutFlag, S + dS, X, t1, T2, T, r1, r2, b, v - 0.01) _
        - GBlackScholesSettlementAdj(CallPutFlag, S - dS, X, t1, T2, T, r1, r2, b, v + 0.01) + GBlackScholesSettlementAdj(CallPutFlag, S - dS, X, t1, T2, T, r1, r2, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v + 0.01) - GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v + 0.01) - 2 * GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v) + GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EGBlackScholesSettlementAdj = v / 0.1 * (GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v + 0.01) - GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v + 0.01) - 2 * GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v) + GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EGBlackScholesSettlementAdj = GBlackScholesSettlementAdj(CallPutFlag, S, X, 0.00001, T2 - 1 / 365, 0.00001, r1, r2, b, v) - GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v)
        Else
                EGBlackScholesSettlementAdj = GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2 - 1 / 365, T - 1 / 365, r1, r2, b, v) - GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1 + 0.01, r2 + 0.01, b + 0.01, v) - GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1 - 0.01, r2 - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1 + 0.01, r2 + 0.01, b, v) - GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1 - 0.01, r2 - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b - 0.01, v) - GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b + 0.01, v) - GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EGBlackScholesSettlementAdj = 1 / dS ^ 3 * (GBlackScholesSettlementAdj(CallPutFlag, S + 2 * dS, X, t1, T2, T, r1, r2, b, v) - 3 * GBlackScholesSettlementAdj(CallPutFlag, S + dS, X, t1, T2, T, r1, r2, b, v) _
                                + 3 * GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v) - GBlackScholesSettlementAdj(CallPutFlag, S - dS, X, t1, T2, T, r1, r2, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S, X + dS, t1, T2, T, r1, r2, b, v) - GBlackScholesSettlementAdj(CallPutFlag, S, X - dS, t1, T2, T, r1, r2, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EGBlackScholesSettlementAdj = (GBlackScholesSettlementAdj(CallPutFlag, S, X + dS, t1, T2, T, r1, r2, b, v) - 2 * GBlackScholesSettlementAdj(CallPutFlag, S, X, t1, T2, T, r1, r2, b, v) + GBlackScholesSettlementAdj(CallPutFlag, S, X - dS, t1, T2, T, r1, r2, b, v)) / dS ^ 2
    End If
End Function
