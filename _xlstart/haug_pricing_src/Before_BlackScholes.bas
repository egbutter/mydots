Attribute VB_Name = "Before_BlackScholes"
Option Explicit

' Programmer Espen Gaarder Haug
' Copyright Espen G. Haug


'//  Bachelier original 1900 formula
Public Function Bachelier(CallPutFlag As String, S As Double, x _
                As Double, T As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    d1 = (S - x) / (v * Sqr(T))
  

    If CallPutFlag = "c" Then
        Bachelier = (S - x) * CND(d1) + v * Sqr(T) * ND(d1)
    ElseIf CallPutFlag = "p" Then
        Bachelier = (x - S) * CND(-d1) + v * Sqr(T) * ND(d1)
    End If
    
End Function


Public Function EBachelier(OutPutFlag As String, CallPutFlag As String, S As Double, x As Double, T As Double, _
                 v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EBachelier = Bachelier(CallPutFlag, S, x, T, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EBachelier = (Bachelier(CallPutFlag, S + dS, x, T, v) - Bachelier(CallPutFlag, S - dS, x, T, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EBachelier = (Bachelier(CallPutFlag, S + dS, x, T, v) - Bachelier(CallPutFlag, S - dS, x, T, v)) / (2 * dS) * S / Bachelier(CallPutFlag, S, x, T, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EBachelier = (Bachelier(CallPutFlag, S + dS, x, T, v) - 2 * Bachelier(CallPutFlag, S, x, T, v) + Bachelier(CallPutFlag, S - dS, x, T, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EBachelier = (Bachelier(CallPutFlag, S + dS, x, T, v + 0.01) - 2 * Bachelier(CallPutFlag, S, x, T, v + 0.01) + Bachelier(CallPutFlag, S - dS, x, T, v + 0.01) _
      - Bachelier(CallPutFlag, S + dS, x, T, v - 0.01) + 2 * Bachelier(CallPutFlag, S, x, T, v - 0.01) - Bachelier(CallPutFlag, S - dS, x, T, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EBachelier = S / 100 * (Bachelier(CallPutFlag, S + dS, x, T, v) - 2 * Bachelier(CallPutFlag, S, x, T, v) + Bachelier(CallPutFlag, S - dS, x, T, v)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        EBachelier = (Bachelier(CallPutFlag, S, x, T + 1 / 365, v) - 2 * Bachelier(CallPutFlag, S, x, T, v) + Bachelier(CallPutFlag, S, x, T - 1 / 365, v)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EBachelier = 1 / (4 * dS * 0.01) * (Bachelier(CallPutFlag, S + dS, x, T, v + 0.01) - Bachelier(CallPutFlag, S + dS, x, T, v - 0.01) _
        - Bachelier(CallPutFlag, S - dS, x, T, v + 0.01) + Bachelier(CallPutFlag, S - dS, x, T, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EBachelier = (Bachelier(CallPutFlag, S, x, T, v + 0.01) - Bachelier(CallPutFlag, S, x, T, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EBachelier = (Bachelier(CallPutFlag, S, x, T, v + 0.01) - 2 * Bachelier(CallPutFlag, S, x, T, v) + Bachelier(CallPutFlag, S, x, T, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EBachelier = v / 0.1 * (Bachelier(CallPutFlag, S, x, T, v + 0.01) - Bachelier(CallPutFlag, S, x, T, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EBachelier = (Bachelier(CallPutFlag, S, x, T, v + 0.01) - 2 * Bachelier(CallPutFlag, S, x, T, v) + Bachelier(CallPutFlag, S, x, T, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EBachelier = Bachelier(CallPutFlag, S, x, 0.00001, v) - Bachelier(CallPutFlag, S, x, T, v)
        Else
                EBachelier = Bachelier(CallPutFlag, S, x, T - 1 / 365, v) - Bachelier(CallPutFlag, S, x, T, v)
        End If
      ElseIf OutPutFlag = "s" Then 'Speed
        EBachelier = 1 / dS ^ 3 * (Bachelier(CallPutFlag, S + 2 * dS, x, T, v) - 3 * Bachelier(CallPutFlag, S + dS, x, T, v) _
                                + 3 * Bachelier(CallPutFlag, S, x, T, v) - Bachelier(CallPutFlag, S - dS, x, T, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EBachelier = (Bachelier(CallPutFlag, S, x + dS, T, v) - Bachelier(CallPutFlag, S, x - dS, T, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EBachelier = (Bachelier(CallPutFlag, S, x + dS, T, v) - 2 * Bachelier(CallPutFlag, S, x, T, v) + Bachelier(CallPutFlag, S, x - dS, T, v)) / dS ^ 2
    End If
End Function


'//  Bachelier Modified  formula
Public Function BachelierModified(CallPutFlag As String, S As Double, x _
                , T As Double, r As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    d1 = (S - x) / (v * Sqr(T))
  

    If CallPutFlag = "c" Then
        BachelierModified = S * CND(d1) - x * Exp(-r * T) * CND(d1) + v * Sqr(T) * ND(d1)
    ElseIf CallPutFlag = "p" Then
        BachelierModified = x * Exp(-r * T) * CND(-d1) - S * CND(-d1) + v * Sqr(T) * ND(d1)
    End If
End Function


Public Function EBachelierModified(OutPutFlag As String, CallPutFlag As String, S As Double, x As Double, T As Double, _
                 r As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EBachelierModified = BachelierModified(CallPutFlag, S, x, T, r, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EBachelierModified = (BachelierModified(CallPutFlag, S + dS, x, T, r, v) - BachelierModified(CallPutFlag, S - dS, x, T, r, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EBachelierModified = (BachelierModified(CallPutFlag, S + dS, x, T, r, v) - BachelierModified(CallPutFlag, S - dS, x, T, r, v)) / (2 * dS) * S / BachelierModified(CallPutFlag, S, x, T, r, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EBachelierModified = (BachelierModified(CallPutFlag, S + dS, x, T, r, v) - 2 * BachelierModified(CallPutFlag, S, x, T, r, v) + BachelierModified(CallPutFlag, S - dS, x, T, r, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EBachelierModified = (BachelierModified(CallPutFlag, S + dS, x, T, r, v + 0.01) - 2 * BachelierModified(CallPutFlag, S, x, T, r, v + 0.01) + BachelierModified(CallPutFlag, S - dS, x, T, r, v + 0.01) _
      - BachelierModified(CallPutFlag, S + dS, x, T, r, v - 0.01) + 2 * BachelierModified(CallPutFlag, S, x, T, r, v - 0.01) - BachelierModified(CallPutFlag, S - dS, x, T, r, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EBachelierModified = S / 100 * (BachelierModified(CallPutFlag, S + dS, x, T, r, v) - 2 * BachelierModified(CallPutFlag, S, x, T, r, v) + BachelierModified(CallPutFlag, S - dS, x, T, r, v)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        EBachelierModified = (BachelierModified(CallPutFlag, S, x, T + 1 / 365, r, v) - 2 * BachelierModified(CallPutFlag, S, x, T, r, v) + BachelierModified(CallPutFlag, S, x, T - 1 / 365, r, v)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EBachelierModified = 1 / (4 * dS * 0.01) * (BachelierModified(CallPutFlag, S + dS, x, T, r, v + 0.01) - BachelierModified(CallPutFlag, S + dS, x, T, r, v - 0.01) _
        - BachelierModified(CallPutFlag, S - dS, x, T, r, v + 0.01) + BachelierModified(CallPutFlag, S - dS, x, T, r, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EBachelierModified = (BachelierModified(CallPutFlag, S, x, T, r, v + 0.01) - BachelierModified(CallPutFlag, S, x, T, r, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EBachelierModified = (BachelierModified(CallPutFlag, S, x, T, r, v + 0.01) - 2 * BachelierModified(CallPutFlag, S, x, T, r, v) + BachelierModified(CallPutFlag, S, x, T, r, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EBachelierModified = v / 0.1 * (BachelierModified(CallPutFlag, S, x, T, r, v + 0.01) - BachelierModified(CallPutFlag, S, x, T, r, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EBachelierModified = (BachelierModified(CallPutFlag, S, x, T, r, v + 0.01) - 2 * BachelierModified(CallPutFlag, S, x, T, r, v) + BachelierModified(CallPutFlag, S, x, T, r, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EBachelierModified = BachelierModified(CallPutFlag, S, x, 0.00001, r, v) - BachelierModified(CallPutFlag, S, x, T, r, v)
        Else
                EBachelierModified = BachelierModified(CallPutFlag, S, x, T - 1 / 365, r, v) - BachelierModified(CallPutFlag, S, x, T, r, v)
        End If
      ElseIf OutPutFlag = "r" Then 'Rho
         EBachelierModified = (BachelierModified(CallPutFlag, S, x, T, r + 0.01, v) - BachelierModified(CallPutFlag, S, x, T, r - 0.01, v)) / 2
   
      ElseIf OutPutFlag = "s" Then 'Speed
        EBachelierModified = 1 / dS ^ 3 * (BachelierModified(CallPutFlag, S + 2 * dS, x, T, r, v) - 3 * BachelierModified(CallPutFlag, S + dS, x, T, r, v) _
                                + 3 * BachelierModified(CallPutFlag, S, x, T, r, v) - BachelierModified(CallPutFlag, S - dS, x, T, r, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EBachelierModified = (BachelierModified(CallPutFlag, S, x + dS, T, r, v) - BachelierModified(CallPutFlag, S, x - dS, T, r, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EBachelierModified = (BachelierModified(CallPutFlag, S, x + dS, T, r, v) - 2 * BachelierModified(CallPutFlag, S, x, T, r, v) + BachelierModified(CallPutFlag, S, x - dS, T, r, v)) / dS ^ 2
    End If
End Function



'//  Sprenkle 1964 formula
Public Function Sprenkle(CallPutFlag As String, S As Double, x _
                , T As Double, rho As Double, k As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    d1 = (Log(S / x) + (rho + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
  
    If CallPutFlag = "c" Then
        Sprenkle = S * Exp(rho * T) * CND(d1) - (1 - k) * x * CND(d2)
    ElseIf CallPutFlag = "p" Then
        Sprenkle = (1 - k) * x * CND(-d2) - S * Exp(rho * T) * CND(-d1)
    End If
End Function



Public Function ESprenkle(OutPutFlag As String, CallPutFlag As String, S As Double, x As Double, T As Double, _
                rho As Double, k As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        ESprenkle = Sprenkle(CallPutFlag, S, x, T, rho, k, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         ESprenkle = (Sprenkle(CallPutFlag, S + dS, x, T, rho, k, v) - Sprenkle(CallPutFlag, S - dS, x, T, rho, k, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         ESprenkle = (Sprenkle(CallPutFlag, S + dS, x, T, rho, k, v) - Sprenkle(CallPutFlag, S - dS, x, T, rho, k, v)) / (2 * dS) * S / Sprenkle(CallPutFlag, S, x, T, rho, k, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ESprenkle = (Sprenkle(CallPutFlag, S + dS, x, T, rho, k, v) - 2 * Sprenkle(CallPutFlag, S, x, T, rho, k, v) + Sprenkle(CallPutFlag, S - dS, x, T, rho, k, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ESprenkle = (Sprenkle(CallPutFlag, S + dS, x, T, rho, k, v + 0.01) - 2 * Sprenkle(CallPutFlag, S, x, T, rho, k, v + 0.01) + Sprenkle(CallPutFlag, S - dS, x, T, rho, k, v + 0.01) _
      - Sprenkle(CallPutFlag, S + dS, x, T, rho, k, v - 0.01) + 2 * Sprenkle(CallPutFlag, S, x, T, rho, k, v - 0.01) - Sprenkle(CallPutFlag, S - dS, x, T, rho, k, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        ESprenkle = S / 100 * (Sprenkle(CallPutFlag, S + dS, x, T, rho, k, v) - 2 * Sprenkle(CallPutFlag, S, x, T, rho, k, v) + Sprenkle(CallPutFlag, S - dS, x, T, rho, k, v)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        ESprenkle = (Sprenkle(CallPutFlag, S, x, T + 1 / 365, rho, k, v) - 2 * Sprenkle(CallPutFlag, S, x, T, rho, k, v) + Sprenkle(CallPutFlag, S, x, T - 1 / 365, rho, k, v)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ESprenkle = 1 / (4 * dS * 0.01) * (Sprenkle(CallPutFlag, S + dS, x, T, rho, k, v + 0.01) - Sprenkle(CallPutFlag, S + dS, x, T, rho, k, v - 0.01) _
        - Sprenkle(CallPutFlag, S - dS, x, T, rho, k, v + 0.01) + Sprenkle(CallPutFlag, S - dS, x, T, rho, k, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ESprenkle = (Sprenkle(CallPutFlag, S, x, T, rho, k, v + 0.01) - Sprenkle(CallPutFlag, S, x, T, rho, k, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        ESprenkle = (Sprenkle(CallPutFlag, S, x, T, rho, k, v + 0.01) - 2 * Sprenkle(CallPutFlag, S, x, T, rho, k, v) + Sprenkle(CallPutFlag, S, x, T, rho, k, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ESprenkle = v / 0.1 * (Sprenkle(CallPutFlag, S, x, T, rho, k, v + 0.01) - Sprenkle(CallPutFlag, S, x, T, rho, k, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ESprenkle = (Sprenkle(CallPutFlag, S, x, T, rho, k, v + 0.01) - 2 * Sprenkle(CallPutFlag, S, x, T, rho, k, v) + Sprenkle(CallPutFlag, S, x, T, rho, k, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ESprenkle = Sprenkle(CallPutFlag, S, x, 0.00001, rho, k, v) - Sprenkle(CallPutFlag, S, x, T, rho, k, v)
        Else
                ESprenkle = Sprenkle(CallPutFlag, S, x, T - 1 / 365, rho, k, v) - Sprenkle(CallPutFlag, S, x, T, rho, k, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         ESprenkle = (Sprenkle(CallPutFlag, S, x, T, rho + 0.01, k, v) - Sprenkle(CallPutFlag, S, x, T, rho - 0.01, k, v)) / (2)
        ElseIf OutPutFlag = "k" Then 'k
        ESprenkle = (Sprenkle(CallPutFlag, S, x, T, rho, k + 0.01, v) - Sprenkle(CallPutFlag, S, x, T, rho, k - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ESprenkle = 1 / dS ^ 3 * (Sprenkle(CallPutFlag, S + 2 * dS, x, T, rho, k, v) - 3 * Sprenkle(CallPutFlag, S + dS, x, T, rho, k, v) _
                                + 3 * Sprenkle(CallPutFlag, S, x, T, rho, k, v) - Sprenkle(CallPutFlag, S - dS, x, T, rho, k, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ESprenkle = (Sprenkle(CallPutFlag, S, x + dS, T, rho, k, v) - Sprenkle(CallPutFlag, S, x - dS, T, rho, k, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        ESprenkle = (Sprenkle(CallPutFlag, S, x + dS, T, rho, k, v) - 2 * Sprenkle(CallPutFlag, S, x, T, rho, k, v) + Sprenkle(CallPutFlag, S, x - dS, T, rho, k, v)) / dS ^ 2
    End If
End Function


