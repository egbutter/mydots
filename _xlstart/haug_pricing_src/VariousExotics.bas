Attribute VB_Name = "VariousExotics"
Option Explicit

' Programmer Espen Gaarder Haug
' Copyright 2006, Espen Gaarder Haug

 Public Function MirrorOption(LongShortFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, v As Double) As Double
    Dim F As Double, d1 As Double, d2 As Double
    
    Dim y As Integer, z As Integer
    
    z = 1
    If CallPutFlag = "p" Then
        z = -1
    End If
    
    y = 1
    If LongShortFlag = "s" And z = 1 Then
        y = -1
    ElseIf LongShortFlag = "l" And z = -1 Then
        y = -1
    ElseIf LongShortFlag = "s" And z = -1 Then
       y = 1
    End If
    
    F = S * Exp((v ^ 2 / 2 + y * Abs(b - v ^ 2 / 2)) * T)
    

    d1 = (Log(F / X) + T * v ^ 2 / 2) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    d2 = (Log(F / X) - T * v ^ 2 / 2) / (v * Sqr(T))
    
    If CallPutFlag = "c" Then
        MirrorOption = y * Exp(-r * T) * (F * CND(d1) - X * CND(d2))
    ElseIf CallPutFlag = "p" Then
        MirrorOption = -y * Exp(-r * T) * (X * CND(-d2) - F * CND(-d1))
    End If
 
 End Function

  Public Function VPO(S As Double, X As Double, L As Double, U As Double, D As Double, T As Double, r As Double, b As Double, v As Double) As Double
 
    Dim d1 As Double, d2 As Double, d3 As Double, d4 As Double
    Dim d5 As Double, d6 As Double
    Dim NMin As Double, NMax As Double
    
    NMin = X / (U * (1 - D))
    NMax = X / (L * (1 - D))
    
    d1 = (Log(S / U) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    d3 = (Log(S / L) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d4 = d3 - v * Sqr(T)
    d5 = (Log(S / (L * (1 - D))) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d6 = d5 - v * Sqr(T)

    VPO = X * D / (1 - D) * Exp(-r * T) + NMin * (S * Exp((b - r) * T) * CND(d1) - U * Exp(-r * T) * CND(d2)) _
        - NMax * (L * Exp(-r * T) * CND(-d4) - S * Exp((b - r) * T) * CND(-d3)) _
        + NMax * (L * (1 - D) * Exp(-r * T) * CND(-d6) - S * Exp((b - r) * T) * CND(-d5))
 
 End Function

Public Function FadeInOption(CallPutFlag As String, S As Double, X As Double, L As Double, H As Double, T As Double, r As Double, b As Double, v As Double, n As Double) As Double

    Dim d1 As Double, d2 As Double, d3 As Double, d4 As Double, d5 As Double, d6 As Double
    Dim rho As Double, dt As Double, sum As Double, t1 As Double
    Dim i As Integer
    
    dt = T / n
    d1 = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    
    sum = 0
    For i = 1 To n
        t1 = i * dt
        rho = Sqr(t1) / Sqr(T)
    
        d3 = (Log(S / L) + (b + v ^ 2 / 2) * t1) / (v * Sqr(t1))
        d4 = d3 - v * Sqr(t1)
        d5 = (Log(S / H) + (b + v ^ 2 / 2) * t1) / (v * Sqr(t1))
        d6 = d5 - v * Sqr(t1)
      
        If CallPutFlag = "c" Then
            sum = sum + S * Exp((b - r) * T) * (CBND(-d5, d1, -rho) - CBND(-d3, d1, -rho)) _
                   - X * Exp(-r * T) * (CBND(-d6, d2, -rho) - CBND(-d4, d2, -rho))
        ElseIf CallPutFlag = "p" Then
            sum = sum + X * Exp(-r * T) * (CBND(-d6, -d2, rho) - CBND(-d4, -d2, rho)) _
                - S * Exp((b - r) * T) * (CBND(-d5, -d1, rho) - CBND(-d3, -d1, rho))
        End If
    Next
    
    FadeInOption = 1 / n * sum
    
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


Public Function EExecutive(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, lambda As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        EExecutive = Executive(CallPutFlag, S, X, T, r, b, v, lambda)
    ElseIf OutPutFlag = "d" Then 'Delta
         EExecutive = (Executive(CallPutFlag, S + dS, X, T, r, b, v, lambda) - Executive(CallPutFlag, S - dS, X, T, r, b, v, lambda)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EExecutive = (Executive(CallPutFlag, S + dS, X, T, r, b, v, lambda) - Executive(CallPutFlag, S - dS, X, T, r, b, v, lambda)) / (2 * dS) * S / Executive(CallPutFlag, S, X, T, r, b, v, lambda)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EExecutive = (Executive(CallPutFlag, S + dS, X, T, r, b, v, lambda) - 2 * Executive(CallPutFlag, S, X, T, r, b, v, lambda) + Executive(CallPutFlag, S - dS, X, T, r, b, v, lambda)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EExecutive = (Executive(CallPutFlag, S + dS, X, T, r, b, v + 0.01, lambda) - 2 * Executive(CallPutFlag, S, X, T, r, b, v + 0.01, lambda) + Executive(CallPutFlag, S - dS, X, T, r, b, v + 0.01, lambda) _
      - Executive(CallPutFlag, S + dS, X, T, r, b, v - 0.01, lambda) + 2 * Executive(CallPutFlag, S, X, T, r, b, v - 0.01, lambda) - Executive(CallPutFlag, S - dS, X, T, r, b, v - 0.01, lambda)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EExecutive = S / 100 * (Executive(CallPutFlag, S + dS, X, T, r, b, v, lambda) - 2 * Executive(CallPutFlag, S, X, T, r, b, v, lambda) + Executive(CallPutFlag, S - dS, X, T, r, b, v, lambda)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EExecutive = 1 / (4 * dS * 0.01) * (Executive(CallPutFlag, S + dS, X, T, r, b, v + 0.01, lambda) - Executive(CallPutFlag, S + dS, X, T, r, b, v - 0.01, lambda) _
        - Executive(CallPutFlag, S - dS, X, T, r, b, v + 0.01, lambda) + Executive(CallPutFlag, S - dS, X, T, r, b, v - 0.01, lambda)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EExecutive = (Executive(CallPutFlag, S, X, T, r, b, v + 0.01, lambda) - Executive(CallPutFlag, S, X, T, r, b, v - 0.01, lambda)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EExecutive = (Executive(CallPutFlag, S, X, T, r, b, v + 0.01, lambda) - 2 * Executive(CallPutFlag, S, X, T, r, b, v, lambda) + Executive(CallPutFlag, S, X, T, r, b, v - 0.01, lambda)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EExecutive = v / 0.1 * (Executive(CallPutFlag, S, X, T, r, b, v + 0.01, lambda) - Executive(CallPutFlag, S, X, T, r, b, v - 0.01, lambda)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EExecutive = (Executive(CallPutFlag, S, X, T, r, b, v + 0.01, lambda) - 2 * Executive(CallPutFlag, S, X, T, r, b, v, lambda) + Executive(CallPutFlag, S, X, T, r, b, v - 0.01, lambda))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EExecutive = Executive(CallPutFlag, S, X, 0.00001, r, b, v, lambda) - Executive(CallPutFlag, S, X, T, r, b, v, lambda)
        Else
                EExecutive = Executive(CallPutFlag, S, X, T - 1 / 365, r, b, v, lambda) - Executive(CallPutFlag, S, X, T, r, b, v, lambda)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EExecutive = (Executive(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v, lambda) - Executive(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v, lambda)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EExecutive = (Executive(CallPutFlag, S, X, T, r + 0.01, b, v, lambda) - Executive(CallPutFlag, S, X, T, r - 0.01, b, v, lambda)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EExecutive = (Executive(CallPutFlag, S, X, T, r, b - 0.01, v, lambda) - Executive(CallPutFlag, S, X, T, r, b + 0.01, v, lambda)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EExecutive = (Executive(CallPutFlag, S, X, T, r, b + 0.01, v, lambda) - Executive(CallPutFlag, S, X, T, r, b - 0.01, v, lambda)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EExecutive = 1 / dS ^ 3 * (Executive(CallPutFlag, S + 2 * dS, X, T, r, b, v, lambda) - 3 * Executive(CallPutFlag, S + dS, X, T, r, b, v, lambda) _
                                + 3 * Executive(CallPutFlag, S, X, T, r, b, v, lambda) - Executive(CallPutFlag, S - dS, X, T, r, b, v, lambda))
     ElseIf OutPutFlag = "dx" Then 'StrikeDelta
         EExecutive = (Executive(CallPutFlag, S, X + dS, T, r, b, v, lambda) - Executive(CallPutFlag, S, X - dS, T, r, b, v, lambda)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike gamma
        EExecutive = (Executive(CallPutFlag, S, X + dS, T, r, b, v, lambda) - 2 * Executive(CallPutFlag, S, X, T, r, b, v, lambda) + Executive(CallPutFlag, S, X - dS, T, r, b, v, lambda)) / dS ^ 2
    ElseIf OutPutFlag = "j" Then 'Sensitivity to jump
         EExecutive = (Executive(CallPutFlag, S, X + dS, T, r, b, v, lambda + 0.01) - Executive(CallPutFlag, S, X - dS, T, r, b, v, lambda - 0.01)) / (2)
    
    End If
End Function

'// Executive stock options
Public Function Executive(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, v As Double, lambda As Double) As Double
    
    Executive = Exp(-lambda * T) * GBlackScholes(CallPutFlag, S, X, T, r, b, v)

End Function

'// Forward start options
Public Function ForwardStartOption(CallPutFlag As String, S As Double, Alpha As Double, t1 As Double, _
                T As Double, r As Double, b As Double, v As Double) As Double

    ForwardStartOption = S * Exp((b - r) * t1) * GBlackScholes(CallPutFlag, 1, Alpha, T - t1, r, b, v)

End Function


Public Function MoneynessOption(Moneyness As Double, T As Double, r As Double, v As Double) As Double
 
        Dim d1 As Double, d2 As Double
        
        d1 = (-Log(Moneyness) + v ^ 2 / 2 * T) / (v * Sqr(T))
        d2 = d1 - v * Sqr(T)
        
        
        MoneynessOption = Exp(-r * T) * (CND(d1) - Moneyness * CND(d2))
     
        
 End Function
 
 
 Public Function EMoneynessOption(OutPutFlag As String, S As Double, T As Double, _
                r As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EMoneynessOption = MoneynessOption(S, T, r, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EMoneynessOption = (MoneynessOption(S + dS, T, r, v) - MoneynessOption(S - dS, T, r, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EMoneynessOption = (MoneynessOption(S + dS, T, r, v) - MoneynessOption(S - dS, T, r, v)) / (2 * dS) * S / MoneynessOption(S, T, r, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EMoneynessOption = (MoneynessOption(S + dS, T, r, v) - 2 * MoneynessOption(S, T, r, v) + MoneynessOption(S - dS, T, r, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EMoneynessOption = (MoneynessOption(S + dS, T, r, v + 0.01) - 2 * MoneynessOption(S, T, r, v + 0.01) + MoneynessOption(S - dS, T, r, v + 0.01) _
      - MoneynessOption(S + dS, T, r, v - 0.01) + 2 * MoneynessOption(S, T, r, v - 0.01) - MoneynessOption(S - dS, T, r, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "gp" Then 'GammaP
        EMoneynessOption = S / 100 * (MoneynessOption(S + dS, T, r, v) - 2 * MoneynessOption(S, T, r, v) + MoneynessOption(S - dS, T, r, v)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        EMoneynessOption = (MoneynessOption(S, T + 1 / 365, r, v) - 2 * MoneynessOption(S, T, r, v) + MoneynessOption(S, T - 1 / 365, r, v)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EMoneynessOption = 1 / (4 * dS * 0.01) * (MoneynessOption(S + dS, T, r, v + 0.01) - MoneynessOption(S + dS, T, r, v - 0.01) _
        - MoneynessOption(S - dS, T, r, v + 0.01) + MoneynessOption(S - dS, T, r, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EMoneynessOption = (MoneynessOption(S, T, r, v + 0.01) - MoneynessOption(S, T, r, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EMoneynessOption = (MoneynessOption(S, T, r, v + 0.01) - 2 * MoneynessOption(S, T, r, v) + MoneynessOption(S, T, r, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EMoneynessOption = v / 0.1 * (MoneynessOption(S, T, r, v + 0.01) - MoneynessOption(S, T, r, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EMoneynessOption = (MoneynessOption(S, T, r, v + 0.01) - 2 * MoneynessOption(S, T, r, v) + MoneynessOption(S, T, r, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EMoneynessOption = MoneynessOption(S, 0.00001, r, v) - MoneynessOption(S, T, r, v)
        Else
                EMoneynessOption = MoneynessOption(S, T - 1 / 365, r, v) - MoneynessOption(S, T, r, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EMoneynessOption = (MoneynessOption(S, T, r + 0.01, v) - MoneynessOption(S, T, r - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EMoneynessOption = (MoneynessOption(S, T, r + 0.01, v) - MoneynessOption(S, T, r - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EMoneynessOption = 1 / dS ^ 3 * (MoneynessOption(S + 2 * dS, T, r, v) - 3 * MoneynessOption(S + dS, T, r, v) _
                                + 3 * MoneynessOption(S, T, r, v) - MoneynessOption(S - dS, T, r, v))
    End If
End Function
