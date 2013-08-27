Attribute VB_Name = "AsianOptions"
Option Explicit


' Implementation By Espen Gaarder Haug
' Copyright Espen Gaarder Haug 2006



Public Function EAsianCurranApprox(OutPutFlag As String, CallPutFlag As String, S As Double, SA As Double, X As Double, t1 As Double, T As Double, n As Long, m As Long, _
                r As Double, b As Double, v As Double, Optional dS)
            
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        EAsianCurranApprox = AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v) - AsianCurranApprox(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v) - AsianCurranApprox(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v)) / (2 * dS) * S / AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v) - 2 * AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) + AsianCurranApprox(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v + 0.01) - 2 * AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v + 0.01) + AsianCurranApprox(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v + 0.01) _
      - AsianCurranApprox(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v - 0.01) + 2 * AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v - 0.01) - AsianCurranApprox(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "gp" Then 'GammaP
        EAsianCurranApprox = S / 100 * (AsianCurranApprox(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v) - 2 * AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) + AsianCurranApprox(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EAsianCurranApprox = 1 / (4 * dS * 0.01) * (AsianCurranApprox(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v + 0.01) - AsianCurranApprox(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v - 0.01) _
        - AsianCurranApprox(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v + 0.01) + AsianCurranApprox(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v + 0.01) - AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v + 0.01) - 2 * AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) + AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EAsianCurranApprox = v / 0.1 * (AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v + 0.01) - AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v + 0.01) - 2 * AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) + AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If t1 > 1 / 365 And T > 1 / 365 Then
                       EAsianCurranApprox = AsianCurranApprox(CallPutFlag, S, SA, X, t1 - 1 / 365, T - 1 / 365, n, m, r, b, v) - AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r + 0.01, b + 0.01, v) - AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r + 0.01, b, v) - AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b - 0.01, v) - AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b + 0.01, v) - AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EAsianCurranApprox = 1 / dS ^ 3 * (AsianCurranApprox(CallPutFlag, S + 2 * dS, SA, X, t1, T, n, m, r, b, v) - 3 * AsianCurranApprox(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v) _
                                + 3 * AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) - AsianCurranApprox(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S, SA, X + dS, t1, T, n, m, r, b, v) - AsianCurranApprox(CallPutFlag, S, SA, X - dS, t1, T, n, m, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EAsianCurranApprox = (AsianCurranApprox(CallPutFlag, S, SA, X + dS, t1, T, n, m, r, b, v) - 2 * AsianCurranApprox(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) + AsianCurranApprox(CallPutFlag, S, SA, X - dS, t1, T, n, m, r, b, v)) / dS ^ 2
    End If
End Function

Public Function AsianCurranApprox(CallPutFlag As String, S As Double, SA As Double, X As Double, t1 As Double, _
     T As Double, n As Long, m As Long, r As Double, b As Double, v As Double) As Double

    Dim dt As Double, my As Double, myi As Double
    Dim vxi As Double, vi As Double, vx As Double
    Dim Km As Double, sum1 As Double, sum2 As Double
    Dim ti As Double, EA As Double
    Dim z As Integer, i As Long

    
    z = 1
    If CallPutFlag = "p" Then
        z = -1
    End If
        

    dt = (T - t1) / (n - 1)
    
     If b = 0 Then
        EA = S
     Else
        EA = S / n * Exp(b * t1) * (1 - Exp(b * dt * n)) / (1 - Exp(b * dt))
     End If

    If m > 0 Then
        If SA > n / m * X Then
        '//Exercise is certain for call, put must be out-of-the-money:
            If CallPutFlag = "p" Then
                AsianCurranApprox = 0
            ElseIf CallPutFlag = "c" Then
                SA = SA * m / n + EA * (n - m) / n
                AsianCurranApprox = (SA - X) * Exp(-r * T)
            End If
            Exit Function
         End If
   End If

    If m = n - 1 Then
    ' // Only one fix left use Black-Scholes weighted with time
            X = n * X - (n - 1) * SA
             AsianCurranApprox = GBlackScholes(CallPutFlag, S, X, T, r, b, v) * 1 / n
              Exit Function
    End If

    If m > 0 Then
        X = n / (n - m) * X - m / (n - m) * SA
    End If


    vx = v * Sqr(t1 + dt * (n - 1) * (2 * n - 1) / (6 * n))
    my = Log(S) + (b - v * v * 0.5) * (t1 + (n - 1) * dt / 2)

    sum1 = 0
    For i = 1 To n Step 1
    
        ti = dt * i + t1 - dt
        vi = v * Sqr(t1 + (i - 1) * dt)
        vxi = v * v * (t1 + dt * ((i - 1) - i * (i - 1) / (2 * n)))
        myi = Log(S) + (b - v * v * 0.5) * ti
        sum1 = sum1 + Exp(myi + vxi / (vx * vx) * _
        (Log(X) - my) + (vi * vi - vxi * vxi / (vx * vx)) * 0.5)
    Next
    Km = 2 * X - 1 / n * sum1
    sum2 = 0


    For i = 1 To n Step 1
    
        ti = dt * i + t1 - dt
        vi = v * Sqr(t1 + (i - 1) * dt)
        vxi = v * v * (t1 + dt * ((i - 1) - i * (i - 1) / (2 * n)))
        myi = Log(S) + (b - v * v * 0.5) * ti
        sum2 = sum2 + Exp(myi + vi * vi * 0.5) * CND(z * ((my - Log(Km)) / vx + vxi / vx))
    
    Next

    AsianCurranApprox = Exp(-r * T) * z * (1 / n * sum2 - X * CND(z * (my - Log(Km)) / vx)) * (n - m) / n
End Function


Public Function DiscreteAsianHHM(CallPutFlag As String, S As Double, SA As Double, X As Double, _
              t1 As Double, T As Double, n As Double, m As Double, r As Double, b As Double, v As Double) As Double

'This is a modified version of the Levy formula, this is the formula published in "Asian Pyramid Power" By
' Haug, Haug and Margrabe
            
    Dim d1 As Double, d2 As Double, h As Double, EA As Double, EA2 As Double
    Dim vA As Double, OptionValue As Double
    
  

   h = (T - t1) / (n - 1)

   If b = 0 Then
        EA = S
   Else
        EA = S / n * Exp(b * t1) * (1 - Exp(b * h * n)) / (1 - Exp(b * h))
   End If
   
   If m > 0 Then
        If SA > n / m * X Then   '// Exercise is certain for call, put must be out-of-the-money
        
            If CallPutFlag = "p" Then
                DiscreteAsianHHM = 0
            ElseIf CallPutFlag = "c" Then
                SA = SA * m / n + EA * (n - m) / n
                DiscreteAsianHHM = (SA - X) * Exp(-r * T)
            End If
            Exit Function

       End If
    End If

 If m = n - 1 Then ' // Only one fix left use Black-Scholes weighted with time
   
         X = n * X - (n - 1) * SA
         DiscreteAsianHHM = GBlackScholes(CallPutFlag, S, X, T, r, b, v) * 1 / n
         Exit Function
   End If

    If b = 0 Then
    
         EA2 = S * S * Exp(v * v * t1) / (n * n) _
            * ((1 - Exp(v * v * h * n)) / (1 - Exp(v * v * h)) _
           + 2 / (1 - Exp(v * v * h)) * (n - (1 - Exp(v * v * h * n)) / (1 - Exp(v * v * h))))
    Else
    
         EA2 = S * S * Exp((2 * b + v * v) * t1) / (n * n) _
            * ((1 - Exp((2 * b + v * v) * h * n)) / (1 - Exp((2 * b + v * v) * h)) _
            + 2 / (1 - Exp((b + v * v) * h)) * ((1 - Exp(b * h * n)) / (1 - Exp(b * h)) _
            - (1 - Exp((2 * b + v * v) * h * n)) / _
            (1 - Exp((2 * b + v * v) * h))))
    End If

    vA = Sqr((Log(EA2) - 2 * Log(EA)) / T)

    OptionValue = 0
    
    If m > 0 Then
        X = n / (n - m) * X - m / (n - m) * SA
    End If
    
    d1 = (Log(EA / X) + vA ^ 2 / 2 * T) / (vA * Sqr(T))
    d2 = d1 - vA * Sqr(T)

    If CallPutFlag = "c" Then
        OptionValue = Exp(-r * T) * (EA * CND(d1) - X * CND(d2))
    ElseIf (CallPutFlag = "p") Then
        OptionValue = Exp(-r * T) * (X * CND(-d2) - EA * CND(-d1))
    End If

    DiscreteAsianHHM = OptionValue * (n - m) / n

End Function


Public Function EDiscreteAsianHHM(OutPutFlag As String, CallPutFlag As String, S As Double, SA As Double, X As Double, t1 As Double, T As Double, n As Double, m As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EDiscreteAsianHHM = DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v) - DiscreteAsianHHM(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v) - DiscreteAsianHHM(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v)) / (2 * dS) * S / DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v) - 2 * DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) + DiscreteAsianHHM(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v + 0.01) - 2 * DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v + 0.01) + DiscreteAsianHHM(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v + 0.01) _
      - DiscreteAsianHHM(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v - 0.01) + 2 * DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v - 0.01) - DiscreteAsianHHM(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "gp" Then 'GammaP
        EDiscreteAsianHHM = S / 100 * (DiscreteAsianHHM(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v) - 2 * DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) + DiscreteAsianHHM(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EDiscreteAsianHHM = 1 / (4 * dS * 0.01) * (DiscreteAsianHHM(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v + 0.01) - DiscreteAsianHHM(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v - 0.01) _
        - DiscreteAsianHHM(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v + 0.01) + DiscreteAsianHHM(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v + 0.01) - DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v + 0.01) - 2 * DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) + DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EDiscreteAsianHHM = v / 0.1 * (DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v + 0.01) - DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v + 0.01) - 2 * DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) + DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If t1 > 1 / 365 And T > 1 / 365 Then
                       EDiscreteAsianHHM = DiscreteAsianHHM(CallPutFlag, S, SA, X, t1 - 1 / 365, T - 1 / 365, n, m, r, b, v) - DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r + 0.01, b + 0.01, v) - DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r + 0.01, b, v) - DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b - 0.01, v) - DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b + 0.01, v) - DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EDiscreteAsianHHM = 1 / dS ^ 3 * (DiscreteAsianHHM(CallPutFlag, S + 2 * dS, SA, X, t1, T, n, m, r, b, v) - 3 * DiscreteAsianHHM(CallPutFlag, S + dS, SA, X, t1, T, n, m, r, b, v) _
                                + 3 * DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) - DiscreteAsianHHM(CallPutFlag, S - dS, SA, X, t1, T, n, m, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S, SA, X + dS, t1, T, n, m, r, b, v) - DiscreteAsianHHM(CallPutFlag, S, SA, X - dS, t1, T, n, m, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EDiscreteAsianHHM = (DiscreteAsianHHM(CallPutFlag, S, SA, X + dS, t1, T, n, m, r, b, v) - 2 * DiscreteAsianHHM(CallPutFlag, S, SA, X, t1, T, n, m, r, b, v) + DiscreteAsianHHM(CallPutFlag, S, SA, X - dS, t1, T, n, m, r, b, v)) / dS ^ 2
    End If
End Function



' // Arithmetic average rate option
 Public Function TurnbullWakemanAsian(CallPutFlag As String, S As Double, SA As Double, X As Double, _
              T As Double, T2 As Double, r As Double, b As Double, v As Double) As Double

    ' // CallPutFlag = "c" for call and "p" for put option
    ' // S = Asset price
    ' // SA= Realized average so far
    ' // X = Strike price
    ' // t1 = Time to start of average period in years
    ' // T =  Time to maturity in years of option  T
    ' // T2 = Original time in average period in years, constant over life of option
    ' // r = risk-free rate
    ' // b = cost of carry underlying asset can be positive and negative
    ' // v = annualized volatility of asset price


    Dim M1 As Double, M2 As Double, tau As Double, t1 As Double
    Dim bA As Double, vA As Double
    
    '//tau: reminding time of average perios

    t1 = Max(0, T - T2)
    tau = T2 - T
   
    If b = 0 Then
    
        M1 = 1
       
    Else
    
         M1 = (Exp(b * T) - Exp(b * t1)) / (b * (T - t1))
    End If

    '//Take into account when option wil  be exercised
    If tau > 0 Then
    
        If T2 / T * X - tau / T * SA < 0 Then
    
            If CallPutFlag = "c" Then
                 SA = SA * (T2 - T) / T2 + S * M1 * T / T2 ' //Expected average at maturity
                TurnbullWakemanAsian = Max(0, SA - X) * Exp(-r * T)
            Else
                TurnbullWakemanAsian = 0
            End If
            Exit Function
             
         End If
    End If

   If b = 0 Then  '//   Extended to hold for options on futures 16 May 1999 Espen Haug
   
       M2 = 2 * Exp(v * v * T) / (v ^ 4 * (T - t1) ^ 2) _
                - 2 * Exp(v * v * t1) * (1 + v * v * (T - t1)) / (v ^ 4 * (T - t1) ^ 2)
   
   Else
   
        M2 = 2 * Exp((2 * b + v * v) * T) / ((b + v * v) * (2 * b + v * v) * (T - t1) ^ 2) _
              + 2 * Exp((2 * b + v * v) * t1) / (b * (T - t1) ^ 2) * (1 / (2 * b + v * v) - Exp(b * (T - t1)) / (b + v * v))
    End If
        bA = Log(M1) / T
        vA = Sqr(Log(M2) / T - 2 * bA)

    If tau > 0 Then
        X = T2 / T * X - tau / T * SA
      
         TurnbullWakemanAsian = GBlackScholes(CallPutFlag, S, X, T, r, bA, vA) * T / T2
    Else
        TurnbullWakemanAsian = GBlackScholes(CallPutFlag, S, X, T, r, bA, vA)
    End If
End Function



Public Function ETurnbullWakemanAsian(OutPutFlag As String, CallPutFlag As String, S As Double, SA As Double, X As Double, T As Double, T2 As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        ETurnbullWakemanAsian = TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S + dS, SA, X, T, T2, r, b, v) - TurnbullWakemanAsian(CallPutFlag, S - dS, SA, X, T, T2, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S + dS, SA, X, T, T2, r, b, v) - TurnbullWakemanAsian(CallPutFlag, S - dS, SA, X, T, T2, r, b, v)) / (2 * dS) * S / TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S + dS, SA, X, T, T2, r, b, v) - 2 * TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v) + TurnbullWakemanAsian(CallPutFlag, S - dS, SA, X, T, T2, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S + dS, SA, X, T, T2, r, b, v + 0.01) - 2 * TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v + 0.01) + TurnbullWakemanAsian(CallPutFlag, S - dS, SA, X, T, T2, r, b, v + 0.01) _
      - TurnbullWakemanAsian(CallPutFlag, S + dS, SA, X, T, T2, r, b, v - 0.01) + 2 * TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v - 0.01) - TurnbullWakemanAsian(CallPutFlag, S - dS, SA, X, T, T2, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "gp" Then 'GammaP
        ETurnbullWakemanAsian = S / 100 * (TurnbullWakemanAsian(CallPutFlag, S + dS, SA, X, T, T2, r, b, v) - 2 * TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v) + TurnbullWakemanAsian(CallPutFlag, S - dS, SA, X, T, T2, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ETurnbullWakemanAsian = 1 / (4 * dS * 0.01) * (TurnbullWakemanAsian(CallPutFlag, S + dS, SA, X, T, T2, r, b, v + 0.01) - TurnbullWakemanAsian(CallPutFlag, S + dS, SA, X, T, T2, r, b, v - 0.01) _
        - TurnbullWakemanAsian(CallPutFlag, S - dS, SA, X, T, T2, r, b, v + 0.01) + TurnbullWakemanAsian(CallPutFlag, S - dS, SA, X, T, T2, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v + 0.01) - TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v + 0.01) - 2 * TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v) + TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ETurnbullWakemanAsian = v / 0.1 * (TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v + 0.01) - TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v + 0.01) - 2 * TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v) + TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                ETurnbullWakemanAsian = TurnbullWakemanAsian(CallPutFlag, S, SA, X, 0.00001, T2, r, b, v) - TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v)
        Else
                       ETurnbullWakemanAsian = TurnbullWakemanAsian(CallPutFlag, S, SA, X, T - 1 / 365, T2, r, b, v) - TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r + 0.01, b + 0.01, v) - TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r + 0.01, b, v) - TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b - 0.01, v) - TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b + 0.01, v) - TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ETurnbullWakemanAsian = 1 / dS ^ 3 * (TurnbullWakemanAsian(CallPutFlag, S + 2 * dS, SA, X, T, T2, r, b, v) - 3 * TurnbullWakemanAsian(CallPutFlag, S + dS, SA, X, T, T2, r, b, v) _
                                + 3 * TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v) - TurnbullWakemanAsian(CallPutFlag, S - dS, SA, X, T, T2, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S, SA, X + dS, T, T2, r, b, v) - TurnbullWakemanAsian(CallPutFlag, S, SA, X - dS, T, T2, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        ETurnbullWakemanAsian = (TurnbullWakemanAsian(CallPutFlag, S, SA, X + dS, T, T2, r, b, v) - 2 * TurnbullWakemanAsian(CallPutFlag, S, SA, X, T, T2, r, b, v) + TurnbullWakemanAsian(CallPutFlag, S, SA, X - dS, T, T2, r, b, v)) / dS ^ 2
    End If
End Function
