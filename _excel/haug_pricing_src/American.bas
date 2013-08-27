Attribute VB_Name = "American"
Option Explicit

' Programmer Espen Gaarder Haug
' Copyright Espen G. Haug 2006

Public Function PerpetualOption(CallPutFlag As String, S As Double, X _
                As Double, r As Double, b As Double, v As Double) As Double

    Dim y1 As Double, y2 As Double, h As Double
    
    y1 = 1 / 2 - b / v ^ 2 + Sqr((b / v ^ 2 - 1 / 2) ^ 2 + 2 * r / v ^ 2)
    y2 = 1 / 2 - b / v ^ 2 - Sqr((b / v ^ 2 - 1 / 2) ^ 2 + 2 * r / v ^ 2)
    If CallPutFlag = "c" Then
        PerpetualOption = X / (y1 - 1) * ((y1 - 1) / y1 * S / X) ^ y1
    ElseIf CallPutFlag = "p" Then
        PerpetualOption = X / (1 - y2) * ((y2 - 1) / y2 * S / X) ^ y2
    End If
    
End Function


Public Function EBSAmericanApprox2002(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EBSAmericanApprox2002 = BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S + dS, X, T, r, b, v) - BSAmericanApprox2002(CallPutFlag, S - dS, X, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S + dS, X, T, r, b, v) - BSAmericanApprox2002(CallPutFlag, S - dS, X, T, r, b, v)) / (2 * dS) * S / BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S + dS, X, T, r, b, v) - 2 * BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox2002(CallPutFlag, S - dS, X, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S + dS, X, T, r, b, v + 0.01) - 2 * BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v + 0.01) + BSAmericanApprox2002(CallPutFlag, S - dS, X, T, r, b, v + 0.01) _
      - BSAmericanApprox2002(CallPutFlag, S + dS, X, T, r, b, v - 0.01) + 2 * BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v - 0.01) - BSAmericanApprox2002(CallPutFlag, S - dS, X, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
       ElseIf OutPutFlag = "gp" Then 'GammaP
        EBSAmericanApprox2002 = S / 100 * (BSAmericanApprox2002(CallPutFlag, S + dS, X, T, r, b, v) - 2 * BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox2002(CallPutFlag, S - dS, X, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S, X, T + 1 / 365, r, b, v) - 2 * BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox2002(CallPutFlag, S, X, T - 1 / 365, r, b, v)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EBSAmericanApprox2002 = 1 / (4 * dS * 0.01) * (BSAmericanApprox2002(CallPutFlag, S + dS, X, T, r, b, v + 0.01) - BSAmericanApprox2002(CallPutFlag, S + dS, X, T, r, b, v - 0.01) _
        - BSAmericanApprox2002(CallPutFlag, S - dS, X, T, r, b, v + 0.01) + BSAmericanApprox2002(CallPutFlag, S - dS, X, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v + 0.01) - BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v + 0.01) - 2 * BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EBSAmericanApprox2002 = v / 0.1 * (BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v + 0.01) - BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v + 0.01) - 2 * BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EBSAmericanApprox2002 = BSAmericanApprox2002(CallPutFlag, S, X, 0.00001, r, b, v) - BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v)
        Else
                EBSAmericanApprox2002 = BSAmericanApprox2002(CallPutFlag, S, X, T - 1 / 365, r, b, v) - BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v) - BSAmericanApprox2002(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S, X, T, r + 0.01, b, v) - BSAmericanApprox2002(CallPutFlag, S, X, T, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S, X, T, r, b - 0.01, v) - BSAmericanApprox2002(CallPutFlag, S, X, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S, X, T, r, b + 0.01, v) - BSAmericanApprox2002(CallPutFlag, S, X, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EBSAmericanApprox2002 = 1 / dS ^ 3 * (BSAmericanApprox2002(CallPutFlag, S + 2 * dS, X, T, r, b, v) - 3 * BSAmericanApprox2002(CallPutFlag, S + dS, X, T, r, b, v) _
                                + 3 * BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v) - BSAmericanApprox2002(CallPutFlag, S - dS, X, T, r, b, v))
     ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S, X + dS, T, r, b, v) - BSAmericanApprox2002(CallPutFlag, S, X - dS, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        EBSAmericanApprox2002 = (BSAmericanApprox2002(CallPutFlag, S, X + dS, T, r, b, v) - 2 * BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox2002(CallPutFlag, S, X - dS, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "di" Then 'Difference in value between BS Approx and Black-Scholes Merton value
        EBSAmericanApprox2002 = BSAmericanApprox2002(CallPutFlag, S, X, T, r, b, v) - GBlackScholes(CallPutFlag, S, X, T, r, b, v)
    End If
End Function



'// The Bjerksund and Stensland (2002) American approximation
Public Function BSAmericanApprox2002(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, v As Double) As Double
    If CallPutFlag = "c" Then
        BSAmericanApprox2002 = BSAmericanCallApprox2002(S, X, T, r, b, v)
    ElseIf CallPutFlag = "p" Then  '// Use the Bjerksund and Stensland put-call transformation
        BSAmericanApprox2002 = BSAmericanCallApprox2002(X, S, T, r - b, -b, v)
    End If
End Function

'// The Bjerksund and Stensland (1993) American approximation
Public Function BSAmericanApprox(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, v As Double) As Double
    If CallPutFlag = "c" Then
        BSAmericanApprox = BSAmericanCallApprox(S, X, T, r, b, v)
    ElseIf CallPutFlag = "p" Then  '// Use the Bjerksund and Stensland put-call transformation
        BSAmericanApprox = BSAmericanCallApprox(X, S, T, r - b, -b, v)
    End If
End Function

Public Function BSAmericanCallApprox(S As Double, X As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim BInfinity As Double, B0 As Double
    Dim ht As Double, i As Double
    Dim Alpha As Double, Beta As Double
    
    If b >= r Then '// Never optimal to exersice before maturity
            BSAmericanCallApprox = GBlackScholes("c", S, X, T, r, b, v)
    Else
        Beta = (1 / 2 - b / v ^ 2) + Sqr((b / v ^ 2 - 1 / 2) ^ 2 + 2 * r / v ^ 2)
        BInfinity = Beta / (Beta - 1) * X
        B0 = Max(X, r / (r - b) * X)
        ht = -(b * T + 2 * v * Sqr(T)) * B0 / (BInfinity - B0)
        i = B0 + (BInfinity - B0) * (1 - Exp(ht))
        Alpha = (i - X) * i ^ (-Beta)
        If S >= i Then
            BSAmericanCallApprox = S - X
        Else
            BSAmericanCallApprox = Alpha * S ^ Beta - Alpha * phi(S, T, Beta, i, i, r, b, v) + phi(S, T, 1, i, i, r, b, v) - phi(S, T, 1, X, i, r, b, v) - X * phi(S, T, 0, i, i, r, b, v) + X * phi(S, T, 0, X, i, r, b, v)
        End If
    End If
End Function


Public Function BSAmericanCallApprox2002(S As Double, X As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
    Dim BInfinity As Double, B0 As Double
    Dim ht1 As Double, ht2 As Double, I1 As Double, I2 As Double
    Dim alfa1 As Double, alfa2 As Double, Beta As Double, t1 As Double
    
    t1 = 1 / 2 * (Sqr(5) - 1) * T
    
    If b >= r Then  '// Never optimal to exersice before maturity
            BSAmericanCallApprox2002 = GBlackScholes("c", S, X, T, r, b, v)
    Else
        
        Beta = (1 / 2 - b / v ^ 2) + Sqr((b / v ^ 2 - 1 / 2) ^ 2 + 2 * r / v ^ 2)
        BInfinity = Beta / (Beta - 1) * X
        B0 = Max(X, r / (r - b) * X)
        
        ht1 = -(b * t1 + 2 * v * Sqr(t1)) * X ^ 2 / ((BInfinity - B0) * B0)
        ht2 = -(b * T + 2 * v * Sqr(T)) * X ^ 2 / ((BInfinity - B0) * B0)
        I1 = B0 + (BInfinity - B0) * (1 - Exp(ht1))
        I2 = B0 + (BInfinity - B0) * (1 - Exp(ht2))
        alfa1 = (I1 - X) * I1 ^ (-Beta)
        alfa2 = (I2 - X) * I2 ^ (-Beta)
    
        If S >= I2 Then
            BSAmericanCallApprox2002 = S - X
        Else
            BSAmericanCallApprox2002 = alfa2 * S ^ Beta - alfa2 * phi(S, t1, Beta, I2, I2, r, b, v) _
                + phi(S, t1, 1, I2, I2, r, b, v) - phi(S, t1, 1, I1, I2, r, b, v) _
                - X * phi(S, t1, 0, I2, I2, r, b, v) + X * phi(S, t1, 0, I1, I2, r, b, v) _
                + alfa1 * phi(S, t1, Beta, I1, I2, r, b, v) - alfa1 * ksi(S, T, Beta, I1, I2, I1, t1, r, b, v) _
                + ksi(S, T, 1, I1, I2, I1, t1, r, b, v) - ksi(S, T, 1, X, I2, I1, t1, r, b, v) _
                - X * ksi(S, T, 0, I1, I2, I1, t1, r, b, v) + X * ksi(S, T, 0, X, I2, I1, t1, r, b, v)
           
        End If
    End If
End Function
Private Function phi(S As Double, T As Double, gamma As Double, h As Double, i As Double, _
        r As Double, b As Double, v As Double) As Double

    Dim lambda As Double, kappa As Double
    Dim d As Double
    
    lambda = (-r + gamma * b + 0.5 * gamma * (gamma - 1) * v ^ 2) * T
    d = -(Log(S / h) + (b + (gamma - 0.5) * v ^ 2) * T) / (v * Sqr(T))
    kappa = 2 * b / v ^ 2 + 2 * gamma - 1
    phi = Exp(lambda) * S ^ gamma * (CND(d) - (i / S) ^ kappa * CND(d - 2 * Log(i / S) / (v * Sqr(T))))
End Function

' Muligens forskjellig fra phi i Bjerksun Stensland 1993
Private Function phi2(S As Double, T2 As Double, gamma As Double, h As Double, i As Double, _
         r As Double, b As Double, v As Double) As Double

    Dim lambda As Double, kappa As Double
    Dim d As Double, d2 As Double
    
    lambda = -r + gamma * b + 0.5 * gamma * (gamma - 1) * v ^ 2
    kappa = 2 * b / v ^ 2 + 2 * gamma - 1
    
    d = (Log(S / h) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Sqr(T2))
    d2 = (Log(i ^ 2 / (S * h)) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Sqr(T2))
    
    phi2 = Exp(lambda * T2) * S ^ gamma * (CND(-d) - (i / S) ^ kappa * CND(-d2))
    
End Function

Public Function ksi(S As Double, T2 As Double, gamma As Double, h As Double, I2 As Double, I1 As Double, t1 As Double, r As Double, b As Double, v As Double) As Double

    Dim e1 As Double, e2 As Double, e3 As Double, e4 As Double
    Dim f1 As Double, f2 As Double, f3 As Double, f4 As Double
    Dim rho As Double, kappa As Double, lambda As Double
    
    e1 = (Log(S / I1) + (b + (gamma - 0.5) * v ^ 2) * t1) / (v * Sqr(t1))
    e2 = (Log(I2 ^ 2 / (S * I1)) + (b + (gamma - 0.5) * v ^ 2) * t1) / (v * Sqr(t1))
    e3 = (Log(S / I1) - (b + (gamma - 0.5) * v ^ 2) * t1) / (v * Sqr(t1))
    e4 = (Log(I2 ^ 2 / (S * I1)) - (b + (gamma - 0.5) * v ^ 2) * t1) / (v * Sqr(t1))
    
    f1 = (Log(S / h) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Sqr(T2))
    f2 = (Log(I2 ^ 2 / (S * h)) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Sqr(T2))
    f3 = (Log(I1 ^ 2 / (S * h)) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Sqr(T2))
    f4 = (Log(S * I1 ^ 2 / (h * I2 ^ 2)) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Sqr(T2))
    
    rho = Sqr(t1 / T2)
    lambda = -r + gamma * b + 0.5 * gamma * (gamma - 1) * v ^ 2
    kappa = 2 * b / (v ^ 2) + (2 * gamma - 1)
    
    ksi = Exp(lambda * T2) * S ^ gamma * (CBND(-e1, -f1, rho) - (I2 / S) ^ kappa * CBND(-e2, -f2, rho) _
            - (I1 / S) ^ kappa * CBND(-e3, -f3, -rho) + (I1 / I2) ^ kappa * CBND(-e4, -f4, -rho))


End Function


Public Function EBSAmericanApprox(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EBSAmericanApprox = BSAmericanApprox(CallPutFlag, S, X, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v) - BSAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v) - BSAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v)) / (2 * dS) * S / BSAmericanApprox(CallPutFlag, S, X, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v) - 2 * BSAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v + 0.01) - 2 * BSAmericanApprox(CallPutFlag, S, X, T, r, b, v + 0.01) + BSAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v + 0.01) _
      - BSAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v - 0.01) + 2 * BSAmericanApprox(CallPutFlag, S, X, T, r, b, v - 0.01) - BSAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
       ElseIf OutPutFlag = "gp" Then 'GammaP
        EBSAmericanApprox = S / 100 * (BSAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v) - 2 * BSAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S, X, T + 1 / 365, r, b, v) - 2 * BSAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox(CallPutFlag, S, X, T - 1 / 365, r, b, v)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EBSAmericanApprox = 1 / (4 * dS * 0.01) * (BSAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v + 0.01) - BSAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v - 0.01) _
        - BSAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v + 0.01) + BSAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S, X, T, r, b, v + 0.01) - BSAmericanApprox(CallPutFlag, S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S, X, T, r, b, v + 0.01) - 2 * BSAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox(CallPutFlag, S, X, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EBSAmericanApprox = v / 0.1 * (BSAmericanApprox(CallPutFlag, S, X, T, r, b, v + 0.01) - BSAmericanApprox(CallPutFlag, S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S, X, T, r, b, v + 0.01) - 2 * BSAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox(CallPutFlag, S, X, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EBSAmericanApprox = BSAmericanApprox(CallPutFlag, S, X, 0.00001, r, b, v) - BSAmericanApprox(CallPutFlag, S, X, T, r, b, v)
        Else
                EBSAmericanApprox = BSAmericanApprox(CallPutFlag, S, X, T - 1 / 365, r, b, v) - BSAmericanApprox(CallPutFlag, S, X, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v) - BSAmericanApprox(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S, X, T, r + 0.01, b, v) - BSAmericanApprox(CallPutFlag, S, X, T, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S, X, T, r, b - 0.01, v) - BSAmericanApprox(CallPutFlag, S, X, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S, X, T, r, b + 0.01, v) - BSAmericanApprox(CallPutFlag, S, X, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EBSAmericanApprox = 1 / dS ^ 3 * (BSAmericanApprox(CallPutFlag, S + 2 * dS, X, T, r, b, v) - 3 * BSAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v) _
                                + 3 * BSAmericanApprox(CallPutFlag, S, X, T, r, b, v) - BSAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v))
     ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S, X + dS, T, r, b, v) - BSAmericanApprox(CallPutFlag, S, X - dS, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        EBSAmericanApprox = (BSAmericanApprox(CallPutFlag, S, X + dS, T, r, b, v) - 2 * BSAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BSAmericanApprox(CallPutFlag, S, X - dS, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "di" Then 'Difference in value between BS Approx and Black-Scholes Merton value
        EBSAmericanApprox = BSAmericanApprox(CallPutFlag, S, X, T, r, b, v) - GBlackScholes(CallPutFlag, S, X, T, r, b, v)
   ElseIf OutPutFlag = "BSIVol" Then 'Equivalent Black-Scholes-Merton implied volatility
       CallPutFlag = "p"
       If S >= S * Exp(b * T) Then CallPutFlag = "c"
        EBSAmericanApprox = ImpliedVolGBlackScholes(CallPutFlag, S, X, T, r, b, BSAmericanApprox(CallPutFlag, S, X, T, r, b, v))
    End If
End Function



Public Function EPerpetualOption(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EPerpetualOption = PerpetualOption(CallPutFlag, S, X, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EPerpetualOption = (PerpetualOption(CallPutFlag, S + dS, X, r, b, v) - PerpetualOption(CallPutFlag, S - dS, X, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EPerpetualOption = (PerpetualOption(CallPutFlag, S + dS, X, r, b, v) - PerpetualOption(CallPutFlag, S - dS, X, r, b, v)) / (2 * dS) * S / PerpetualOption(CallPutFlag, S, X, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EPerpetualOption = (PerpetualOption(CallPutFlag, S + dS, X, r, b, v) - 2 * PerpetualOption(CallPutFlag, S, X, r, b, v) + PerpetualOption(CallPutFlag, S - dS, X, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EPerpetualOption = (PerpetualOption(CallPutFlag, S + dS, X, r, b, v + 0.01) - 2 * PerpetualOption(CallPutFlag, S, X, r, b, v + 0.01) + PerpetualOption(CallPutFlag, S - dS, X, r, b, v + 0.01) _
      - PerpetualOption(CallPutFlag, S + dS, X, r, b, v - 0.01) + 2 * PerpetualOption(CallPutFlag, S, X, r, b, v - 0.01) - PerpetualOption(CallPutFlag, S - dS, X, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
       ElseIf OutPutFlag = "gp" Then 'GammaP
        EPerpetualOption = S / 100 * (PerpetualOption(CallPutFlag, S + dS, X, r, b, v) - 2 * PerpetualOption(CallPutFlag, S, X, r, b, v) + PerpetualOption(CallPutFlag, S - dS, X, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EPerpetualOption = 1 / (4 * dS * 0.01) * (PerpetualOption(CallPutFlag, S + dS, X, r, b, v + 0.01) - PerpetualOption(CallPutFlag, S + dS, X, r, b, v - 0.01) _
        - PerpetualOption(CallPutFlag, S - dS, X, r, b, v + 0.01) + PerpetualOption(CallPutFlag, S - dS, X, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EPerpetualOption = (PerpetualOption(CallPutFlag, S, X, r, b, v + 0.01) - PerpetualOption(CallPutFlag, S, X, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EPerpetualOption = (PerpetualOption(CallPutFlag, S, X, r, b, v + 0.01) - 2 * PerpetualOption(CallPutFlag, S, X, r, b, v) + PerpetualOption(CallPutFlag, S, X, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EPerpetualOption = v / 0.1 * (PerpetualOption(CallPutFlag, S, X, r, b, v + 0.01) - PerpetualOption(CallPutFlag, S, X, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EPerpetualOption = (PerpetualOption(CallPutFlag, S, X, r, b, v + 0.01) - 2 * PerpetualOption(CallPutFlag, S, X, r, b, v) + PerpetualOption(CallPutFlag, S, X, r, b, v - 0.01))
     ElseIf OutPutFlag = "r" Then 'Rho
         EPerpetualOption = (PerpetualOption(CallPutFlag, S, X, r + 0.01, b + 0.01, v) - PerpetualOption(CallPutFlag, S, X, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EPerpetualOption = (PerpetualOption(CallPutFlag, S, X, r + 0.01, b, v) - PerpetualOption(CallPutFlag, S, X, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EPerpetualOption = (PerpetualOption(CallPutFlag, S, X, r, b - 0.01, v) - PerpetualOption(CallPutFlag, S, X, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EPerpetualOption = (PerpetualOption(CallPutFlag, S, X, r, b + 0.01, v) - PerpetualOption(CallPutFlag, S, X, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EPerpetualOption = 1 / dS ^ 3 * (PerpetualOption(CallPutFlag, S + 2 * dS, X, r, b, v) - 3 * PerpetualOption(CallPutFlag, S + dS, X, r, b, v) _
                                + 3 * PerpetualOption(CallPutFlag, S, X, r, b, v) - PerpetualOption(CallPutFlag, S - dS, X, r, b, v))
    ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EPerpetualOption = (PerpetualOption(CallPutFlag, S, X + dS, r, b, v) - PerpetualOption(CallPutFlag, S, X - dS, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        EPerpetualOption = (PerpetualOption(CallPutFlag, S, X + dS, r, b, v) - 2 * PerpetualOption(CallPutFlag, S, X, r, b, v) + PerpetualOption(CallPutFlag, S, X - dS, r, b, v)) / dS ^ 2
    End If
End Function


'// The Barone-Adesi and Whaley (1987) American approximation
Public Function BAWAmericanApprox(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, v As Double) As Double
    If CallPutFlag = "c" Then
        BAWAmericanApprox = BAWAmericanCallApprox(S, X, T, r, b, v)
    ElseIf CallPutFlag = "p" Then
        BAWAmericanApprox = BAWAmericanPutApprox(S, X, T, r, b, v)
    End If
End Function

'// American call
Private Function BAWAmericanCallApprox(S As Double, X As Double, T As Double, r As Double, b As Double, v As Double) As Double

    Dim Sk As Double, N As Double, k As Double
    Dim d1 As Double, Q2 As Double, a2 As Double

    If b >= r Then
        BAWAmericanCallApprox = GBlackScholes("c", S, X, T, r, b, v)
    Else
        Sk = Kc(X, T, r, b, v)
        N = 2 * b / v ^ 2                                           '
        k = 2 * r / (v ^ 2 * (1 - Exp(-r * T)))
        d1 = (Log(Sk / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
        Q2 = (-(N - 1) + Sqr((N - 1) ^ 2 + 4 * k)) / 2
        a2 = (Sk / Q2) * (1 - Exp((b - r) * T) * CND(d1))
        If S < Sk Then
            BAWAmericanCallApprox = GBlackScholes("c", S, X, T, r, b, v) + a2 * (S / Sk) ^ Q2
        Else
            BAWAmericanCallApprox = S - X
        End If
    End If
End Function

'// Newton Raphson algorithm to solve for the critical commodity price for a Call
Private Function Kc(X As Double, T As Double, r As Double, b As Double, v As Double) As Double

    Dim N As Double, m As Double
    Dim su As Double, Si As Double
    Dim h2 As Double, k As Double
    Dim d1 As Double, Q2 As Double, q2u As Double
    Dim LHS As Double, RHS As Double
    Dim bi As Double, E As Double
    
    '// Calculation of seed value, Si
    N = 2 * b / v ^ 2
    m = 2 * r / v ^ 2
    q2u = (-(N - 1) + Sqr((N - 1) ^ 2 + 4 * m)) / 2
    su = X / (1 - 1 / q2u)
    h2 = -(b * T + 2 * v * Sqr(T)) * X / (su - X)
    Si = X + (su - X) * (1 - Exp(h2))

    k = 2 * r / (v ^ 2 * (1 - Exp(-r * T)))
    d1 = (Log(Si / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    Q2 = (-(N - 1) + Sqr((N - 1) ^ 2 + 4 * k)) / 2
    LHS = Si - X
    RHS = GBlackScholes("c", Si, X, T, r, b, v) + (1 - Exp((b - r) * T) * CND(d1)) * Si / Q2
    bi = Exp((b - r) * T) * CND(d1) * (1 - 1 / Q2) + (1 - Exp((b - r) * T) * CND(d1) / (v * Sqr(T))) / Q2
    E = 0.000001
    '// Newton Raphson algorithm for finding critical price Si
    While Abs(LHS - RHS) / X > E
        Si = (X + RHS - bi * Si) / (1 - bi)
        d1 = (Log(Si / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
        LHS = Si - X
        RHS = GBlackScholes("c", Si, X, T, r, b, v) + (1 - Exp((b - r) * T) * CND(d1)) * Si / Q2
        bi = Exp((b - r) * T) * CND(d1) * (1 - 1 / Q2) + (1 - Exp((b - r) * T) * ND(d1) / (v * Sqr(T))) / Q2
    Wend
        Kc = Si
End Function
'// American put
Private Function BAWAmericanPutApprox(S As Double, X As Double, T As Double, r As Double, b As Double, v As Double) As Double

    Dim Sk As Double, N As Double, k As Double
    Dim d1 As Double, Q1 As Double, a1 As Double

    Sk = Kp(X, T, r, b, v)
    N = 2 * b / v ^ 2
    k = 2 * r / (v ^ 2 * (1 - Exp(-r * T)))
    d1 = (Log(Sk / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    Q1 = (-(N - 1) - Sqr((N - 1) ^ 2 + 4 * k)) / 2
    a1 = -(Sk / Q1) * (1 - Exp((b - r) * T) * CND(-d1))

    If S > Sk Then
        BAWAmericanPutApprox = GBlackScholes("p", S, X, T, r, b, v) + a1 * (S / Sk) ^ Q1
    Else
        BAWAmericanPutApprox = X - S
    End If
End Function

'// Newton Raphson algorithm to solve for the critical commodity price for a Put
Private Function Kp(X As Double, T As Double, r As Double, b As Double, v As Double) As Double
    
   
    Dim N As Double, m As Double
    Dim su As Double, Si As Double
    Dim h1 As Double, k As Double
    Dim d1 As Double, q1u As Double, Q1 As Double
    Dim LHS As Double, RHS As Double
    Dim bi As Double, E As Double
    
    '// Calculation of seed value, Si
    N = 2 * b / v ^ 2
    m = 2 * r / v ^ 2
    q1u = (-(N - 1) - Sqr((N - 1) ^ 2 + 4 * m)) / 2
    su = X / (1 - 1 / q1u)
    h1 = (b * T - 2 * v * Sqr(T)) * X / (X - su)
    Si = su + (X - su) * Exp(h1)

    
    k = 2 * r / (v ^ 2 * (1 - Exp(-r * T)))
    d1 = (Log(Si / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    Q1 = (-(N - 1) - Sqr((N - 1) ^ 2 + 4 * k)) / 2
    LHS = X - Si
    RHS = GBlackScholes("p", Si, X, T, r, b, v) - (1 - Exp((b - r) * T) * CND(-d1)) * Si / Q1
    bi = -Exp((b - r) * T) * CND(-d1) * (1 - 1 / Q1) - (1 + Exp((b - r) * T) * ND(-d1) / (v * Sqr(T))) / Q1
    E = 0.000001
    '// Newton Raphson algorithm for finding critical price Si
    While Abs(LHS - RHS) / X > E
        Si = (X - RHS + bi * Si) / (1 + bi)
        d1 = (Log(Si / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
        LHS = X - Si
        RHS = GBlackScholes("p", Si, X, T, r, b, v) - (1 - Exp((b - r) * T) * CND(-d1)) * Si / Q1
        bi = -Exp((b - r) * T) * CND(-d1) * (1 - 1 / Q1) - (1 + Exp((b - r) * T) * CND(-d1) / (v * Sqr(T))) / Q1
    Wend
    Kp = Si
End Function


Public Function EBAWAmericanApprox(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EBAWAmericanApprox = BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v) - BAWAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v) - BAWAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v)) / (2 * dS) * S / BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v) - 2 * BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BAWAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v + 0.01) - 2 * BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v + 0.01) + BAWAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v + 0.01) _
      - BAWAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v - 0.01) + 2 * BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v - 0.01) - BAWAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
    ElseIf OutPutFlag = "gp" Then 'GammaP
        EBAWAmericanApprox = S / 100 * (BAWAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v) - 2 * BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BAWAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v)) / dS ^ 2
     ElseIf OutPutFlag = "tg" Then 'time Gamma
        EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S, X, T + 1 / 365, r, b, v) - 2 * BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BAWAmericanApprox(CallPutFlag, S, X, T - 1 / 365, r, b, v)) / (1 / 365) ^ 2
     ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EBAWAmericanApprox = 1 / (4 * dS * 0.01) * (BAWAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v + 0.01) - BAWAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v - 0.01) _
        - BAWAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v + 0.01) + BAWAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v + 0.01) - BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v + 0.01) - 2 * BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EBAWAmericanApprox = v / 0.1 * (BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v + 0.01) - BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v + 0.01) - 2 * BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EBAWAmericanApprox = BAWAmericanApprox(CallPutFlag, S, X, 0.00001, r, b, v) - BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v)
        Else
                EBAWAmericanApprox = BAWAmericanApprox(CallPutFlag, S, X, T - 1 / 365, r, b, v) - BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S, X, T, r + 0.01, b + 0.01, v) - BAWAmericanApprox(CallPutFlag, S, X, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures options rho
         EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S, X, T, r + 0.01, b, v) - BAWAmericanApprox(CallPutFlag, S, X, T, r - 0.01, b, v)) / (2)
     ElseIf OutPutFlag = "f" Then 'Rho2
         EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S, X, T, r, b - 0.01, v) - BAWAmericanApprox(CallPutFlag, S, X, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S, X, T, r, b + 0.01, v) - BAWAmericanApprox(CallPutFlag, S, X, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EBAWAmericanApprox = 1 / dS ^ 3 * (BAWAmericanApprox(CallPutFlag, S + 2 * dS, X, T, r, b, v) - 3 * BAWAmericanApprox(CallPutFlag, S + dS, X, T, r, b, v) _
                                + 3 * BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v) - BAWAmericanApprox(CallPutFlag, S - dS, X, T, r, b, v))
       ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S, X + dS, T, r, b, v) - BAWAmericanApprox(CallPutFlag, S, X - dS, T, r, b, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        EBAWAmericanApprox = (BAWAmericanApprox(CallPutFlag, S, X + dS, T, r, b, v) - 2 * BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v) + BAWAmericanApprox(CallPutFlag, S, X - dS, T, r, b, v)) / dS ^ 2
      ElseIf OutPutFlag = "di" Then 'Difference in value between BS Approx and Black-Scholes Merton value
        EBAWAmericanApprox = BAWAmericanApprox(CallPutFlag, S, X, T, r, b, v) - GBlackScholes(CallPutFlag, S, X, T, r, b, v)
    End If
End Function

