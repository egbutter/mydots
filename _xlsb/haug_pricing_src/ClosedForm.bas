Attribute VB_Name = "ClosedForm"
Option Explicit
Global Const Pi = 3.14159265358979


' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug


'// American Calls on stocks with known dividends, Roll-Geske-Whaley
Public Function RollGeskeWhaley(S As Double, X As Double, t1 As Double, T2 As Double, r As Double, d As Double, v As Double) As Double
    't1 time to dividend payout
    'T2 time to option expiration
    
    Dim Sx As Double, i As Double
    Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double
    Dim HighS As Double, LowS As Double, epsilon As Double
    Dim ci As Double, infinity As Double
    
    infinity = 1000000000
    epsilon = 0.00000001
    Sx = S - d * Exp(-r * t1)
    If d <= X * (1 - Exp(-r * (T2 - t1))) Then '// Not optimal to exercise
        RollGeskeWhaley = GBlackScholes("c", Sx, X, T2, r, r, v)
        Exit Function
    End If
    ci = GBlackScholes("c", S, X, T2 - t1, r, r, v)
    HighS = S
    While (ci - HighS - d + X) > 0 And HighS < infinity
        HighS = HighS * 2
        ci = GBlackScholes("c", HighS, X, T2 - t1, r, r, v)
    Wend
    If HighS > infinity Then
        RollGeskeWhaley = GBlackScholes("c", Sx, X, T2, r, r, v)
        Exit Function
    End If
    
    LowS = 0
    i = HighS * 0.5
    ci = GBlackScholes("c", i, X, T2 - t1, r, r, v)
    
    '// Search algorithm to find the critical stock price I
    While Abs(ci - i - d + X) > epsilon And HighS - LowS > epsilon
        If (ci - i - d + X) < 0 Then
            HighS = i
        Else
            LowS = i
        End If
        i = (HighS + LowS) / 2
        ci = GBlackScholes("c", i, X, T2 - t1, r, r, v)
    Wend
    a1 = (Log(Sx / X) + (r + v ^ 2 / 2) * T2) / (v * Sqr(T2))
    a2 = a1 - v * Sqr(T2)
    b1 = (Log(Sx / i) + (r + v ^ 2 / 2) * t1) / (v * Sqr(t1))
    b2 = b1 - v * Sqr(t1)
   
    RollGeskeWhaley = Sx * CND(b1) + Sx * CBND(a1, -b1, -Sqr(t1 / T2)) - X * Exp(-r * T2) * CBND(a2, -b2, -Sqr(t1 / T2)) - (X - d) * Exp(-r * t1) * CND(b2)

End Function




'// The Black-Scholes formula adjusted for discrete dividend yield
Public Function EEuropeanDiscreteDividendYield(OutPutFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, Dy As Double, n As Integer, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    S = S * (1 - Dy) ^ n
    
    If OutPutFlag = "p" Then ' Value
        EEuropeanDiscreteDividendYield = GBlackScholes(CallPutFlag, S, X, T, r, r, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EEuropeanDiscreteDividendYield = (GBlackScholes(CallPutFlag, S + dS, X, T, r, r, v) - GBlackScholes(CallPutFlag, S - dS, X, T, r, r, v)) / (2 * dS)
    ElseIf OutPutFlag = "e" Then 'Elasticity
         EEuropeanDiscreteDividendYield = (GBlackScholes(CallPutFlag, S + dS, X, T, r, r, v) - GBlackScholes(CallPutFlag, S - dS, X, T, r, r, v)) / (2 * dS) * S / GBlackScholes(CallPutFlag, S, X, T, r, r, v)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EEuropeanDiscreteDividendYield = (GBlackScholes(CallPutFlag, S + dS, X, T, r, r, v) - 2 * GBlackScholes(CallPutFlag, S, X, T, r, r, v) + GBlackScholes(CallPutFlag, S - dS, X, T, r, r, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EEuropeanDiscreteDividendYield = (GBlackScholes(CallPutFlag, S + dS, X, T, r, r, v + 0.01) - 2 * GBlackScholes(CallPutFlag, S, X, T, r, r, v + 0.01) + GBlackScholes(CallPutFlag, S - dS, X, T, r, r, v + 0.01) _
      - GBlackScholes(CallPutFlag, S + dS, X, T, r, r, v - 0.01) + 2 * GBlackScholes(CallPutFlag, S, X, T, r, r, v - 0.01) - GBlackScholes(CallPutFlag, S - dS, X, T, r, r, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "gp" Then 'GammaP
        EEuropeanDiscreteDividendYield = S / 100 * (GBlackScholes(CallPutFlag, S + dS, X, T, r, r, v) - 2 * GBlackScholes(CallPutFlag, S, X, T, r, r, v) + GBlackScholes(CallPutFlag, S - dS, X, T, r, r, v)) / dS ^ 2
        ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EEuropeanDiscreteDividendYield = 1 / (4 * dS * 0.01) * (GBlackScholes(CallPutFlag, S + dS, X, T, r, r, v + 0.01) - GBlackScholes(CallPutFlag, S + dS, X, T, r, r, v - 0.01) _
        - GBlackScholes(CallPutFlag, S - dS, X, T, r, r, v + 0.01) + GBlackScholes(CallPutFlag, S - dS, X, T, r, r, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EEuropeanDiscreteDividendYield = (GBlackScholes(CallPutFlag, S, X, T, r, r, v + 0.01) - GBlackScholes(CallPutFlag, S, X, T, r, r, v - 0.01)) / 2
     ElseIf OutPutFlag = "vv" Then 'DvegaDvol/vomma
        EEuropeanDiscreteDividendYield = (GBlackScholes(CallPutFlag, S, X, T, r, r, v + 0.01) - 2 * GBlackScholes(CallPutFlag, S, X, T, r, r, v) + GBlackScholes(CallPutFlag, S, X, T, r, r, v - 0.01)) / 0.01 ^ 2 / 10000
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EEuropeanDiscreteDividendYield = v / 0.1 * (GBlackScholes(CallPutFlag, S, X, T, r, r, v + 0.01) - GBlackScholes(CallPutFlag, S, X, T, r, r, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EEuropeanDiscreteDividendYield = (GBlackScholes(CallPutFlag, S, X, T, r, r, v + 0.01) - 2 * GBlackScholes(CallPutFlag, S, X, T, r, r, v) + GBlackScholes(CallPutFlag, S, X, T, r, r, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
                EEuropeanDiscreteDividendYield = GBlackScholes(CallPutFlag, S, X, 0.00001, r, r, v) - GBlackScholes(CallPutFlag, S, X, T, r, r, v)
        Else
                EEuropeanDiscreteDividendYield = GBlackScholes(CallPutFlag, S, X, T - 1 / 365, r, r, v) - GBlackScholes(CallPutFlag, S, X, T, r, r, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EEuropeanDiscreteDividendYield = (GBlackScholes(CallPutFlag, S, X, T, r + 0.01, r + 0.01, v) - GBlackScholes(CallPutFlag, S, X, T, r - 0.01, r - 0.01, v)) / (2)
         ElseIf OutPutFlag = "r" Then 'Rho
         EEuropeanDiscreteDividendYield = (GBlackScholes(CallPutFlag, S, X, T, r + 0.01, r + 0.01, v) - GBlackScholes(CallPutFlag, S, X, T, r - 0.01, r - 0.01, v)) / (2)
      ElseIf OutPutFlag = "s" Then 'Speed
        EEuropeanDiscreteDividendYield = 1 / dS ^ 3 * (GBlackScholes(CallPutFlag, S + 2 * dS, X, T, r, r, v) - 3 * GBlackScholes(CallPutFlag, S + dS, X, T, r, r, v) _
                                + 3 * GBlackScholes(CallPutFlag, S, X, T, r, r, v) - GBlackScholes(CallPutFlag, S - dS, X, T, r, r, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EEuropeanDiscreteDividendYield = (GBlackScholes(CallPutFlag, S, X + dS, T, r, r, v) - GBlackScholes(CallPutFlag, S, X - dS, T, r, r, v)) / (2 * dS)
     ElseIf OutPutFlag = "dxdx" Then 'Gamma
        EEuropeanDiscreteDividendYield = (GBlackScholes(CallPutFlag, S, X + dS, T, r, r, v) - 2 * GBlackScholes(CallPutFlag, S, X, T, r, r, v) + GBlackScholes(CallPutFlag, S, X - dS, T, r, r, v)) / dS ^ 2
    End If
End Function


'// American Calls on stocks with known discrete dividend yield, Villiger (2005)
'// Implementation By Espen Gaarder Haug
Public Function DiscreteDividenYieldAnalytic(EurAmeFlag As String, S As Double, X As Double, t1 As Double, T2 As Double, r As Double, Dy As Double, v As Double) As Double
    't1 time to dividend payout
    'T2 time to option expiration
    
    Dim Sx As Double, i As Double
    Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double
    Dim HighS As Double, LowS As Double, epsilon As Double
    Dim ci As Double, infinity As Double
    
    infinity = 1000000000
    epsilon = 0.00000001
    Sx = S * (1 - Dy)
    If EurAmeFlag = "e" Or S * (1 - Dy) <= X * (1 - Exp(-r * (T2 - t1))) Then '// Not optimal to exercise
        DiscreteDividenYieldAnalytic = GBlackScholes("c", Sx, X, T2, r, r, v)
        Exit Function
    End If
    ci = GBlackScholes("c", S, X, T2 - t1, r, r, v)
    HighS = S
    While (ci - HighS + X) > 0 And HighS < infinity
        HighS = HighS * 2
        ci = GBlackScholes("c", HighS * (1 - Dy), X, T2 - t1, r, r, v)
    Wend
    If HighS > infinity Then
        DiscreteDividenYieldAnalytic = GBlackScholes("c", Sx, X, T2, r, r, v)
        Exit Function
    End If
    
    LowS = 0
    i = HighS * 0.5
    ci = GBlackScholes("c", i * (1 - Dy), X, T2 - t1, r, r, v)
    
    '// Search algorithm to find the critical stock price I
    While Abs(ci - i + X) > epsilon And HighS - LowS > epsilon
        If (ci - i + X) < 0 Then
            HighS = i
        Else
            LowS = i
        End If
        i = (HighS + LowS) / 2
        ci = GBlackScholes("c", i * (1 - Dy), X, T2 - t1, r, r, v)
    Wend
    
    a1 = (Log(Sx / X) + (r + v ^ 2 / 2) * T2) / (v * Sqr(T2))
    a2 = a1 - v * Sqr(T2)
    b1 = (Log(S / i) + (r + v ^ 2 / 2) * t1) / (v * Sqr(t1))
    b2 = b1 - v * Sqr(t1)
   
    DiscreteDividenYieldAnalytic = S * CND(b1) + Sx * CBND(a1, -b1, -Sqr(t1 / T2)) _
        - X * Exp(-r * T2) * CBND(a2, -b2, -Sqr(t1 / T2)) - X * Exp(-r * t1) * CND(b2)
    
  End Function



Public Function BosVandermarkCashDividend(CallPutFlag As String, S As Double, X As Double _
                , T As Double, r As Double, b As Double, v As Double, Optional Dividends As Object, Optional DividendTimes As Object) As Double

        Dim i As Integer, n As Integer
        Dim Xn As Double, Xf As Double

        n = Application.Count(Dividends)
        Xn = 0
        Xf = 0

        For i = 1 To n
            Xn = Xn + (T - DividendTimes(i)) / T * Dividends(i) * Exp(-r * DividendTimes(i))
            Xf = Xf + (DividendTimes(i)) / T * Dividends(i) * Exp(-r * DividendTimes(i))
        Next
        
        BosVandermarkCashDividend = GBlackScholes(CallPutFlag, S - Xn, X + Xf * Exp(r * T), T, r, b, v)
       

End Function


Public Function BosGaiShepVol(S As Double, X As Double, T As Double, r As Double, v As Double, Optional DividendTimes As Object, Optional Dividends As Object)

    Dim n As Integer, i As Integer, j As Integer
    Dim sum1 As Double, sum2 As Double
    Dim z1 As Double, z2 As Double
    Dim ti As Double, tj As Double
    Dim dt As Double
    
    n = Application.Count(Dividends)
    dt = 0
    For i = 1 To n
        dt = dt + Dividends(i) * Exp(-r * DividendTimes(i))
    Next
        
    S = Log(S)
    X = Log((X + dt) * Exp(-r * T))
    
    z1 = (S - X) / (v * Sqr(T)) + v * Sqr(T) / 2
    z2 = (S - X) / (v * Sqr(T)) + v * Sqr(T)
    
    sum1 = 0
    sum2 = 0
    
    For i = 1 To n
        ti = DividendTimes(i)
        sum1 = sum1 + Dividends(i) * Exp(-r * ti) * (CND(z1) - CND(z1 - v * ti / Sqr(T)))
        For j = 1 To n
            tj = DividendTimes(j)
            sum2 = sum2 + Dividends(i) * Dividends(j) * Exp(-r * (ti + tj)) * (CND(z2) - CND(z2 - 2 * v * Min(ti, tj) / Sqr(T)))
        Next
       
    Next
    
    BosGaiShepVol = Sqr(v ^ 2 + v * Sqr(Pi / (2 * T)) * (4 * Exp(z1 ^ 2 / 2 - S) * sum1 + Exp(z2 ^ 2 / 2 - 2 * S) * sum2))
    
End Function

'// Equity forward price with same rate for each cash low or different rate for each cashflow.
Function EquityForwardPrice(TradeDate As Double, ForwardDate As Double, S As Double, r As Double, DividendDates As Object, Dividends As Object, Optional rDividendDates) As Double
        
        '// Variable description:
        '// TradeDate: the date the trade is done
        '// ForwardDate: the date for calculating the forward price
        '// S: stock price spot
        '// r: risk-free rate, a continuous compounding zero coupon rate with maturity equal to the forward period
        '// DividendDates: a single or hole array of the dates the stock payes dividend
        '// Dividends: cash dividends at each dividend date
        '// rDividendDates: array of continuous compunding zero coupon rates. One for each dividend date, if not given the
        '// Function will use the same rate r for all cash flows
        
        Dim NoOfDividends As Integer
        Dim NoOfDividendTimes As Integer
        Dim i As Integer
        
        NoOfDividendTimes = Application.Count(DividendDates)
        NoOfDividends = Application.Count(Dividends)
    
       For i = 1 To NoOfDividendTimes
            If DividendDates(i) <= ForwardDate Then
                If IsMissing(rDividendDates) Then
                    S = S - Dividends(i) * Exp(-r * (DividendDates(i) - TradeDate) / 365)
                Else
                    S = S - Dividends(i) * Exp(-rDividendDates(i) * (DividendDates(i) - TradeDate) / 365)
                End If
            Else
                Exit For
            End If
       Next
       
       EquityForwardPrice = S * Exp(r * (ForwardDate - TradeDate) / 365)

End Function


'// Equity forward price with same rate for each cash low or different rate for each cashflow.
Function StockMinusNPVDividend(TradeDate As Double, ExpiryDate As Double, S As Double, r As Double, DividendDates As Object, Dividends As Object, Optional rDividendDates) As Double
        
        '// Variable description
        '// S: stock price spot
        '// r: risk-free rate, a continuous compounding zero coupon rate with maturity equal to the forward period
        '// DividendDates: a single or whole array of the dates the stock payes dividend
        '// Dividends: cash dividends at each dividend date
        '// rDividendDates: array of continuous compunding zero coupon rates. One for each dividend date, if not given the
        '// function will use the same rate r for all cash flows
        
        Dim NoOfDividends As Integer
        Dim NoOfDividendTimes As Integer
        Dim i As Integer
        
        NoOfDividendTimes = Application.Count(DividendDates)
        NoOfDividends = Application.Count(Dividends)
    
       For i = 1 To NoOfDividendTimes
            If DividendDates(i) <= ExpiryDate Then
                If IsMissing(rDividendDates) Then
                    S = S - Dividends(i) * Exp(-r * (DividendDates(i) - TradeDate) / 365)
                Else
                    S = S - Dividends(i) * Exp(-rDividendDates(i) * (DividendDates(i) - TradeDate) / 365)
                End If
            Else
                Exit For
            End If
       Next
       
       StockMinusNPVDividend = S

End Function


'// Equity forward price with same rate for each cash low or different rate for each cashflow.
Function StockMinusNPVDividend2(ExpiryDate As Double, S As Double, r As Double, DividendDates As Object, Dividends As Object, Optional rDividendDates) As Double
        
        '// Variable description
        '// S: stock price spot
        '// r: risk-free rate, a continuous compounding zero coupon rate with maturity equal to the forward period
        '// DividendDates: a single or whole array of the dates the stock payes dividend
        '// Dividends: cash dividends at each dividend date
        '// rDividendDates: array of continuous compunding zero coupon rates. One for each dividend date, if not given the
        '// function will use the same rate r for all cash flows
        
        Dim NoOfDividends As Integer
        Dim NoOfDividendTimes As Integer
        Dim i As Integer
        
        NoOfDividendTimes = Application.Count(DividendDates)
        NoOfDividends = Application.Count(Dividends)
    
       For i = 1 To NoOfDividendTimes
            If DividendDates(i) <= ExpiryDate Then
                If IsMissing(rDividendDates) Then
                    S = S - Dividends(i) * Exp(-r * DividendDates(i))
                Else
                    S = S - Dividends(i) * Exp(-rDividendDates(i) * DividendDates(i))
                End If
            Else
                Exit For
            End If
       Next
       
       StockMinusNPVDividend2 = S

End Function


'
 'Volatility adjusted for discrete dividends European model
Public Function HaugHaugVol(S As Double, T As Double, r As Double, _
    Dividends As Variant, DividendTimes As Variant, v As Double) As Double
    
    Dim SumDividends As Double, sumVolatilities As Double
    Dim n As Integer, j As Integer, i As Integer
    
    n = Application.Count(Dividends) ' number of dividends
    
    sumVolatilities = 0
    For j = 1 To n + 1
        SumDividends = 0
        For i = j To n
             SumDividends = SumDividends + Dividends(i) * Exp(-r * DividendTimes(i))
        Next
        If j = 1 Then
                    sumVolatilities = sumVolatilities + (S * v / (S - SumDividends)) ^ 2 * DividendTimes(j)
        ElseIf j < n + 1 Then
            sumVolatilities = sumVolatilities + (S * v / (S - SumDividends)) ^ 2 * (DividendTimes(j) - DividendTimes(j - 1))
        Else
            sumVolatilities = sumVolatilities + v ^ 2 * (T - DividendTimes(j - 1))
        End If
    Next
    HaugHaugVol = Sqr(sumVolatilities / T)
End Function



'// The generalized Black and Scholes formula
Public Function GBlackScholes(CallPutFlag As String, S As Double, X _
                As Double, T As Double, r As Double, b As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then
        GBlackScholes = S * Exp((b - r) * T) * CND(d1) - X * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
        GBlackScholes = X * Exp(-r * T) * CND(-d2) - S * Exp((b - r) * T) * CND(-d1)
    End If
End Function
