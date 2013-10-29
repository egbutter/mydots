Attribute VB_Name = "FixedIncome"
Global Const Pi = 3.14159265358979

Option Explicit     'Requirs that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.

' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug

Public Function YieldVolToPriceVol(YieldVol As Double, Duration As Double, BondYield As Double) As Double
    
    YieldVolToPriceVol = YieldVol * BondYield * Duration / (1 + BondYield)

End Function

Public Function PriceVolToYieldVol(PriceVol As Double, Duration As Double, BondYield As Double) As Double
    
    PriceVolToYieldVol = PriceVol / (BondYield * Duration / (1 + BondYield))

End Function


'//  Black-76 European swaption
Public Function Swaption(CallPutFlag As String, t1 As Double, m As Double, F As Double, X As Double, T As Double, _
                r As Double, v As Double) As Double
 
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(F / X) + v ^ 2 / 2 * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then 'Payer swaption
        Swaption = ((1 - 1 / (1 + F / m) ^ (t1 * m)) / F) * Exp(-r * T) * (F * CND(d1) - X * CND(d2))
    ElseIf CallPutFlag = "p" Then  'Receiver swaption
        Swaption = ((1 - 1 / (1 + F / m) ^ (t1 * m)) / F) * Exp(-r * T) * (X * CND(-d2) - F * CND(-d1))
    End If

End Function

Public Function ESwaption(OutputFlag As String, CallPutFlag As String, t1 As Double, m As Double, F As Double, X As Double, T As Double, _
                r As Double, v As Double) As Double
                
                If OutputFlag = "p" Then
                    ESwaption = Swaption(CallPutFlag, t1, m, F, X, T, r, v)
                ElseIf OutputFlag = "d" Then
                    ESwaption = (Swaption(CallPutFlag, t1, m, F + 0.001, X, T, r, v) - Swaption(CallPutFlag, t1, m, F - 0.001, X, T, r, v)) / (20)
                ElseIf OutputFlag = "v" Then
                    ESwaption = (Swaption(CallPutFlag, t1, m, F, X, T, r, v + 0.01) - Swaption(CallPutFlag, t1, m, F, X, T, r, v - 0.01)) / (200 * 0.01)
                ElseIf OutputFlag = "t" Then
                    ESwaption = Swaption(CallPutFlag, t1, m, F, X, T - 1 / 365, r, v) - Swaption(CallPutFlag, t1, m, F, X, T, r, v)
                  ElseIf OutputFlag = "r" Then
                    ESwaption = (Swaption(CallPutFlag, t1, m, F, X, T, r + 0.01, v) - Swaption(CallPutFlag, t1, m, F, X, T, r - 0.01, v)) / (200 * 0.01)
                End If
                
                
End Function


'//  Black-76 European modified for Options on Money Market Futures
'// This is simply call on price is put on yield and put on price is call on yield
Public Function MoneyMarketFuturesOption(CallPutFlag As String, F As Double, X As Double, T As Double, _
                r As Double, v As Double) As Double
 
    Dim d1 As Double, d2 As Double
    
    F = 100 - F
    X = 100 - X
    
    d1 = (Log(F / X) + v ^ 2 / 2 * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    If CallPutFlag = "c" Then 'Call on price is put on implied money market yield
        MoneyMarketFuturesOption = Exp(-r * T) * (X * CND(-d2) - F * CND(-d1))
    ElseIf CallPutFlag = "p" Then  'Put on price is call on implied money market yield
        MoneyMarketFuturesOption = Exp(-r * T) * (F * CND(d1) - X * CND(d2))
    End If

End Function


'// Vasicek: options on zero coupon bonds
Function VasicekBondOption(CallPutFlag As String, F As Double, X As Double, tau As Double, T As Double, _
        r As Double, theta As Double, kappa As Double, v As Double) As Double

  
    Dim PtT As Double, Pt_tau As Double
    Dim h As Double, vp As Double

    X = X / F
    PtT = VasicekBondPrice(0, T, r, theta, kappa, v)
    Pt_tau = VasicekBondPrice(0, tau, r, theta, kappa, v)
    vp = Sqr(v ^ 2 * (1 - Exp(-2 * kappa * T)) / (2 * kappa)) * (1 - Exp(-kappa * (tau - T))) / kappa
   
    h = 1 / vp * Log(Pt_tau / (PtT * X)) + vp / 2
    
    If CallPutFlag = "c" Then
        VasicekBondOption = F * (Pt_tau * CND(h) - X * PtT * CND(h - vp))
    Else
        VasicekBondOption = F * (X * PtT * CND(-h + vp) - Pt_tau * CND(-h))
    End If
End Function


'// Vasicek: value zero coupon bond
Public Function VasicekBondPrice(t1 As Double, T As Double, r As Double, theta As Double, kappa As Double, v As Double) As Double
    Dim BtT As Double, AtT As Double, PtT As Double

    BtT = (1 - Exp(-kappa * (T - t1))) / kappa
    AtT = Exp((BtT - T + t1) * (kappa ^ 2 * theta - v ^ 2 / 2) / kappa ^ 2 - v ^ 2 * BtT ^ 2 / (4 * kappa))
    PtT = AtT * Exp(-BtT * r)
    VasicekBondPrice = PtT
End Function


Public Function MMFutures(F As Double, t1 As Double, v As Double, Basis As Double, tau As Double, kappa As Double) As Double

    Dim Z As Double, T2 As Double
    
    T2 = t1 + tau / 365
    
    If kappa = 0 Then
        Z = v ^ 2 * t1 * (T2 - t1) ^ 2 + v ^ 2 / 2 * t1 ^ 2 * (T2 - t1)
    Else
        Z = v ^ 2 * (1 - Exp(-2 * kappa * t1)) / (2 * kappa) * ((1 - Exp(-kappa * (T2 - t1))) / kappa) ^ 2 _
            + v ^ 2 / (2 * kappa ^ 3) * (1 - Exp(-kappa * (T2 - t1))) * (1 - Exp(-kappa * t1)) ^ 2
    End If
        
        MMFutures = F + (1 - Exp(-Z)) * (100 - F + 100 * Basis / tau)

End Function
                                                                                                         
                                                                                                         
Public Function BDTYieldOnly(ReturnFlag As String, v As Double, N As Integer, T As Double, InputZeroRates As Variant, YieldMatu As Variant)
         
    Dim ZeroR() As Double, ZeroBond() As Double, U() As Double
    Dim r() As Double, Lambda() As Double, Df() As Double
    Dim dt As Double, epsilon As Double
    Dim Pi As Double, di As Double
    Dim i As Integer, j As Integer, m As Integer
    
    ReDim ZeroR(0 To N + 1)
    ReDim r(0 To N * 2, 0 To N * 2)
    ReDim Lambda(0 To N * 2, 0 To N * 2)
    ReDim ZeroBond(0 To N + 1)
    ReDim U(0 To N)
    ReDim Df(0 To N * 2, 0 To N * 2)
    
    dt = T / (N + 1)
    
    For i = 1 To N + 1
        ZeroR(i) = InputZeroRates(i)
        ZeroBond(i) = 1 / (1 + ZeroR(i) * dt) ^ (i)
    Next
       
    Lambda(0, 0) = 1
    U(0) = ZeroR(1)
    r(0, 0) = ZeroR(1)
    Df(0, 0) = 1 / (1 + r(0, 0) * dt)

    For i = 1 To N
        '// Calculate the Arrow-Debreu prices by forward induction:
        Lambda(i, 0) = 0.5 * Lambda(i - 1, 0) * Df(i - 1, 0) ' //Arrow-Debreu  at lowest node
        Lambda(i, i) = 0.5 * Lambda(i - 1, i - 1) * Df(i - 1, i - 1) ' //Arrow-Debreu  at uppest node
        For j = 1 To i - 1 ' //Arrow-Debreua  between lowest and uppest node
            Lambda(i, j) = 0.5 * Lambda(i - 1, j - 1) * Df(i - 1, j - 1) + 0.5 * Lambda(i - 1, j) * Df(i - 1, j)
        Next
        
        '// Newton-Raphson method to find the unknown median u(i)
        U(i) = U(i - 1)     '// Seed value
        di = 0
        Pi = 0
        For j = 0 To i
            m = j * 2 - i
            Pi = Pi + Lambda(i, j) / (1 + U(i) * Exp(v * m * Sqr(dt)) * dt)
            di = di - Lambda(i, j) * (Exp(v * m * Sqr(dt)) * dt) / (1 + U(i) * Exp(v * m * Sqr(dt)) * dt) ^ 2
       Next
        epsilon = 0.000000001
        While Abs(Pi - ZeroBond(i + 1)) > epsilon
            U(i) = U(i) - (Pi - ZeroBond(i + 1)) / di
            di = 0
            Pi = 0
            For j = 0 To i
                m = -i + j * 2
                Pi = Pi + Lambda(i, j) / (1 + U(i) * Exp(v * m * Sqr(dt)) * dt)
                di = di - Lambda(i, j) * (Exp(v * m * Sqr(dt)) * dt) / (1 + U(i) * Exp(v * m * Sqr(dt)) * dt) ^ 2
            Next
        Wend
   
        '// Given u(i) from the search above we can calculate the short rates and
        '// the corresponding discount factors
        For j = 0 To i
            m = (-i + j * 2)
            r(i, j) = U(i) * Exp(v * m * Sqr(dt))
            Df(i, j) = 1 / (1 + r(i, j) * dt)
        Next j
        
    Next i
    
    '// Output
    If ReturnFlag = "r" Then '// Will return the short rate tree as a matrix
        BDTYieldOnly = Application.Transpose(r())
    ElseIf ReturnFlag = "d" Then '// Will return the discount factor tree as a matrix
        BDTYieldOnly = Application.Transpose(Df())
    ElseIf ReturnFlag = "a" Then '// Will return Arrow-Debreu tree as a matrix
        BDTYieldOnly = Application.Transpose(Lambda())
    End If
    
End Function
