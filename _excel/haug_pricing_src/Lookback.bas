Attribute VB_Name = "Lookback"
Option Explicit

' Programmer Espen Gaarder Haug
' Copyright Espen Gaarder Haug  2006


'// Partial-time fixed strike lookback options
Public Function PartialFixedLB(CallPutFlag As String, S As Double, X As Double, t1 As Double, _
                T2 As Double, r As Double, b As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double
    Dim e1 As Double, e2 As Double
    Dim f1 As Double, f2 As Double

    d1 = (Log(S / X) + (b + v ^ 2 / 2) * T2) / (v * Sqr(T2))
    d2 = d1 - v * Sqr(T2)
    e1 = ((b + v ^ 2 / 2) * (T2 - t1)) / (v * Sqr(T2 - t1))
    e2 = e1 - v * Sqr(T2 - t1)
    f1 = (Log(S / X) + (b + v ^ 2 / 2) * t1) / (v * Sqr(t1))
    f2 = f1 - v * Sqr(t1)
    If CallPutFlag = "c" Then
        PartialFixedLB = S * Exp((b - r) * T2) * CND(d1) - Exp(-r * T2) * X * CND(d2) + S * Exp(-r * T2) * v ^ 2 / (2 * b) * (-(S / X) ^ (-2 * b / v ^ 2) * CBND(d1 - 2 * b * Sqr(T2) / v, -f1 + 2 * b * Sqr(t1) / v, -Sqr(t1 / T2)) + Exp(b * T2) * CBND(e1, d1, Sqr(1 - t1 / T2))) - S * Exp((b - r) * T2) * CBND(-e1, d1, -Sqr(1 - t1 / T2)) - X * Exp(-r * T2) * CBND(f2, -d2, -Sqr(t1 / T2)) + Exp(-b * (T2 - t1)) * (1 - v ^ 2 / (2 * b)) * S * Exp((b - r) * T2) * CND(f1) * CND(-e2)
    ElseIf CallPutFlag = "p" Then
        PartialFixedLB = X * Exp(-r * T2) * CND(-d2) - S * Exp((b - r) * T2) * CND(-d1) + S * Exp(-r * T2) * v ^ 2 / (2 * b) * ((S / X) ^ (-2 * b / v ^ 2) * CBND(-d1 + 2 * b * Sqr(T2) / v, f1 - 2 * b * Sqr(t1) / v, -Sqr(t1 / T2)) - Exp(b * T2) * CBND(-e1, -d1, Sqr(1 - t1 / T2))) + S * Exp((b - r) * T2) * CBND(e1, -d1, -Sqr(1 - t1 / T2)) + X * Exp(-r * T2) * CBND(-f2, d2, -Sqr(t1 / T2)) - Exp(-b * (T2 - t1)) * (1 - v ^ 2 / (2 * b)) * S * Exp((b - r) * T2) * CND(-f1) * CND(e2)
    End If
End Function
