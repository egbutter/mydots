Attribute VB_Name = "Analytics"
Option Explicit

' Programmer Espen Gaarder Haug
' Copyright 2006 Espen GaarderHaug


Function CND(X As Double) As Double
    Dim y As Double, Exponential As Double, SumA As Double, SumB As Double
    
    y = Abs(X)
    If y > 37 Then
        CND = 0
    Else
        Exponential = Exp(-y ^ 2 / 2)
        If y < 7.07106781186547 Then
            SumA = 3.52624965998911E-02 * y + 0.700383064443688
            SumA = SumA * y + 6.37396220353165
            SumA = SumA * y + 33.912866078383
            SumA = SumA * y + 112.079291497871
            SumA = SumA * y + 221.213596169931
            SumA = SumA * y + 220.206867912376
            SumB = 8.83883476483184E-02 * y + 1.75566716318264
            SumB = SumB * y + 16.064177579207
            SumB = SumB * y + 86.7807322029461
            SumB = SumB * y + 296.564248779674
            SumB = SumB * y + 637.333633378831
            SumB = SumB * y + 793.826512519948
            SumB = SumB * y + 440.413735824752
            CND = Exponential * SumA / SumB
        Else
            SumA = y + 0.65
            SumA = y + 4 / SumA
            SumA = y + 3 / SumA
            SumA = y + 2 / SumA
            SumA = y + 1 / SumA
            CND = Exponential / (SumA * 2.506628274631)
        End If
  End If
  
  If X > 0 Then CND = 1 - CND

End Function


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

Public Function StandardBarrier(TypeFlag As String, S As Double, X As Double, H As Double, K As Double, T As Double, _
            r As Double, b As Double, v As Double)

    'TypeFlag:      The "TypeFlag" gives you 8 different standard barrier options:
    '               1) "cdi"=Down-and-in call,    2) "cui"=Up-and-in call
    '               3) "pdi"=Down-and-in put,     4) "pui"=Up-and-in put
    '               5) "cdo"=Down-and-out call,   6) "cuo"=Up-out-in call
    '               7) "pdo"=Down-and-out put,    8) "puo"=Up-out-in put
    
    Dim mu As Double
    Dim lambda As Double
    Dim X1 As Double, X2 As Double
    Dim y1 As Double, y2 As Double
    Dim z As Double
    
    Dim eta As Integer    'Binary variable that can take the value of 1 or -1
    Dim phi As Integer    'Binary variable that can take the value of 1 or -1
    
    Dim f1 As Double    'Equal to formula "A" in the book
    Dim f2 As Double    'Equal to formula "B" in the book
    Dim f3 As Double    'Equal to formula "C" in the book
    Dim f4 As Double    'Equal to formula "D" in the book
    Dim f5 As Double    'Equal to formula "E" in the book
    Dim f6 As Double    'Equal to formula "F" in the book

    mu = (b - v ^ 2 / 2) / v ^ 2
    lambda = Sqr(mu ^ 2 + 2 * r / v ^ 2)
    X1 = Log(S / X) / (v * Sqr(T)) + (1 + mu) * v * Sqr(T)
    X2 = Log(S / H) / (v * Sqr(T)) + (1 + mu) * v * Sqr(T)
    y1 = Log(H ^ 2 / (S * X)) / (v * Sqr(T)) + (1 + mu) * v * Sqr(T)
    y2 = Log(H / S) / (v * Sqr(T)) + (1 + mu) * v * Sqr(T)
    z = Log(H / S) / (v * Sqr(T)) + lambda * v * Sqr(T)
    
    If TypeFlag = "cdi" Or TypeFlag = "cdo" Then
        eta = 1
        phi = 1
    ElseIf TypeFlag = "cui" Or TypeFlag = "cuo" Then
        eta = -1
        phi = 1
    ElseIf TypeFlag = "pdi" Or TypeFlag = "pdo" Then
        eta = 1
        phi = -1
    ElseIf TypeFlag = "pui" Or TypeFlag = "puo" Then
        eta = -1
        phi = -1
    End If
    
    f1 = phi * S * Exp((b - r) * T) * CND(phi * X1) - phi * X * Exp(-r * T) * CND(phi * X1 - phi * v * Sqr(T))
    f2 = phi * S * Exp((b - r) * T) * CND(phi * X2) - phi * X * Exp(-r * T) * CND(phi * X2 - phi * v * Sqr(T))
    f3 = phi * S * Exp((b - r) * T) * (H / S) ^ (2 * (mu + 1)) * CND(eta * y1) - phi * X * Exp(-r * T) * (H / S) ^ (2 * mu) * CND(eta * y1 - eta * v * Sqr(T))
    f4 = phi * S * Exp((b - r) * T) * (H / S) ^ (2 * (mu + 1)) * CND(eta * y2) - phi * X * Exp(-r * T) * (H / S) ^ (2 * mu) * CND(eta * y2 - eta * v * Sqr(T))
    f5 = K * Exp(-r * T) * (CND(eta * X2 - eta * v * Sqr(T)) - (H / S) ^ (2 * mu) * CND(eta * y2 - eta * v * Sqr(T)))
    f6 = K * ((H / S) ^ (mu + lambda) * CND(eta * z) + (H / S) ^ (mu - lambda) * CND(eta * z - 2 * eta * lambda * v * Sqr(T)))
    
    
    If X > H Then
        Select Case TypeFlag
            Case Is = "cdi"      '1a) cdi
                StandardBarrier = f3 + f5
            Case Is = "cui"   '2a) cui
                StandardBarrier = f1 + f5
            Case Is = "pdi"    '3a) pdi
                StandardBarrier = f2 - f3 + f4 + f5
            Case Is = "pui" '4a) pui
                StandardBarrier = f1 - f2 + f4 + f5
            Case Is = "cdo"    '5a) cdo
                StandardBarrier = f1 - f3 + f6
            Case Is = "cuo"   '6a) cuo
                StandardBarrier = f6
            Case Is = "pdo"   '7a) pdo
                StandardBarrier = f1 - f2 + f3 - f4 + f6
            Case Is = "puo" '8a) puo
                StandardBarrier = f2 - f4 + f6
            End Select
    ElseIf X < H Then
        Select Case TypeFlag
            Case Is = "cdi" '1b) cdi
                StandardBarrier = f1 - f2 + f4 + f5
            Case Is = "cui"  '2b) cui
                StandardBarrier = f2 - f3 + f4 + f5
            Case Is = "pdi" '3b) pdi
                StandardBarrier = f1 + f5
            Case Is = "pui"   '4b) pui
                StandardBarrier = f3 + f5
            Case Is = "cdo" '5b) cdo
                StandardBarrier = f2 + f6 - f4
            Case Is = "cuo" '6b) cuo
                StandardBarrier = f1 - f2 + f3 - f4 + f6
            Case Is = "pdo"   '7b) pdo
                StandardBarrier = f6
            Case Is = "puo"  '8b) puo
                StandardBarrier = f1 - f3 + f6
        End Select
    End If
End Function
