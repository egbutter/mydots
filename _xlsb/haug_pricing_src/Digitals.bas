Attribute VB_Name = "Digitals"
Option Explicit

' Implementation By Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug

Global Const Pi = 3.14159265358979

'// Discrete barrier monitoring adjustment
Public Function DiscreteAdjustedBarrier(S As Double, h As Double, v As Double, dt As Double) As Double

    If h > S Then
        DiscreteAdjustedBarrier = h * Exp(0.5826 * v * Sqr(dt))
    ElseIf h < S Then
        DiscreteAdjustedBarrier = h * Exp(-0.5826 * v * Sqr(dt))
    End If
End Function

Public Function EBinaryBarrier(OutPutFlag As String, TypeFlag As Integer, S As Double, x As Double, h As Double, k As Double, _
                T As Double, r As Double, b As Double, v As Double, eta As Integer, phi As Integer, Optional dS) As Double

            
    If IsMissing(dS) Then
        dS = 0.01
    End If
   
    If S >= h And TypeFlag / 2 - Int(TypeFlag / 2) = 0 Then
        If OutPutFlag = "p" Then
            If TypeFlag = 2 Then
                EBinaryBarrier = k
            ElseIf TypeFlag = 4 Then
                EBinaryBarrier = S
            ElseIf TypeFlag = 6 Then
                EBinaryBarrier = Exp(-r * T) * k
            ElseIf TypeFlag = 8 Then
                EBinaryBarrier = Exp(-r * T) * S
            ElseIf TypeFlag = 10 Then
                EBinaryBarrier = 0
            ElseIf TypeFlag = 12 Then
                EBinaryBarrier = 0
            ElseIf TypeFlag = 14 Then
                If S >= x Then
                    EBinaryBarrier = Exp(-r * T) * k
                Else
                    EBinaryBarrier = 0
                End If
            ElseIf TypeFlag = 16 Then
                If S >= x Then
                    EBinaryBarrier = Exp(-r * T) * S
                Else
                    EBinaryBarrier = 0
                End If
            ElseIf TypeFlag = 18 Then
                If S <= x Then
                    EBinaryBarrier = Exp(-r * T) * k
                Else
                    EBinaryBarrier = 0
                End If
            ElseIf TypeFlag = 20 Then
                If S <= x Then
                    EBinaryBarrier = Exp(-r * T) * S
                Else
                    EBinaryBarrier = 0
                End If
            ElseIf TypeFlag = 22 Then
                EBinaryBarrier = 0
            ElseIf TypeFlag = 24 Then
                EBinaryBarrier = 0
            ElseIf TypeFlag = 26 Then
                EBinaryBarrier = 0
            ElseIf TypeFlag = 28 Then
                EBinaryBarrier = 0
            End If
        ElseIf OutPutFlag = "d" Then
            If TypeFlag = 4 Then
                EBinaryBarrier = 1
            ElseIf TypeFlag = 8 Or TypeFlag = 16 Or TypeFlag = 20 Then
                EBinaryBarrier = Exp((b - r) * T) * 1
            End If
        Else
            EBinaryBarrier = 0
        End If
            Exit Function
    End If
    
    If S <= h And TypeFlag / 2 - Int(TypeFlag / 2) <> 0 Then
        If OutPutFlag = "p" Then
            If TypeFlag = 1 Then
                EBinaryBarrier = k
            ElseIf TypeFlag = 3 Then
                EBinaryBarrier = S
            ElseIf TypeFlag = 5 Then
                EBinaryBarrier = Exp(-r * T) * k
            ElseIf TypeFlag = 7 Then
                EBinaryBarrier = Exp(-r * T) * S
            ElseIf TypeFlag = 9 Then
                EBinaryBarrier = 0
            ElseIf TypeFlag = 11 Then
                EBinaryBarrier = 0
            ElseIf TypeFlag = 13 Then
                If S >= x Then
                    EBinaryBarrier = Exp(-r * T) * k
                Else
                    EBinaryBarrier = 0
                End If
            ElseIf TypeFlag = 15 Then
                If S >= x Then
                    EBinaryBarrier = Exp(-r * T) * S
                Else
                    EBinaryBarrier = 0
                End If
            ElseIf TypeFlag = 17 Then
                If S <= x Then
                    EBinaryBarrier = Exp(-r * T) * k
                Else
                    EBinaryBarrier = 0
                End If
            ElseIf TypeFlag = 19 Then
                If S <= x Then
                    EBinaryBarrier = Exp(-r * T) * S
                Else
                    EBinaryBarrier = 0
                End If
            ElseIf TypeFlag = 21 Then
                EBinaryBarrier = 0
            ElseIf TypeFlag = 23 Then
                EBinaryBarrier = 0
            ElseIf TypeFlag = 25 Then
                EBinaryBarrier = 0
            ElseIf TypeFlag = 27 Then
                EBinaryBarrier = 0
            End If
        ElseIf OutPutFlag = "d" Then
            If TypeFlag = 3 Then
                EBinaryBarrier = 1
            ElseIf TypeFlag = 7 Or TypeFlag = 15 Or TypeFlag = 19 Then
                EBinaryBarrier = Exp((b - r) * T) * 1
            End If
        Else
            EBinaryBarrier = 0
        End If
        Exit Function
    End If
    
    
    If OutPutFlag = "p" Then ' Value
        EBinaryBarrier = BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v, eta, phi)
    ElseIf OutPutFlag = "d" Then 'Delta
         EBinaryBarrier = (BinaryBarrier(TypeFlag, S + dS, x, h, k, T, r, b, v, eta, phi) - BinaryBarrier(TypeFlag, S - dS, x, h, k, T, r, b, v, eta, phi)) / (2 * dS)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EBinaryBarrier = (BinaryBarrier(TypeFlag, S + dS, x, h, k, T, r, b, v, eta, phi) - 2 * BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v, eta, phi) + BinaryBarrier(TypeFlag, S - dS, x, h, k, T, r, b, v, eta, phi)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EBinaryBarrier = (BinaryBarrier(TypeFlag, S + dS, x, h, k, T, r, b, v + 0.01, eta, phi) - 2 * BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v + 0.01, eta, phi) + BinaryBarrier(TypeFlag, S - dS, x, h, k, T, r, b, v + 0.01, eta, phi) _
      - BinaryBarrier(TypeFlag, S + dS, x, h, k, T, r, b, v - 0.01, eta, phi) + 2 * BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v - 0.01, eta, phi) - BinaryBarrier(TypeFlag, S - dS, x, h, k, T, r, b, v - 0.01, eta, phi)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EBinaryBarrier = 1 / (4 * dS * 0.01) * (BinaryBarrier(TypeFlag, S + dS, x, h, k, T, r, b, v + 0.01, eta, phi) - BinaryBarrier(TypeFlag, S + dS, x, h, k, T, r, b, v - 0.01, eta, phi) _
        - BinaryBarrier(TypeFlag, S - dS, x, h, k, T, r, b, v + 0.01, eta, phi) + BinaryBarrier(TypeFlag, S - dS, x, h, k, T, r, b, v - 0.01, eta, phi)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EBinaryBarrier = (BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v + 0.01, eta, phi) - BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v - 0.01, eta, phi)) / 2
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EBinaryBarrier = v / 0.1 * (BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v + 0.01, eta, phi) - BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v - 0.01, eta, phi)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EBinaryBarrier = (BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v + 0.01, eta, phi) - 2 * BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v, eta, phi) + BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v - 0.01, eta, phi))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
            EBinaryBarrier = BinaryBarrier(TypeFlag, S, x, h, k, 0.00001, r, b, v, eta, phi) - BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v, eta, phi)
        Else
                EBinaryBarrier = BinaryBarrier(TypeFlag, S, x, h, k, T - 1 / 365, r, b, v, eta, phi) - BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v, eta, phi)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EBinaryBarrier = (BinaryBarrier(TypeFlag, S, x, h, k, T, r + 0.01, b + 0.01, v, eta, phi) - BinaryBarrier(TypeFlag, S, x, h, k, T, r - 0.01, b - 0.01, v, eta, phi)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures Rho
         EBinaryBarrier = (BinaryBarrier(TypeFlag, S, x, h, k, T, r + 0.01, b, v, eta, phi) - BinaryBarrier(TypeFlag, S, x, h, k, T, r - 0.01, b, v, eta, phi)) / (2)
    ElseIf OutPutFlag = "f" Then 'Rho2
         EBinaryBarrier = (BinaryBarrier(TypeFlag, S, x, h, k, T, r, b - 0.01, v, eta, phi) - BinaryBarrier(TypeFlag, S, x, h, k, T, r, b + 0.01, v, eta, phi)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EBinaryBarrier = (BinaryBarrier(TypeFlag, S, x, h, k, T, r, b + 0.01, v, eta, phi) - BinaryBarrier(TypeFlag, S, x, h, k, T, r, b - 0.01, v, eta, phi)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EBinaryBarrier = 1 / dS ^ 3 * (BinaryBarrier(TypeFlag, S + 2 * dS, x, h, k, T, r, b, v, eta, phi) - 3 * BinaryBarrier(TypeFlag, S + dS, x, h, k, T, r, b, v, eta, phi) _
                                + 3 * BinaryBarrier(TypeFlag, S, x, h, k, T, r, b, v, eta, phi) - BinaryBarrier(TypeFlag, S - dS, x, h, k, T, r, b, v, eta, phi))
    End If
    
End Function


'// Binary barrier options
Public Function BinaryBarrier(TypeFlag As Integer, S As Double, x As Double, h As Double, k As Double, _
                T As Double, r As Double, b As Double, v As Double, eta As Integer, phi As Integer) As Double

    '// TypeFlag:  Value 1 to 28 dependent on binary option type,
    '//            look in the book for spesifications.
    
    Dim X1 As Double, X2 As Double
    Dim y1 As Double, y2 As Double
    Dim Z As Double, mu As Double, lambda As Double
    Dim a1 As Double, a2 As Double, a3 As Double, a4 As Double, a5 As Double
    Dim b1 As Double, b2 As Double, b3 As Double, b4 As Double

    mu = (b - v ^ 2 / 2) / v ^ 2
    lambda = Sqr(mu ^ 2 + 2 * r / v ^ 2)
    X1 = Log(S / x) / (v * Sqr(T)) + (mu + 1) * v * Sqr(T)
    X2 = Log(S / h) / (v * Sqr(T)) + (mu + 1) * v * Sqr(T)
    y1 = Log(h ^ 2 / (S * x)) / (v * Sqr(T)) + (mu + 1) * v * Sqr(T)
    y2 = Log(h / S) / (v * Sqr(T)) + (mu + 1) * v * Sqr(T)
    Z = Log(h / S) / (v * Sqr(T)) + lambda * v * Sqr(T)
    
    a1 = S * Exp((b - r) * T) * CND(phi * X1)
    b1 = k * Exp(-r * T) * CND(phi * X1 - phi * v * Sqr(T))
    a2 = S * Exp((b - r) * T) * CND(phi * X2)
    b2 = k * Exp(-r * T) * CND(phi * X2 - phi * v * Sqr(T))
    a3 = S * Exp((b - r) * T) * (h / S) ^ (2 * (mu + 1)) * CND(eta * y1)
    b3 = k * Exp(-r * T) * (h / S) ^ (2 * mu) * CND(eta * y1 - eta * v * Sqr(T))
    a4 = S * Exp((b - r) * T) * (h / S) ^ (2 * (mu + 1)) * CND(eta * y2)
    b4 = k * Exp(-r * T) * (h / S) ^ (2 * mu) * CND(eta * y2 - eta * v * Sqr(T))
    a5 = k * ((h / S) ^ (mu + lambda) * CND(eta * Z) + (h / S) ^ (mu - lambda) * CND(eta * Z - 2 * eta * lambda * v * Sqr(T)))
    
    If x > h Then
        Select Case TypeFlag
            Case Is < 5
                BinaryBarrier = a5
            Case Is < 7
                BinaryBarrier = b2 + b4
            Case Is < 9
                BinaryBarrier = a2 + a4
            Case Is < 11
                BinaryBarrier = b2 - b4
            Case Is < 13
                BinaryBarrier = a2 - a4
            Case Is = 13
                BinaryBarrier = b3
            Case Is = 14
                BinaryBarrier = b3
            Case Is = 15
                BinaryBarrier = a3
            Case Is = 16
                BinaryBarrier = a1
            Case Is = 17
                BinaryBarrier = b2 - b3 + b4
            Case Is = 18
                BinaryBarrier = b1 - b2 + b4
            Case Is = 19
                BinaryBarrier = a2 - a3 + a4
            Case Is = 20
                BinaryBarrier = a1 - a2 + a3
            Case Is = 21
                BinaryBarrier = b1 - b3
            Case Is = 22
                BinaryBarrier = 0
            Case Is = 23
                BinaryBarrier = a1 - a3
            Case Is = 24
               BinaryBarrier = 0
            Case Is = 25
                BinaryBarrier = b1 - b2 + b3 - b4
            Case Is = 26
                BinaryBarrier = b2 - b4
            Case Is = 27
                BinaryBarrier = a1 - a2 + a3 - a4
            Case Is = 28
                BinaryBarrier = a2 - a4
        End Select
    ElseIf x < h Then
        Select Case TypeFlag
            Case Is < 5
                BinaryBarrier = a5
            Case Is < 7
                BinaryBarrier = b2 + b4
            Case Is < 9
                BinaryBarrier = a2 + a4
            Case Is < 11
                BinaryBarrier = b2 - b4
            Case Is < 13
                BinaryBarrier = a2 - a4
            Case Is = 13
                BinaryBarrier = b1 - b2 + b4
            Case Is = 14
                BinaryBarrier = b2 - b3 + b4
            Case Is = 15
                BinaryBarrier = a1 - a2 + a4
            Case Is = 16
                BinaryBarrier = a2 - a3 + a4
            Case Is = 17
                BinaryBarrier = b1
            Case Is = 18
                BinaryBarrier = b3
            Case Is = 19
                BinaryBarrier = a1
            Case Is = 20
                BinaryBarrier = a3
            Case Is = 21
                BinaryBarrier = b2 - b4
            Case Is = 22
                BinaryBarrier = b1 - b2 + b3 - b4
            Case Is = 23
                BinaryBarrier = a2 - a4
            Case Is = 24
                BinaryBarrier = a1 - a2 + a3 - a4
            Case Is = 25
                BinaryBarrier = 0
            Case Is = 26
                BinaryBarrier = b1 - b3
            Case Is = 27
                BinaryBarrier = 0
            Case Is = 28
                BinaryBarrier = a1 - a3
        End Select
    End If
End Function



Public Function ECashOrNothing(OutPutFlag As String, CallPutFlag As String, S As Double, x As Double, k As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        ECashOrNothing = CashOrNothing(CallPutFlag, S, x, k, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         ECashOrNothing = (CashOrNothing(CallPutFlag, S + dS, x, k, T, r, b, v) - CashOrNothing(CallPutFlag, S - dS, x, k, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ECashOrNothing = (CashOrNothing(CallPutFlag, S + dS, x, k, T, r, b, v) - 2 * CashOrNothing(CallPutFlag, S, x, k, T, r, b, v) + CashOrNothing(CallPutFlag, S - dS, x, k, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ECashOrNothing = (CashOrNothing(CallPutFlag, S + dS, x, k, T, r, b, v + 0.01) - 2 * CashOrNothing(CallPutFlag, S, x, k, T, r, b, v + 0.01) + CashOrNothing(CallPutFlag, S - dS, x, k, T, r, b, v + 0.01) _
      - CashOrNothing(CallPutFlag, S + dS, x, k, T, r, b, v - 0.01) + 2 * CashOrNothing(CallPutFlag, S, x, k, T, r, b, v - 0.01) - CashOrNothing(CallPutFlag, S - dS, x, k, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ECashOrNothing = 1 / (4 * dS * 0.01) * (CashOrNothing(CallPutFlag, S + dS, x, k, T, r, b, v + 0.01) - CashOrNothing(CallPutFlag, S + dS, x, k, T, r, b, v - 0.01) _
        - CashOrNothing(CallPutFlag, S - dS, x, k, T, r, b, v + 0.01) + CashOrNothing(CallPutFlag, S - dS, x, k, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ECashOrNothing = (CashOrNothing(CallPutFlag, S, x, k, T, r, b, v + 0.01) - CashOrNothing(CallPutFlag, S, x, k, T, r, b, v - 0.01)) / 2
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ECashOrNothing = v / 0.1 * (CashOrNothing(CallPutFlag, S, x, k, T, r, b, v + 0.01) - CashOrNothing(CallPutFlag, S, x, k, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ECashOrNothing = (CashOrNothing(CallPutFlag, S, x, k, T, r, b, v + 0.01) - 2 * CashOrNothing(CallPutFlag, S, x, k, T, r, b, v) + CashOrNothing(CallPutFlag, S, x, k, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
            ECashOrNothing = CashOrNothing(CallPutFlag, S, x, k, 0.00001, r, b, v) - CashOrNothing(CallPutFlag, S, x, k, T, r, b, v)
        Else
                ECashOrNothing = CashOrNothing(CallPutFlag, S, x, k, T - 1 / 365, r, b, v) - CashOrNothing(CallPutFlag, S, x, k, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         ECashOrNothing = (CashOrNothing(CallPutFlag, S, x, k, T, r + 0.01, b + 0.01, v) - CashOrNothing(CallPutFlag, S, x, k, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures Rho
         ECashOrNothing = (CashOrNothing(CallPutFlag, S, x, k, T, r + 0.01, b, v) - CashOrNothing(CallPutFlag, S, x, k, T, r - 0.01, b, v)) / (2)
    ElseIf OutPutFlag = "f" Then 'Rho2
         ECashOrNothing = (CashOrNothing(CallPutFlag, S, x, k, T, r, b - 0.01, v) - CashOrNothing(CallPutFlag, S, x, k, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        ECashOrNothing = (CashOrNothing(CallPutFlag, S, x, k, T, r, b + 0.01, v) - CashOrNothing(CallPutFlag, S, x, k, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ECashOrNothing = 1 / dS ^ 3 * (CashOrNothing(CallPutFlag, S + 2 * dS, x, k, T, r, b, v) - 3 * CashOrNothing(CallPutFlag, S + dS, x, k, T, r, b, v) _
                                + 3 * CashOrNothing(CallPutFlag, S, x, k, T, r, b, v) - CashOrNothing(CallPutFlag, S - dS, x, k, T, r, b, v))
      ElseIf OutPutFlag = "dx" Then 'Strike Delta
         ECashOrNothing = (CashOrNothing(CallPutFlag, S, x + dS, k, T, r, b, v) - CashOrNothing(CallPutFlag, S, x - dS, k, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        ECashOrNothing = (CashOrNothing(CallPutFlag, S, x + dS, k, T, r, b, v) - 2 * CashOrNothing(CallPutFlag, S, x, k, T, r, b, v) + CashOrNothing(CallPutFlag, S, x - dS, k, T, r, b, v)) / dS ^ 2

    
    End If
    
End Function


Public Function EAssetOrNothing(OutPutFlag As String, CallPutFlag As String, S As Double, x As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        EAssetOrNothing = AssetOrNothing(CallPutFlag, S, x, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EAssetOrNothing = (AssetOrNothing(CallPutFlag, S + dS, x, T, r, b, v) - AssetOrNothing(CallPutFlag, S - dS, x, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EAssetOrNothing = (AssetOrNothing(CallPutFlag, S + dS, x, T, r, b, v) - 2 * AssetOrNothing(CallPutFlag, S, x, T, r, b, v) + AssetOrNothing(CallPutFlag, S - dS, x, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EAssetOrNothing = (AssetOrNothing(CallPutFlag, S + dS, x, T, r, b, v + 0.01) - 2 * AssetOrNothing(CallPutFlag, S, x, T, r, b, v + 0.01) + AssetOrNothing(CallPutFlag, S - dS, x, T, r, b, v + 0.01) _
      - AssetOrNothing(CallPutFlag, S + dS, x, T, r, b, v - 0.01) + 2 * AssetOrNothing(CallPutFlag, S, x, T, r, b, v - 0.01) - AssetOrNothing(CallPutFlag, S - dS, x, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EAssetOrNothing = 1 / (4 * dS * 0.01) * (AssetOrNothing(CallPutFlag, S + dS, x, T, r, b, v + 0.01) - AssetOrNothing(CallPutFlag, S + dS, x, T, r, b, v - 0.01) _
        - AssetOrNothing(CallPutFlag, S - dS, x, T, r, b, v + 0.01) + AssetOrNothing(CallPutFlag, S - dS, x, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EAssetOrNothing = (AssetOrNothing(CallPutFlag, S, x, T, r, b, v + 0.01) - AssetOrNothing(CallPutFlag, S, x, T, r, b, v - 0.01)) / 2
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EAssetOrNothing = v / 0.1 * (AssetOrNothing(CallPutFlag, S, x, T, r, b, v + 0.01) - AssetOrNothing(CallPutFlag, S, x, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EAssetOrNothing = (AssetOrNothing(CallPutFlag, S, x, T, r, b, v + 0.01) - 2 * AssetOrNothing(CallPutFlag, S, x, T, r, b, v) + AssetOrNothing(CallPutFlag, S, x, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
            EAssetOrNothing = AssetOrNothing(CallPutFlag, S, x, 0.00001, r, b, v) - AssetOrNothing(CallPutFlag, S, x, T, r, b, v)
        Else
                EAssetOrNothing = AssetOrNothing(CallPutFlag, S, x, T - 1 / 365, r, b, v) - AssetOrNothing(CallPutFlag, S, x, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EAssetOrNothing = (AssetOrNothing(CallPutFlag, S, x, T, r + 0.01, b + 0.01, v) - AssetOrNothing(CallPutFlag, S, x, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures Rho
         EAssetOrNothing = (AssetOrNothing(CallPutFlag, S, x, T, r + 0.01, b, v) - AssetOrNothing(CallPutFlag, S, x, T, r - 0.01, b, v)) / (2)
    ElseIf OutPutFlag = "f" Then 'Rho2
         EAssetOrNothing = (AssetOrNothing(CallPutFlag, S, x, T, r, b - 0.01, v) - AssetOrNothing(CallPutFlag, S, x, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EAssetOrNothing = (AssetOrNothing(CallPutFlag, S, x, T, r, b + 0.01, v) - AssetOrNothing(CallPutFlag, S, x, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EAssetOrNothing = 1 / dS ^ 3 * (AssetOrNothing(CallPutFlag, S + 2 * dS, x, T, r, b, v) - 3 * AssetOrNothing(CallPutFlag, S + dS, x, T, r, b, v) _
                                + 3 * AssetOrNothing(CallPutFlag, S, x, T, r, b, v) - AssetOrNothing(CallPutFlag, S - dS, x, T, r, b, v))
    ElseIf OutPutFlag = "dx" Then 'Strike Delta
         EAssetOrNothing = (AssetOrNothing(CallPutFlag, S, x + dS, T, r, b, v) - AssetOrNothing(CallPutFlag, S, x - dS, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "dxdx" Then 'Strike Gamma
        EAssetOrNothing = (AssetOrNothing(CallPutFlag, S, x + dS, T, r, b, v) - 2 * AssetOrNothing(CallPutFlag, S, x, T, r, b, v) + AssetOrNothing(CallPutFlag, S, x - dS, T, r, b, v)) / dS ^ 2

    End If
    
End Function


Public Function ESuperShare(OutPutFlag As String, S As Double, XL As Double, XH As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        ESuperShare = SuperShare(S, XL, XH, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         ESuperShare = (SuperShare(S + dS, XL, XH, T, r, b, v) - SuperShare(S - dS, XL, XH, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "g" Then 'Gamma
        ESuperShare = (SuperShare(S + dS, XL, XH, T, r, b, v) - 2 * SuperShare(S, XL, XH, T, r, b, v) + SuperShare(S - dS, XL, XH, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        ESuperShare = (SuperShare(S + dS, XL, XH, T, r, b, v + 0.01) - 2 * SuperShare(S, XL, XH, T, r, b, v + 0.01) + SuperShare(S - dS, XL, XH, T, r, b, v + 0.01) _
      - SuperShare(S + dS, XL, XH, T, r, b, v - 0.01) + 2 * SuperShare(S, XL, XH, T, r, b, v - 0.01) - SuperShare(S - dS, XL, XH, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        ESuperShare = 1 / (4 * dS * 0.01) * (SuperShare(S + dS, XL, XH, T, r, b, v + 0.01) - SuperShare(S + dS, XL, XH, T, r, b, v - 0.01) _
        - SuperShare(S - dS, XL, XH, T, r, b, v + 0.01) + SuperShare(S - dS, XL, XH, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         ESuperShare = (SuperShare(S, XL, XH, T, r, b, v + 0.01) - SuperShare(S, XL, XH, T, r, b, v - 0.01)) / 2
    ElseIf OutPutFlag = "vp" Then 'VegaP
         ESuperShare = v / 0.1 * (SuperShare(S, XL, XH, T, r, b, v + 0.01) - SuperShare(S, XL, XH, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        ESuperShare = (SuperShare(S, XL, XH, T, r, b, v + 0.01) - 2 * SuperShare(S, XL, XH, T, r, b, v) + SuperShare(S, XL, XH, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
            ESuperShare = SuperShare(S, XL, XH, 0.00001, r, b, v) - SuperShare(S, XL, XH, T, r, b, v)
        Else
                ESuperShare = SuperShare(S, XL, XH, T - 1 / 365, r, b, v) - SuperShare(S, XL, XH, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         ESuperShare = (SuperShare(S, XL, XH, T, r + 0.01, b + 0.01, v) - SuperShare(S, XL, XH, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures Rho
         ESuperShare = (SuperShare(S, XL, XH, T, r + 0.01, b, v) - SuperShare(S, XL, XH, T, r - 0.01, b, v)) / (2)
    ElseIf OutPutFlag = "f" Then 'Rho2
         ESuperShare = (SuperShare(S, XL, XH, T, r, b - 0.01, v) - SuperShare(S, XL, XH, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        ESuperShare = (SuperShare(S, XL, XH, T, r, b + 0.01, v) - SuperShare(S, XL, XH, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        ESuperShare = 1 / dS ^ 3 * (SuperShare(S + 2 * dS, XL, XH, T, r, b, v) - 3 * SuperShare(S + dS, XL, XH, T, r, b, v) _
                                + 3 * SuperShare(S, XL, XH, T, r, b, v) - SuperShare(S - dS, XL, XH, T, r, b, v))
    End If
    
End Function

Public Function EGapOption(OutPutFlag As String, CallPutFlag As String, S As Double, X1 As Double, X2 As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If OutPutFlag = "p" Then ' Value
        EGapOption = GapOption(CallPutFlag, S, X1, X2, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EGapOption = (GapOption(CallPutFlag, S + dS, X1, X2, T, r, b, v) - GapOption(CallPutFlag, S - dS, X1, X2, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EGapOption = (GapOption(CallPutFlag, S + dS, X1, X2, T, r, b, v) - 2 * GapOption(CallPutFlag, S, X1, X2, T, r, b, v) + GapOption(CallPutFlag, S - dS, X1, X2, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EGapOption = (GapOption(CallPutFlag, S + dS, X1, X2, T, r, b, v + 0.01) - 2 * GapOption(CallPutFlag, S, X1, X2, T, r, b, v + 0.01) + GapOption(CallPutFlag, S - dS, X1, X2, T, r, b, v + 0.01) _
      - GapOption(CallPutFlag, S + dS, X1, X2, T, r, b, v - 0.01) + 2 * GapOption(CallPutFlag, S, X1, X2, T, r, b, v - 0.01) - GapOption(CallPutFlag, S - dS, X1, X2, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EGapOption = 1 / (4 * dS * 0.01) * (GapOption(CallPutFlag, S + dS, X1, X2, T, r, b, v + 0.01) - GapOption(CallPutFlag, S + dS, X1, X2, T, r, b, v - 0.01) _
        - GapOption(CallPutFlag, S - dS, X1, X2, T, r, b, v + 0.01) + GapOption(CallPutFlag, S - dS, X1, X2, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EGapOption = (GapOption(CallPutFlag, S, X1, X2, T, r, b, v + 0.01) - GapOption(CallPutFlag, S, X1, X2, T, r, b, v - 0.01)) / 2
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EGapOption = v / 0.1 * (GapOption(CallPutFlag, S, X1, X2, T, r, b, v + 0.01) - GapOption(CallPutFlag, S, X1, X2, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EGapOption = (GapOption(CallPutFlag, S, X1, X2, T, r, b, v + 0.01) - 2 * GapOption(CallPutFlag, S, X1, X2, T, r, b, v) + GapOption(CallPutFlag, S, X1, X2, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
            EGapOption = GapOption(CallPutFlag, S, X1, X2, 0.00001, r, b, v) - GapOption(CallPutFlag, S, X1, X2, T, r, b, v)
        Else
                EGapOption = GapOption(CallPutFlag, S, X1, X2, T - 1 / 365, r, b, v) - GapOption(CallPutFlag, S, X1, X2, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EGapOption = (GapOption(CallPutFlag, S, X1, X2, T, r + 0.01, b + 0.01, v) - GapOption(CallPutFlag, S, X1, X2, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures Rho
         EGapOption = (GapOption(CallPutFlag, S, X1, X2, T, r + 0.01, b, v) - GapOption(CallPutFlag, S, X1, X2, T, r - 0.01, b, v)) / (2)
    ElseIf OutPutFlag = "f" Then 'Rho2
         EGapOption = (GapOption(CallPutFlag, S, X1, X2, T, r, b - 0.01, v) - GapOption(CallPutFlag, S, X1, X2, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EGapOption = (GapOption(CallPutFlag, S, X1, X2, T, r, b + 0.01, v) - GapOption(CallPutFlag, S, X1, X2, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EGapOption = 1 / dS ^ 3 * (GapOption(CallPutFlag, S + 2 * dS, X1, X2, T, r, b, v) - 3 * GapOption(CallPutFlag, S + dS, X1, X2, T, r, b, v) _
                                + 3 * GapOption(CallPutFlag, S, X1, X2, T, r, b, v) - GapOption(CallPutFlag, S - dS, X1, X2, T, r, b, v))
    
    End If
    
End Function




'// Cash-or-nothing options
Public Function CashOrNothing(CallPutFlag As String, S As Double, x As Double, k As Double, T As Double, _
                r As Double, b As Double, v As Double) As Double

    Dim d As Double

    d = (Log(S / x) + (b - v ^ 2 / 2) * T) / (v * Sqr(T))

    If CallPutFlag = "c" Then
        CashOrNothing = k * Exp(-r * T) * CND(d)
    ElseIf CallPutFlag = "p" Then
        CashOrNothing = k * Exp(-r * T) * CND(-d)
    End If
End Function


'// Asset-or-nothing options
Public Function AssetOrNothing(CallPutFlag As String, S As Double, x As Double, T As Double, r As Double, _
                b As Double, v As Double) As Double

    Dim d As Double
    
    d = (Log(S / x) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    
    If CallPutFlag = "c" Then
        AssetOrNothing = S * Exp((b - r) * T) * CND(d)
    ElseIf CallPutFlag = "p" Then
        AssetOrNothing = S * Exp((b - r) * T) * CND(-d)
    End If
End Function

'// Supershare options
Public Function SuperShare(S As Double, XL As Double, XH As Double, T As Double, _
                r As Double, b As Double, v As Double) As Double
 
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S / XL) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = (Log(S / XH) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))

    SuperShare = S * Exp((b - r) * T) / XL * (CND(d1) - CND(d2))
End Function

'// Gap options
Public Function GapOption(CallPutFlag As String, S As Double, X1 As Double, X2 As Double, _
                T As Double, r As Double, b As Double, v As Double) As Double

    Dim d1 As Double, d2 As Double

    d1 = (Log(S / X1) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)
    
    If CallPutFlag = "c" Then
        GapOption = S * Exp((b - r) * T) * CND(d1) - X2 * Exp(-r * T) * CND(d2)
    ElseIf CallPutFlag = "p" Then
        GapOption = X2 * Exp(-r * T) * CND(-d2) - S * Exp((b - r) * T) * CND(-d1)
    End If
End Function


'Double-Barrier Binary Range Option
Public Function DoubleBarrierBinary(TypeFlag As String, S As Double, L As Double, U As Double, k As Double, T As Double, r As Double, b As Double, v As Double) As Double

    Dim Alfa As Double, Beta As Double
    Dim Z As Double, sum As Double
    Dim i As Integer
    
    Alfa = -0.5 * (2 * b / v ^ 2 - 1)
    Beta = -0.25 * (2 * b / v ^ 2 - 1) ^ 2 - 2 * r / v ^ 2
    
    Z = Log(U / L)
    sum = 0
    For i = 1 To 50
        sum = sum + 2 * Pi * i * k / Z ^ 2 _
        * (((S / L) ^ Alfa - (-1) ^ i * (S / U) ^ Alfa) / (Alfa ^ 2 + (i * Pi / Z) ^ 2)) _
        * Sin(i * Pi / Z * Log(S / L)) * Exp(-0.5 * ((i * Pi / Z) ^ 2 - Beta) * v ^ 2 * T)
    Next
    
    If TypeFlag = "o" Then '// Knock-out
        DoubleBarrierBinary = sum
    ElseIf TypeFlag = "i" Then '// Knock-in
        DoubleBarrierBinary = k * Exp(-r * T) - sum
    End If
End Function


Public Function EDoubleBarrierBinary(OutPutFlag As String, TypeFlag As String, S As Double, L As Double, U As Double, k As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If S >= U Or S <= L Then
        If TypeFlag = "i" And OutPutFlag = "p" Then
            EDoubleBarrierBinary = k * Exp(-r * T)
        Else
            EDoubleBarrierBinary = 0
        End If
        Exit Function
    End If
    
    If OutPutFlag = "p" Then ' Value
        EDoubleBarrierBinary = DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EDoubleBarrierBinary = (DoubleBarrierBinary(TypeFlag, S + dS, L, U, k, T, r, b, v) - DoubleBarrierBinary(TypeFlag, S - dS, L, U, k, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EDoubleBarrierBinary = (DoubleBarrierBinary(TypeFlag, S + dS, L, U, k, T, r, b, v) - 2 * DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v) + DoubleBarrierBinary(TypeFlag, S - dS, L, U, k, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EDoubleBarrierBinary = (DoubleBarrierBinary(TypeFlag, S + dS, L, U, k, T, r, b, v + 0.01) - 2 * DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v + 0.01) + DoubleBarrierBinary(TypeFlag, S - dS, L, U, k, T, r, b, v + 0.01) _
      - DoubleBarrierBinary(TypeFlag, S + dS, L, U, k, T, r, b, v - 0.01) + 2 * DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v - 0.01) - DoubleBarrierBinary(TypeFlag, S - dS, L, U, k, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EDoubleBarrierBinary = 1 / (4 * dS * 0.01) * (DoubleBarrierBinary(TypeFlag, S + dS, L, U, k, T, r, b, v + 0.01) - DoubleBarrierBinary(TypeFlag, S + dS, L, U, k, T, r, b, v - 0.01) _
        - DoubleBarrierBinary(TypeFlag, S - dS, L, U, k, T, r, b, v + 0.01) + DoubleBarrierBinary(TypeFlag, S - dS, L, U, k, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EDoubleBarrierBinary = (DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v + 0.01) - DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v - 0.01)) / 2
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EDoubleBarrierBinary = v / 0.1 * (DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v + 0.01) - DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EDoubleBarrierBinary = (DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v + 0.01) - 2 * DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v) + DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
            EDoubleBarrierBinary = DoubleBarrierBinary(TypeFlag, S, L, U, k, 0.00001, r, b, v) - DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v)
        Else
                EDoubleBarrierBinary = DoubleBarrierBinary(TypeFlag, S, L, U, k, T - 1 / 365, r, b, v) - DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EDoubleBarrierBinary = (DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r + 0.01, b + 0.01, v) - DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures Rho
         EDoubleBarrierBinary = (DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r + 0.01, b, v) - DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r - 0.01, b, v)) / (2)
    ElseIf OutPutFlag = "f" Then 'Rho2
         EDoubleBarrierBinary = (DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b - 0.01, v) - DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EDoubleBarrierBinary = (DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b + 0.01, v) - DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EDoubleBarrierBinary = 1 / dS ^ 3 * (DoubleBarrierBinary(TypeFlag, S + 2 * dS, L, U, k, T, r, b, v) - 3 * DoubleBarrierBinary(TypeFlag, S + dS, L, U, k, T, r, b, v) _
                                + 3 * DoubleBarrierBinary(TypeFlag, S, L, U, k, T, r, b, v) - DoubleBarrierBinary(TypeFlag, S - dS, L, U, k, T, r, b, v))
    End If
    
End Function



'Double-Barrier binary assymmetrical
Public Function DoubleBarrierBinaryAsymmetric(TypeFlag As Integer, S As Double, L As Double, U As Double, Rebate As Double, T As Double, r As Double, b As Double, v As Double) As Double
    ' // TypeFlag="1" gives cash imideately at down barrier hit, knock-out at upper barrier
    ' // TypeFlag="2" gives cash imideately at up barrier hit, knock-out at lower barrier
    
    Dim Alfa As Double, Beta As Double
    Dim Z As Double, sum As Double
    Dim i As Integer
    
    If TypeFlag = 2 Then
        DoubleBarrierBinaryAsymmetric = DoubleBarrierBinaryAsymmetric(1, S, U, L, Rebate, T, r, b, v)
        Exit Function
    End If
    
    Alfa = -0.5 * (2 * b / v ^ 2 - 1)
    Beta = -0.25 * (2 * b / v ^ 2 - 1) ^ 2 - 2 * r / v ^ 2
    Z = Log(U / L)
    
    sum = 0
    For i = 1 To 50
        sum = sum + 2 / (i * Pi) * ((Beta - (i * Pi / Z) ^ 2 * Exp(-0.5 * ((i * Pi / Z) ^ 2 - Beta) * v ^ 2 * T)) _
        / ((i * Pi / Z) ^ 2 - Beta)) * Sin(i * Pi / Z * Log(S / L))
    Next
    
    DoubleBarrierBinaryAsymmetric = Rebate * (S / L) ^ Alfa * (sum + (1 - Log(S / L) / Z))
  
End Function


Public Function EDoubleBarrierBinaryAsymmetric(OutPutFlag As String, TypeFlag As Integer, S As Double, L As Double, U As Double, k As Double, T As Double, _
                r As Double, b As Double, v As Double, Optional dS)
            
    If IsMissing(dS) Then
        dS = 0.01
    End If
    
    If S >= U Then
            EDoubleBarrierBinaryAsymmetric = 0
            Exit Function
    ElseIf S <= L And OutPutFlag = "p" Then
        EDoubleBarrierBinaryAsymmetric = k
        Exit Function
    ElseIf S <= L Then
        EDoubleBarrierBinaryAsymmetric = 0
        Exit Function
    End If
    
    If OutPutFlag = "p" Then ' Value
        EDoubleBarrierBinaryAsymmetric = DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v)
    ElseIf OutPutFlag = "d" Then 'Delta
         EDoubleBarrierBinaryAsymmetric = (DoubleBarrierBinaryAsymmetric(TypeFlag, S + dS, L, U, k, T, r, b, v) - DoubleBarrierBinaryAsymmetric(TypeFlag, S - dS, L, U, k, T, r, b, v)) / (2 * dS)
    ElseIf OutPutFlag = "g" Then 'Gamma
        EDoubleBarrierBinaryAsymmetric = (DoubleBarrierBinaryAsymmetric(TypeFlag, S + dS, L, U, k, T, r, b, v) - 2 * DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v) + DoubleBarrierBinaryAsymmetric(TypeFlag, S - dS, L, U, k, T, r, b, v)) / dS ^ 2
    ElseIf OutPutFlag = "gv" Then 'DGammaDVol
        EDoubleBarrierBinaryAsymmetric = (DoubleBarrierBinaryAsymmetric(TypeFlag, S + dS, L, U, k, T, r, b, v + 0.01) - 2 * DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v + 0.01) + DoubleBarrierBinaryAsymmetric(TypeFlag, S - dS, L, U, k, T, r, b, v + 0.01) _
      - DoubleBarrierBinaryAsymmetric(TypeFlag, S + dS, L, U, k, T, r, b, v - 0.01) + 2 * DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v - 0.01) - DoubleBarrierBinaryAsymmetric(TypeFlag, S - dS, L, U, k, T, r, b, v - 0.01)) / (2 * 0.01 * dS ^ 2) / 100
   ElseIf OutPutFlag = "dddv" Then 'DDeltaDvol
        EDoubleBarrierBinaryAsymmetric = 1 / (4 * dS * 0.01) * (DoubleBarrierBinaryAsymmetric(TypeFlag, S + dS, L, U, k, T, r, b, v + 0.01) - DoubleBarrierBinaryAsymmetric(TypeFlag, S + dS, L, U, k, T, r, b, v - 0.01) _
        - DoubleBarrierBinaryAsymmetric(TypeFlag, S - dS, L, U, k, T, r, b, v + 0.01) + DoubleBarrierBinaryAsymmetric(TypeFlag, S - dS, L, U, k, T, r, b, v - 0.01)) / 100
    ElseIf OutPutFlag = "v" Then 'Vega
         EDoubleBarrierBinaryAsymmetric = (DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v + 0.01) - DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v - 0.01)) / 2
    ElseIf OutPutFlag = "vp" Then 'VegaP
         EDoubleBarrierBinaryAsymmetric = v / 0.1 * (DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v + 0.01) - DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v - 0.01)) / 2
     ElseIf OutPutFlag = "dvdv" Then 'DvegaDvol
        EDoubleBarrierBinaryAsymmetric = (DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v + 0.01) - 2 * DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v) + DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v - 0.01))
    ElseIf OutPutFlag = "t" Then 'Theta
         If T <= 1 / 365 Then
            EDoubleBarrierBinaryAsymmetric = DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, 0.00001, r, b, v) - DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v)
        Else
                EDoubleBarrierBinaryAsymmetric = DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T - 1 / 365, r, b, v) - DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v)
        End If
     ElseIf OutPutFlag = "r" Then 'Rho
         EDoubleBarrierBinaryAsymmetric = (DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r + 0.01, b + 0.01, v) - DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r - 0.01, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "fr" Then 'Futures Rho
         EDoubleBarrierBinaryAsymmetric = (DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r + 0.01, b, v) - DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r - 0.01, b, v)) / (2)
    ElseIf OutPutFlag = "f" Then 'Rho2
         EDoubleBarrierBinaryAsymmetric = (DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b - 0.01, v) - DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b + 0.01, v)) / (2)
    ElseIf OutPutFlag = "b" Then 'Carry
        EDoubleBarrierBinaryAsymmetric = (DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b + 0.01, v) - DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b - 0.01, v)) / (2)
    ElseIf OutPutFlag = "s" Then 'Speed
        EDoubleBarrierBinaryAsymmetric = 1 / dS ^ 3 * (DoubleBarrierBinaryAsymmetric(TypeFlag, S + 2 * dS, L, U, k, T, r, b, v) - 3 * DoubleBarrierBinaryAsymmetric(TypeFlag, S + dS, L, U, k, T, r, b, v) _
                                + 3 * DoubleBarrierBinaryAsymmetric(TypeFlag, S, L, U, k, T, r, b, v) - DoubleBarrierBinaryAsymmetric(TypeFlag, S - dS, L, U, k, T, r, b, v))
    End If
    
End Function
