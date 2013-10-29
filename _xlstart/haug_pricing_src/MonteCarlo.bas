Attribute VB_Name = "MonteCarlo"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.


' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug


Public Function Max(X, y)
            Max = Application.Max(X, y)
End Function
'// Monte Carlo window barrier option using Brownian bridge probability
Public Function MonteCarloStandardBarrier(CallPutFlag As String, S As Double, X As Double, H As Double, T As Double, _
                r As Double, b As Double, v As Double, nSimulations As Long) As Double
            
    Dim ST As Double, sum As Double
    Dim Drift As Double, vSqrdt   As Double, BarrierHitProb As Double
    Dim i As Long, z As Integer
   
    Drift = (b - v ^ 2 / 2) * T
    vSqrdt = v * Sqr(T)
    
    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    For i = 0 To nSimulations
            ST = S * Exp(Drift + vSqrdt * Application.NormInv(Rnd(), 0, 1))
            
            If S > H Then
          '//Probability of hitting  barrier below
                If ST <= H Then
                      BarrierHitProb = 1
                Else
                        BarrierHitProb = Exp(-2 / (v ^ 2 * T) * Abs(Log(H / S) * Log(H / ST)))
                     
                End If
            ElseIf S < H Then
            '// Probability of hitting the barrier above
                If ST >= H Then
                    BarrierHitProb = 1
                Else
                    BarrierHitProb = Exp(-2 / (v ^ 2 * T) * Abs(Log(S / H) * Log(ST / H)))
                   
                End If
            End If
            sum = sum + (1 - BarrierHitProb) * Max(z * (ST - X), 0)
    Next

    MonteCarloStandardBarrier = Exp(-r * T) * (sum / nSimulations)

End Function


'// Monte Carlo window barrier option using Brownian bridge probability
Public Function MonteCarloWindowBarrier(CallPutFlag As String, S As Double, X As Double, H As Double, t1 As Double, t2 As Double, T As Double, _
                r As Double, b As Double, v As Double, nSimulations As Long) As Double
            
    Dim St1 As Double, St2 As Double, ST As Double
    Dim sum As Double, BarrierHitProb As Double
    Dim Drift1 As Double, Drift2 As Double, Drift3 As Double
    Dim vSqrdt1 As Double, vSqrdt2 As Double, vSqrdt3 As Double
    Dim i As Long, z As Integer
   
    Drift1 = (b - v ^ 2 / 2) * t1
    Drift2 = (b - v ^ 2 / 2) * (t2 - t1)
    Drift3 = (b - v ^ 2 / 2) * (T - t2)
    vSqrdt1 = v * Sqr(t1)
    vSqrdt2 = v * Sqr(t2 - t1)
    vSqrdt3 = v * Sqr(T - t2)
        
    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    For i = 0 To nSimulations
            St1 = S * Exp(Drift1 + vSqrdt1 * Application.NormInv(Rnd(), 0, 1))
            St2 = St1 * Exp(Drift2 + vSqrdt2 * Application.NormInv(Rnd(), 0, 1))
            ST = St2 * Exp(Drift3 + vSqrdt3 * Application.NormInv(Rnd(), 0, 1))
            
            If S > H Then
          '//Probability of hitting  barrier below
                If St2 <= H Or St1 <= H Then
                      BarrierHitProb = 1
                Else
                        BarrierHitProb = Exp(-2 / (v ^ 2 * T) * Abs(Log(H / St1) * Log(H / St2)))
                     
                End If
            ElseIf S < H Then
            '// Probability of hitting the barrier above
                If St2 >= H Or St1 >= H Then
                    BarrierHitProb = 1
                Else
                    BarrierHitProb = Exp(-2 / (v ^ 2 * T) * Abs(Log(St1 / H) * Log(St2 / H)))
                   
                End If
            End If
            
            sum = sum + (1 - BarrierHitProb) * Max(z * (ST - X), 0)
    Next

    MonteCarloWindowBarrier = Exp(-r * T) * (sum / nSimulations)

End Function

Public Function MonteCarloTripleAsset(CallPutFlag As String, S1 As Double, S2 As Double, _
S3 As Double, X As Double, T As Double, r As Double, b1 As Double, b2 As Double, _
b3 As Double, v1 As Double, v2 As Double, v3 As Double, rho12 As Double, _
rho13 As Double, rho23 As Double, nSimulations As Long) As Double

    Dim dt As Double, St1 As Double, St2 As Double, St3 As Double
    Dim i As Long, z As Integer
    Dim sum As Double, g As Double
    Dim Drift1 As Double, Drift2 As Double, Drift3 As Double
    Dim v1Sqrdt As Double, v2Sqrdt As Double, v3Sqrdt As Double
    Dim Epsilon1 As Double, Epsilon2 As Double, Epsilon3 As Double
    Dim alpha2 As Double, alpha3 As Double

 
    z = 1
    If CallPutFlag = "p" Then
        z = -1
    End If
    
    Drift1 = (b1 - v1 * v1 / 2) * T
    Drift2 = (b2 - v2 * v2 / 2) * T
    Drift3 = (b2 - v3 * v3 / 2) * T
    v1Sqrdt = v1 * Sqr(T)
    v2Sqrdt = v2 * Sqr(T)
    v3Sqrdt = v3 * Sqr(T)
      g = Sqr((1 - rho13 ^ 2) / (1 - rho12 ^ 2 - rho23 ^ 2 _
        - rho13 ^ 2 + 2 * rho12 * rho13 * rho23))
    
    sum = 0

    For i = 1 To nSimulations
       
        St1 = S1
        St2 = S2
        St3 = S3
       
        Epsilon1 = Application.NormInv(Rnd(), 0, 1)
        Epsilon2 = Application.NormInv(Rnd(), 0, 1)
        Epsilon3 = Application.NormInv(Rnd(), 0, 1)
        alpha2 = rho12 * Epsilon1 + Epsilon2 * Sqr(1 - rho12 ^ 2)
        alpha3 = Epsilon3 / g + (rho23 - rho13 * rho12) * Epsilon2 _
        + rho13 * Epsilon1 * Sqr(1 / (1 - rho12 ^ 2))
           
        St1 = St1 * Math.Exp(Drift1 + v1Sqrdt * Epsilon1)
        St2 = St2 * Math.Exp(Drift2 + v2Sqrdt * alpha2)
        St3 = St3 * Math.Exp(Drift3 + v3Sqrdt * alpha3)
            
        sum = sum + Application.Max(z * (St1 - St2 - X), z * (St3 - St2 - X), 0)
    Next

    MonteCarloTripleAsset = Exp(-r * T) * sum / nSimulations
    
End Function


Public Function QuasiMonteCarloTripleAsset(CallPutFlag As String, S1 As Double, S2 As Double, S3 As Double, _
                 X As Double, T As Double, r As Double, b1 As Double, b2 As Double, _
                 b3 As Double, v1 As Double, v2 As Double, v3 As Double, _
                 rho12 As Double, rho13 As Double, rho23 As Double, _
                      nSimulations As Long) As Double

    Dim dt As Double, St1 As Double, St2 As Double, St3 As Double
    Dim i As Long, z As Integer
    Dim sum As Double, g As Double
    Dim Drift1 As Double, Drift2 As Double, Drift3 As Double
    Dim v1Sqrdt As Double, v2Sqrdt As Double, v3Sqrdt As Double
    Dim Epsilon1 As Double, Epsilon2 As Double
    Dim alpha2 As Double, alpha3 As Double

 
    z = 1
    If CallPutFlag = "p" Then
        z = -1
    End If
    
    Drift1 = (b1 - v1 * v1 / 2) * T
    Drift2 = (b2 - v2 * v2 / 2) * T
    Drift3 = (b2 - v3 * v3 / 2) * T
    v1Sqrdt = v1 * Sqr(T)
    v2Sqrdt = v2 * Sqr(T)
    v3Sqrdt = v3 * Sqr(T)
    g = Sqr((1 - rho13 ^ 2) / (1 - rho12 ^ 2 - rho23 ^ 2 - rho13 ^ 2 + 2 * rho12 * rho13 * rho23))
        
    
    sum = 0

    For i = 1 To nSimulations

        St1 = S1
        St2 = S2
        St3 = S3
       
        Epsilon1 = BoxMuller(Halton(i, 3), Halton(i, 5))
        Epsilon2 = BoxMuller(Halton(i, 7), Halton(i, 11))
        alpha2 = rho12 * Epsilon1 + Epsilon2 * Sqr(1 - rho12 ^ 2)
        alpha3 = BoxMuller(Halton(i, 13), Halton(i, 15)) / g + (rho23 - rho13 * rho12) * Epsilon2 + rho13 * Epsilon1 * Sqr(1 / (1 - rho12 ^ 2))
           
        St1 = St1 * Math.Exp(Drift1 + v1Sqrdt * Epsilon1)
        St2 = St2 * Math.Exp(Drift2 + v2Sqrdt * alpha2)
        St3 = St3 * Math.Exp(Drift3 + v3Sqrdt * alpha3)
            
         sum = sum + Application.Max(z * (St1 - St2 - X), z * (St3 - St2 - X), 0)
    Next

    QuasiMonteCarloTripleAsset = Exp(-r * T) * sum / nSimulations
    
End Function




'// Company can force exersise if price MovingDaysN days in a row is above Barrier
Public Function CallableWarrantNDays(CallPutFlag As String, S As Double, X As Double, _
H As Double, T As Double, r As Double, b As Double, _
v As Double, DaysPerYear As Integer, _
nSimulations As Long, MovingDaysN As Integer) As Double

    Dim i As Long, j As Long
    Dim n As Long, Counter As Long
    Dim z As Integer
    
     Dim dt As Double, ST As Double, sum As Double, Drift As Double, vSqrt As Double
     Dim BarrierHitProb As Double
     
    z = 1
     If CallPutFlag = "p" Then
         z = -1
     End If
        
     n = DaysPerYear * T
     dt = T / n
     Drift = (b - v * v * 0.5) * dt
     vSqrt = v * Sqr(dt)
     sum = 0
    
    For j = 1 To nSimulations
        BarrierHitProb = 0
        ST = S
        Counter = 0
        For i = 2 To n '//starts at second fixing
        
            ST = ST * Exp(Drift + vSqrt * Application.NormInv(Rnd(), 0, 1))
       
            If z = 1 Then '// call
                If ST > H Then
                    Counter = Counter + 1
                Else
                    Counter = 0
                End If
            ElseIf z = -1 Then '//put
                If ST < H Then
                    Counter = Counter + 1
                Else
                    Counter = 0
                End If
            End If
          
            If Counter = MovingDaysN Then
              sum = sum + Exp(-r * (i * dt)) * Max(z * (ST - X), 0)
              BarrierHitProb = 1
              Exit For
           End If
       Next
       
        sum = sum + Exp(-r * T) * (1 - BarrierHitProb) * Max(z * (ST - X), 0)
   Next
   
   CallableWarrantNDays = sum / nSimulations
End Function
                        
       '// Monte Carlo plain vanilla European option using Antithetic variance reduction
Public Function MonteCarloStandardOptionAntithetic(CallPutFlag As String, S As Double, _
X As Double, T As Double, r As Double, b As Double, _
v As Double, nSimulations As Long) As Double
            
    Dim St1 As Double, St2 As Double, Epsilon As Double
    Dim sum As Double, Drift As Double, vSqrdt As Double
    Dim i As Long, j As Long, z As Integer

    
    Drift = (b - v ^ 2 / 2) * T
    vSqrdt = v * Sqr(T)

    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    For i = 1 To nSimulations
            Epsilon = Application.NormInv(Rnd(), 0, 1)
            St1 = S * Exp(Drift + vSqrdt * Epsilon)
            St2 = S * Exp(Drift + vSqrdt * (-Epsilon))
           
        sum = sum + (Max(z * (St1 - X), 0) + Max(z * (St2 - X), 0)) / 2
    Next

    MonteCarloStandardOptionAntithetic = Exp(-r * T) * sum / nSimulations

End Function

Public Function BoxMuller(X As Double, y As Double) As Double
    BoxMuller = Sqr(-2 * Log(X)) * Cos(2 * Application.Pi() * y)
End Function

Public Function BoxMuller2(x1 As Double, x2 As Double) As Variant

    Dim T As Double, L As Double
    Dim ReturnVec(1 To 2)

    If x1 = 0 Then
        BoxMuller2(x1, x2) = BoxMuller(x2, x1)
    Else
     '/ Using tan(Pi*x2) instead of cos and sin increases the speed by 30%
        T = Tan(Pi * x2)
        L = Sqr(-2 * Log(x1))
      
        ReturnVec(1) = L * (1 - T * T) / (1 + T * T)
        ReturnVec(2) = L * 2 * T / (1 + T * T)

        BoxMuller2 = ReturnVec()
    End If
End Function

' Halton Quasi Random Number Generator
Function Halton(n, b)
    Dim n0, n1, r As Integer
    Dim H As Double
    Dim f As Double
    n0 = n
    H = 0
    f = 1 / b
    While (n0 > 0)
        n1 = Int(n0 / b)
        r = n0 - n1 * b
        H = H + f * r
        f = f / b
        n0 = n1
    Wend
    Halton = H
End Function
Public Function StandardMCUsingBoxMuller(CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, nSimulations As Long) As Double
            
    Dim ST As Double
    Dim sum As Double, Drift As Double, vSqrdt As Double
    Dim i As Long, z As Integer
   
    
    Drift = (b - v ^ 2 / 2) * T
    vSqrdt = v * Sqr(T)

    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    For i = 1 To nSimulations
           ST = S * Exp(Drift + vSqrdt * BoxMuller(Halton(i, 3), Halton(i, 5)))
           
        sum = sum + Max(z * (ST - X), 0)
    Next

    StandardMCUsingBoxMuller = Exp(-r * T) * sum / nSimulations

End Function
Public Function HaltonMonteCarloStandardOption(CallPutFlag As String, S As Double, _
    X As Double, T As Double, r As Double, b As Double, _
    v As Double, nSimulations As Long) As Double
            
    Dim ST As Double
    Dim sum As Double, Drift As Double, vSqrdt As Double
    Dim i As Long, z As Integer
   
    
    Drift = (b - v ^ 2 / 2) * T
    vSqrdt = v * Sqr(T)

    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    For i = 1 To nSimulations
           ST = S * Exp(Drift + vSqrdt * BoxMuller(Halton(i, 3), Halton(i, 5)))
           
        sum = sum + Max(z * (ST - X), 0)
    Next

    HaltonMonteCarloStandardOption = Exp(-r * T) * sum / nSimulations

End Function

Public Function StandardMCWithGreeks(OutputFlag As String, CallPutFlag As String, S As Double, _
X As Double, T As Double, r As Double, b As Double, _
v As Double, nSimulations As Long) As Variant
            
    Dim ST As Double, Output() As Double
    Dim sum As Double, Drift As Double, vSqrdt As Double, DeltaSum As Double, GammaSum As Double
    Dim i As Long, z As Integer
   
   ReDim Output(0 To 4) As Double
    
    Drift = (b - v ^ 2 / 2) * T
    vSqrdt = v * Sqr(T)

    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    For i = 1 To nSimulations
        ST = S * Exp(Drift + vSqrdt * Application.NormInv(Rnd(), 0, 1))
        sum = sum + Max(z * (ST - X), 0)
        If z = 1 And ST > X Then DeltaSum = DeltaSum + ST
        If z = -1 And ST < X Then DeltaSum = DeltaSum + ST
        If Abs(ST - X) < 2 Then GammaSum = GammaSum + 1
    Next
    
    '// Option value:
    Output(0) = Exp(-r * T) * sum / nSimulations
     '// Delta:
    Output(1) = z * Exp(-r * T) * DeltaSum / (nSimulations * S)
    '// Gamma:
    Output(2) = Exp(-r * T) * (X / S) ^ 2 * GammaSum / (4 * nSimulations)
    '// Theta:
    Output(3) = (r * Output(0) - b * S * Output(1) - 0.5 * v ^ 2 * S ^ 2 * Output(2)) / 365
      '// Vega:
    Output(4) = Output(2) * v * S ^ 2 * T / 100
    
    StandardMCWithGreeks = Application.Transpose(Output())

End Function

'// Monte Carlo plain vanilla European option
Public Function SuperIQMC(CallPutFlag As String, S As Double, _
    X As Double, T As Double, r As Double, b As Double, _
    v As Double, nSimulations As Long) As Double
            
    Dim ST As Double
    Dim sum As Double, Drift As Double, vSqrdt As Double
    Dim i As Long, z As Integer
    Dim d As Double, d2 As Double, Epsilon As Double

    Drift = (b - v ^ 2 / 2) * T
    vSqrdt = v * Sqr(T)
    d = (Log(X / S) - (b - v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = (Log(S / X) + (b - v ^ 2 / 2) * T) / (v * Sqr(T))
    
    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    For i = 1 To nSimulations
        If z = -1 Then
            Epsilon = CND(d) * Halton(i, 5)
        Else
            Epsilon = (1 - CND(d)) * Halton(i, 5) + CND(d)
        End If
        ST = S * Exp(Drift + vSqrdt * Application.NormInv(Epsilon, 0, 1))
        sum = sum + Max(z * (ST - X), 0)
    Next
    SuperIQMC = Exp(-r * T) * sum / nSimulations * CND(z * d2)
    
End Function



'// Monte Carlo plain vanilla European option
Public Function IQMC(CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, nSimulations As Long) As Double
            
    Dim ST As Double
    Dim sum As Double, Drift As Double, vSqrdt As Double
    Dim i As Long, z As Integer
    Dim d As Double, d2 As Double, Epsilon As Double
   
    Drift = (b - v ^ 2 / 2) * T
    vSqrdt = v * Sqr(T)
    d = (Log(X / S) - (b - v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = (Log(S / X) + (b - v ^ 2 / 2) * T) / (v * Sqr(T))
    
    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    For i = 1 To nSimulations
        If z = -1 Then
            Epsilon = CND(d) * Rnd()
        Else
            Epsilon = (1 - CND(d)) * Rnd() + CND(d)
        End If
        ST = S * Exp(Drift + vSqrdt * Application.NormInv(Epsilon, 0, 1))
        sum = sum + Max(z * (ST - X), 0)
    Next
    IQMC = Exp(-r * T) * sum / nSimulations * CND(z * d2)
    
End Function















'// Monte Carlo plain vanilla European option
Public Function MonteCarloStandardOption(CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, nSimulations As Long) As Double
            
    Dim ST As Double
    Dim sum As Double, Drift As Double, vSqrdt As Double
    Dim i As Long, z As Integer
   
    Drift = (b - v ^ 2 / 2) * T
    vSqrdt = v * Sqr(T)

    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    For i = 1 To nSimulations
            ST = S * Exp(Drift + vSqrdt * Application.NormInv(Rnd(), 0, 1))
        sum = sum + Max(z * (ST - X), 0)
    Next

    MonteCarloStandardOption = Exp(-r * T) * sum / nSimulations

End Function



'// Monte Carlo plain vanilla European option
Public Function MonteCarloStandardOptionBoxMuller(CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, nSimulations As Long) As Double
            
    Dim St1 As Double, St2 As Double
    Dim sum As Double, Drift As Double, vSqrdt As Double
    Dim i As Long, z As Integer
    Dim StdRand As Variant
   

    Drift = (b - v ^ 2 / 2) * T
    vSqrdt = v * Sqr(T)

    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    For i = 1 To nSimulations
            StdRand = BoxMuller2(Rnd(), Rnd())
            St1 = S * Exp(Drift + vSqrdt * StdRand(1))
            St2 = S * Exp(Drift + vSqrdt * StdRand(2))
        sum = sum + Max(z * (St1 - X), 0) + Max(z * (St2 - X), 0)
    Next

    MonteCarloStandardOptionBoxMuller = Exp(-r * T) * (sum / (2 * nSimulations))

End Function


'// Monte Carlo plain vanilla European option
Public Function MonteCarloMeanReverting(CallPutFlag As String, S As Double, _
X As Double, T As Double, r As Double, b As Double, v As Double, _
kappa As Double, theta As Double, beta As Double, _
nSteps As Long, nSimulations As Long) As Double
            
    Dim dt As Double, ST As Double
    Dim sum As Double
    Dim i As Long, j As Long, z As Integer
   
    dt = T / nSteps

    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If
   
    For i = 1 To nSimulations
     ST = S
        For j = 0 To nSteps
            ST = ST + kappa * (theta - ST) * dt _
            + v * ST ^ beta * v * Sqr(dt) * Application.NormInv(Rnd(), 0, 1)
        Next
        sum = sum + Max(z * (ST - X), 0)
        
    Next

    MonteCarloMeanReverting = Exp(-r * T) * sum / nSimulations

End Function






                        
'// Monte Carlo two asset Asian spread option
Public Function MonteCarloAsianSpreadOption(CallPutFlag As String, S1 As Double, S2 As Double, _
                X As Double, T As Double, r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, _
                nSteps As Long, nSimulations As Long) As Double
            
    Dim dt As Double, St1 As Double, St2 As Double
    Dim i As Long, j As Long, z As Integer
    Dim sum As Double, Drift1 As Double, Drift2 As Double
    Dim v1Sqrdt As Double, v2Sqrdt As Double
    Dim Epsilon1 As Double, Epsilon2 As Double, Average1 As Double, Average2 As Double

    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    dt = T / nSteps
    Drift1 = (b1 - v1 ^ 2 / 2) * dt
    Drift2 = (b2 - v2 ^ 2 / 2) * dt
    v1Sqrdt = v1 * Sqr(dt)
    v2Sqrdt = v2 * Sqr(dt)

    For i = 1 To nSimulations
        Average1 = 0
        Average2 = 0
        St1 = S1
        St2 = S2
        For j = 1 To nSteps
            Epsilon1 = Application.NormInv(Rnd(), 0, 1)
            Epsilon2 = rho * Epsilon1 + Application.NormInv(Rnd(), 0, 1) * Sqr(1 - rho ^ 2)
            St1 = St1 * Exp(Drift1 + v1Sqrdt * Epsilon1)
            St2 = St2 * Exp(Drift2 + v2Sqrdt * Epsilon2)
            Average1 = Average1 + St1
            Average2 = Average2 + St2
        Next
        Average1 = Average1 / nSteps
        Average2 = Average2 / nSteps
        sum = sum + Max(z * (Average1 - Average2 - X), 0)
    Next

    MonteCarloAsianSpreadOption = Exp(-r * T) * sum / nSimulations
    
End Function



'// Monte Carlo two asset Asian spread option
Public Function MC2Asset(CallPutFlag As String, S1 As Double, S2 As Double, _
                x1 As Double, x2 As Double, T As Double, r As Double, b1 As Double, b2 As Double, v1 As Double, v2 As Double, rho As Double, _
                nSimulations As Long) As Double
            
    Dim dt As Double, St1 As Double, St2 As Double
    Dim i As Long, j As Long, z As Integer
    Dim sum As Double, Drift1 As Double, Drift2 As Double
    Dim v1Sqrdt As Double, v2Sqrdt As Double
    Dim Epsilon1 As Double, Epsilon2 As Double, Average1 As Double, Average2 As Double

    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    Drift1 = (b1 - v1 ^ 2 / 2) * T
    Drift2 = (b2 - v2 ^ 2 / 2) * T
    v1Sqrdt = v1 * Sqr(T)
    v2Sqrdt = v2 * Sqr(T)

    For i = 1 To nSimulations
       
            Epsilon1 = Application.NormInv(Rnd(), 0, 1)
            Epsilon2 = rho * Epsilon1 + Application.NormInv(Rnd(), 0, 1) * Sqr(1 - rho ^ 2)
            St1 = S1 * Exp(Drift1 + v1Sqrdt * Epsilon1)
            St2 = S2 * Exp(Drift2 + v2Sqrdt * Epsilon2)

            If St1 > x1 Then
                sum = sum + Max(z * (St2 - x2), 0)
            End If
    Next

    MC2Asset = Exp(-r * T) * sum / nSimulations
    
End Function


'// IQ-MC two asset correlation option
Public Function IQMC2Asset(CallPutFlag As String, S1 As Double, S2 As Double, _
        x1 As Double, x2 As Double, T As Double, r As Double, b1 As Double, _
        b2 As Double, v1 As Double, v2 As Double, rho As Double, _
        nSimulations As Long) As Double
            
    Dim dt As Double, St1 As Double, St2 As Double
    Dim i As Long, j As Long, z As Integer
    Dim sum As Double, Drift1 As Double, Drift2 As Double
    Dim v1Sqrdt As Double, v2Sqrdt As Double
    Dim y1 As Double, y2 As Double, dd As Double
    Dim Epsilon1 As Double, Epsilon2 As Double
    
    Dim d As Double, d2 As Double

    If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If

    Drift1 = (b1 - v1 ^ 2 / 2) * T
    Drift2 = (b2 - v2 ^ 2 / 2) * T
    v1Sqrdt = v1 * Sqr(T)
    v2Sqrdt = v2 * Sqr(T)
    
    y1 = (Log(S1 / x1) + (b1 - v1 ^ 2 / 2) * T) / (v1 * Sqr(T))
    y2 = (Log(S2 / x2) + (b2 - v2 ^ 2 / 2) * T) / (v2 * Sqr(T))
    d = (Log(x1 / S1) - (b1 - v1 ^ 2 / 2) * T) / (v1 * Sqr(T))
    dd = (Log(x2 / S2) - (b2 - v2 ^ 2 / 2) * T) / (v2 * Sqr(T))
   
    For i = 1 To nSimulations
       
        If z = -1 Then
            Epsilon1 = CND(d) * Rnd()
        Else
            Epsilon1 = (1 - CND(d)) * Rnd() + CND(d)
        End If
        Epsilon1 = Application.NormInv(Epsilon1, 0, 1)
        Epsilon2 = CND((dd - rho * Epsilon1) / Sqr(1 - rho * rho))
        If z = 1 Then
            Epsilon2 = (1 - Epsilon2) * Rnd() + Epsilon2
        Else
            Epsilon2 = Epsilon2 * Rnd()
        End If
        Epsilon2 = rho * Epsilon1 + Application.NormInv(Epsilon2, 0, 1) * Sqr(1 - rho ^ 2)
        St1 = S1 * Exp(Drift1 + v1Sqrdt * Epsilon1)
        St2 = S2 * Exp(Drift2 + v2Sqrdt * Epsilon2)
        sum = sum + z * (St2 - x2)
    Next

    IQMC2Asset = Exp(-r * T) * sum / nSimulations * CBND(z * y1, z * y2, rho)
        
End Function
