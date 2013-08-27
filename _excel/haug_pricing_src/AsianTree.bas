Attribute VB_Name = "AsianTree"
Global Const Pi = 3.14159265358979

Option Explicit     'Requirs that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.
Option Base 0

'Programming By Espen Gaarder Haug, Copyright 2006

Public Function Max(X, y)
            Max = Application.Max(X, y)
End Function

Public Function Min(X, y)
            Min = Application.Min(X, y)
End Function
                                                                                                         


'// Implemented to hold only when number of fixings = number of time steps
Public Function AsianTrinomialTree(AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, _
    v As Double, alpha As Integer, N As Integer) As Double

    ' Alpha integer from 1 and uppwards higher alpha gives higher accurancy

    Dim i As Integer, j As Integer, K As Integer
    Dim pu As Double, pm As Double, pd As Double
    Dim Fu As Double, Fm As Double, Fd As Double
    Dim Cu As Double, Cm As Double, Cd As Double
    Dim St() As Double, nn() As Double, F() As Double, c() As Double
    Dim h As Double, dt As Double, Df As Double
    Dim u As Double, d As Double
    Dim m As Integer, z As Integer
    Dim kiu As Integer, kim As Integer, kid As Integer
    
   ReDim St(0 To N, 0 To N * 2)
   ReDim nn(0 To N, 0 To N * 2)
   ReDim F(0 To N, 0 To N * 2, 0 To 1 + alpha * N)
   ReDim c(0 To N, 0 To N * 2, 0 To 1 + alpha * N)
   
   If CallPutFlag = "c" Then
        z = 1
    Else
        z = -1
    End If
    
    nn(0, 0) = 1
    F(0, 0, 1) = S
    
    dt = T / N
    u = Exp(v * Sqr(2 * dt))
    d = 1 / u
    pu = ((Exp(b * dt / 2) - Exp(-v * Sqr(dt / 2))) / (Exp(v * Sqr(dt / 2)) - Exp(-v * Sqr(dt / 2)))) ^ 2
    pd = ((Exp(v * Sqr(dt / 2)) - Exp(b * dt / 2)) / (Exp(v * Sqr(dt / 2)) - Exp(-v * Sqr(dt / 2)))) ^ 2
    pm = 1 - pu - pd
    
    Df = Exp(-r * dt)
    St(0, 0) = S
    
    For j = N To 0 Step -1  ' Build state space
        For i = 0 To j * 2
            St(j, i) = S * u ^ Max(i - j, 0) * d ^ Max(j * 2 - j - i, 0)
        Next
  Next
                
    For i = 1 To N
        For j = 0 To 2 * i
            ' Minimum at each node
            If j = 0 Then
                F(i, j, 1) = (F(i - 1, 0, 1) * i + St(i, 0)) / (i + 1)
            ElseIf j = 1 Then
                  F(i, j, 1) = (F(i - 1, 0, 1) * i + St(i, 1)) / (i + 1)
            Else
                  F(i, j, 1) = (F(i - 1, j - 2, 1) * i + St(i, j)) / (i + 1)
            End If
            
            m = -i + j
            nn(i, j) = 1 + alpha * (i - Abs(m))
            
            ' Maximum at each node
            If j > 2 * i - 1 Then
                F(i, j, nn(i, j)) = (F(i - 1, j - 2, nn(i - 1, j - 2)) * i + St(i, j)) / (i + 1)
            ElseIf j > 2 * i - 2 Then
                F(i, j, nn(i, j)) = (F(i - 1, j - 1, nn(i - 1, j - 1)) * i + St(i, j)) / (i + 1)
            Else
                F(i, j, nn(i, j)) = (F(i - 1, j, nn(i - 1, j)) * i + St(i, j)) / (i + 1)
            End If
             
             'Intermidiate average values
             If nn(i, j) > 2 Then
                  h = Log(F(i, j, nn(i, j)) / F(i, j, 1)) / (nn(i, j) - 1)
            End If
            For K = 2 To nn(i, j) - 1
                F(i, j, K) = F(i, j, 1) * Exp((K - 1) * h)
            Next
        Next
    Next

    For j = 0 To 2 * N ' Option values at maturity
        For K = 1 To nn(N, j)
            c(N, j, K) = Application.Max(0, z * (F(N, j, K) - X))
        Next
    Next
            
    For i = N - 1 To 0 Step -1
        For j = 0 To 2 * i
            For K = 1 To nn(i, j)
                Fu = (F(i, j, K) * (i + 1) + St(i + 1, j + 2)) / (i + 2)
                Fm = (F(i, j, K) * (i + 1) + St(i + 1, j + 1)) / (i + 2)
                Fd = (F(i, j, K) * (i + 1) + St(i + 1, j)) / (i + 2)
            If j = 0 Then
                Cu = c(i + 1, j + 2, 1)
                Cm = c(i + 1, j + 1, 1)
                Cd = c(i + 1, j, 1)
            ElseIf j = 2 * i Then
                Cu = c(i + 1, j + 2, 1)
                Cm = c(i + 1, j + 1, nn(i + 1, j + 1))
                Cd = c(i + 1, j, nn(i + 1, j))
           Else
                    kiu = 1
                    While Fu < F(i + 1, j + 2, kiu) Or Fu > F(i + 1, j + 2, kiu + 1)
                        kiu = kiu + 1
                    Wend
                    kim = 1
                    While Fm < F(i + 1, j + 1, kim) Or Fm > F(i + 1, j + 1, kim + 1)
                        kim = kim + 1
                    Wend
                    kid = 1
                    While Fd < F(i + 1, j, kid) Or Fd > F(i + 1, j, kid + 1)
                        kid = kid + 1
                    Wend
                    Cu = c(i + 1, j + 2, kiu) + (c(i + 1, j + 2, kiu + 1) - c(i + 1, j + 2, kiu)) / (F(i + 1, j + 2, kiu + 1) - F(i + 1, j + 2, kiu)) * (Fu - F(i + 1, j + 2, kiu))
                    Cm = c(i + 1, j + 1, kim) + (c(i + 1, j + 1, kim + 1) - c(i + 1, j + 1, kim)) / (F(i + 1, j + 1, kim + 1) - F(i + 1, j + 1, kim)) * (Fm - F(i + 1, j + 1, kim))
                    Cd = c(i + 1, j, kid) + (c(i + 1, j, kid + 1) - c(i + 1, j, kid)) / (F(i + 1, j, kid + 1) - F(i + 1, j, kid)) * (Fd - F(i + 1, j, kid))
            End If
            
                c(i, j, K) = Df * (pu * Cu + pm * Cm + pd * Cd)
                If AmeEurFlag = "a" Then
                    c(i, j, K) = Application.Max(c(i, j, K), z * (F(i, j, K) - X))
                End If
                
            Next
        Next
    Next
    
    AsianTrinomialTree = c(0, 0, 1)
    
End Function





