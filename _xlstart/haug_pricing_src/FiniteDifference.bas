Attribute VB_Name = "FiniteDifference"
Option Base 0
Option Explicit


' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug


Private Function Max(X As Double, y As Double) As Double
    
    Max = Application.Max(X, y)

End Function


Public Function ImplicitFiniteDifference(AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, _
                                                                v As Double, N As Integer, M As Integer) As Double

    Dim p() As Variant, CT() As Variant, C As Variant
    Dim dS As Double, dt As Double
    Dim i As Integer, z As Integer, j As Integer
    Dim SGridPt As Integer
    
    z = 1
    If CallPutFlag = "p" Then z = -1
   
   '// Makes sure current asset price falls at grid point
    dS = 2 * S / M
    SGridPt = S / dS
    M = Int(X / dS) * 2
    dt = T / N
    
    ReDim CT(0 To M)
    ReDim p(0 To M, 0 To M)
    
    
    For j = 0 To M
        CT(j) = Max(0, z * (j * dS - X)) '//Option value at maturity
        For i = 0 To M
            p(j, i) = 0
        Next
    Next
            
    p(0, 0) = 1
    For i = 1 To M - 1 Step 1
        p(i, i - 1) = 0.5 * i * (b - v ^ 2 * i) * dt
        p(i, i) = 1 + (r + v ^ 2 * i ^ 2) * dt
        p(i, i + 1) = 0.5 * i * (-b - v ^ 2 * i) * dt
    Next
    
    p(M, M) = 1
   C = Application.MMult(Application.MInverse(p()), Application.Transpose(CT()))
   For j = N - 1 To 1 Step -1
        C = Application.MMult(Application.MInverse(p()), C)
        
        If AmeEurFlag = "a" Then '//American option
            For i = 1 To M
                C(i, 1) = Max(CDbl(C(i, 1)), z * ((i - 1) * dS - X))
             Next
        End If
    Next
    
    ImplicitFiniteDifference = C(SGridPt + 1, 1)
End Function
' Standard Explicite Difference for call option constant volatility
Public Function ExplicitFiniteDifferenceLnS(AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, v As Double, N As Integer, M As Integer) As Double

    Dim C() As Double, St() As Double
    Dim dt As Double, dx As Double
    Dim pu As Double, pm As Double, pd As Double
    Dim i As Integer, j As Integer, z As Integer

    ReDim C(0 To M / 2, 0 To M + 1)
    ReDim St(0 To M + 1)
    
    z = 1
    If CallPutFlag = "p" Then z = -1
    
    dt = T / N
    dx = v * Sqr(3 * dt)
    pu = 0.5 * dt * ((v / dx) ^ 2 + (b - v ^ 2 / 2) / dx)
    pm = 1 - dt * (v / dx) ^ 2 - r * dt
    pd = 0.5 * dt * ((v / dx) ^ 2 - (b - v ^ 2 / 2) / dx)
    
    St(0) = S * Exp(-M / 2 * dx)
    C(N, 0) = Max(0, z * (St(0) - X))
    
    For i = 1 To M
        St(i) = St(i - 1) * Exp(dx)   ' // asset price at maturity
        C(N, i) = Max(0, z * (St(i) - X)) '// Option value at maturity
    Next
    
    For j = N - 1 To 0 Step -1
        For i = 1 To M - 1
            C(j, i) = pu * C(j + 1, i + 1) + pm * C(j + 1, i) + pd * C(j + 1, i - 1)
            If AmeEurFlag = "a" Then '//American option
                C(j, i) = Max(C(j, i), z * (St(i) - X))
            End If
        Next
        C(j, M) = C(j, M - 1) + St(M) - St(M - 1)     '//Upper boundary
        C(j, 0) = C(j, 1) '// Lower boundary
    Next
    ExplicitFiniteDifferenceLnS = C(0, M / 2)
    
End Function

Public Function ExplicitFiniteDifference(AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                                                            r As Double, b As Double, v As Double, M As Integer) As Double

    Dim C() As Double, St() As Double
    Dim dt As Double, dS As Double
    Dim pu As Double, pm As Double, pd As Double, Df As Double
    Dim i As Integer, j As Integer, N As Integer, z As Integer
    Dim SGridPt As Integer
    
    z = 1
    If CallPutFlag = "p" Then z = -1
    
    dS = S / M
    M = Int(X / dS) * 2
    ReDim St(0 To M + 1)
    
    SGridPt = S / dS
    dt = dS ^ 2 / (v ^ 2 * 4 * X ^ 2)
    N = Int(T / dt) + 1
    
    ReDim C(0 To N, 0 To M + 1)
    dt = T / N
    Df = 1 / (1 + r * dt)
    
    For i = 0 To M
         St(i) = i * dS ' // Asset price at maturity
         C(N, i) = Max(0, z * (St(i) - X)) '// Option value at maturity
    Next
    For j = N - 1 To 0 Step -1
        For i = 1 To M - 1
        pu = 0.5 * (v ^ 2 * i ^ 2 + b * i) * dt
        pm = 1 - v ^ 2 * i ^ 2 * dt
        pd = 0.5 * (v ^ 2 * i ^ 2 - b * i) * dt
        
        C(j, i) = Df * (pu * C(j + 1, i + 1) + pm * C(j + 1, i) + pd * C(j + 1, i - 1))
        If AmeEurFlag = "a" Then
            C(j, i) = Max(z * (St(i) - X), C(j, i))
        End If
        Next
        If z = 1 Then '//Call option
            C(j, 0) = 0
            C(j, M) = (St(i) - X)
        Else
            C(j, 0) = X
            C(j, M) = 0
        End If
    Next
    ExplicitFiniteDifference = C(0, SGridPt)
    
End Function

Public Function CrankNickolson(AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, v As Double, _
            N As Integer, M As Integer) As Double


    Dim C() As Double, St() As Double, p() As Double, pmd() As Double
    Dim dt As Double, dx As Double
    Dim i As Integer, j As Integer, z As Integer
    Dim pu As Double, pm As Double, pd As Double
    
    ReDim pmd(0 To M)
    ReDim p(0 To M)
    ReDim C(0 To M / 2 + 1, 0 To M + 1)
    ReDim St(0 To M + 1)
    
    z = 1
    If CallPutFlag = "p" Then z = -1
    
    dt = T / N
    dx = v * Sqr(3 * dt)
    pu = -0.25 * dt * ((v / dx) ^ 2 + (b - 0.5 * v ^ 2) / dx)
    pm = 1 + 0.5 * dt * (v / dx) ^ 2 + 0.5 * r * dt
    pd = -0.25 * dt * ((v / dx) ^ 2 - (b - 0.5 * v ^ 2) / dx)
    
    St(0) = S * Exp(-M / 2 * dx)
    C(0, 0) = Max(0, z * (St(0) - X))
    For i = 1 To M '// Option value at maturity
        St(i) = St(i - 1) * Exp(dx)
        C(0, i) = Max(0, z * (St(i) - X))
    Next
    
    pmd(1) = pm + pd
    p(1) = -pu * C(0, 2) - (pm - 2) * C(0, 1) - pd * C(0, 0) - pd * (St(1) - St(0))
    For j = N - 1 To 0 Step -1

        For i = 2 To M - 1
            p(i) = -pu * C(0, i + 1) - (pm - 2) * C(0, i) - pd * C(0, i - 1) - p(i - 1) * pd / pmd(i - 1)
            pmd(i) = pm - pu * pd / pmd(i - 1)
        Next
        
        
        For i = M - 2 To 1 Step -1
            C(1, i) = (p(i) - pu * C(1, i + 1)) / pmd(i)
        Next
        
            For i = 0 To M
                C(0, i) = C(1, i)
                If AmeEurFlag = "a" Then
                    C(0, i) = Max(C(1, i), z * (St(i) - X))
                 End If
            Next
       
    Next
    
    CrankNickolson = C(0, M / 2)
    
End Function
