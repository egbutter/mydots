Attribute VB_Name = "StandardTrees"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.

Option Base 0       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug


Public Function Max(X, y)
            Max = Application.Max(X, y)
End Function

Public Function TrinomialTree(OutputFLag As String, AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, _
            T As Double, r As Double, b As Double, v As Double, n As Integer) As Variant


    Dim OptionValue() As Double
    ReDim OptionValue(0 To n * 2 + 1)

    Dim ReturnValue(3) As Double
    Dim dt As Double, u As Double, d As Double
    Dim pu As Double, pd As Double, pm As Double, Df As Double
    Dim i As Long, j As Long, z As Integer

     If CallPutFlag = "c" Then
        z = 1
    ElseIf CallPutFlag = "p" Then
        z = -1
    End If
    
    dt = T / n
    u = Exp(v * Sqr(2 * dt))
    d = Exp(-v * Sqr(2 * dt))
    pu = ((Exp(b * dt / 2) - Exp(-v * Sqr(dt / 2))) / (Exp(v * Sqr(dt / 2)) - Exp(-v * Sqr(dt / 2)))) ^ 2
   
    pd = ((Exp(v * Sqr(dt / 2)) - Exp(b * dt / 2)) / (Exp(v * Sqr(dt / 2)) - Exp(-v * Sqr(dt / 2)))) ^ 2
    pm = 1 - pu - pd
    Df = Exp(-r * dt)
    
    For i = 0 To (2 * n)
         OptionValue(i) = Max(0, z * (S * u ^ Max(i - n, 0) * d ^ Max(n - i, 0) - X))
    Next
    
    For j = n - 1 To 0 Step -1
    
        For i = 0 To (j * 2)
        
            OptionValue(i) = (pu * OptionValue(i + 2) _
                + pm * OptionValue(i + 1) + pd * OptionValue(i)) * Df
            
            If AmeEurFlag = "a" Then
                OptionValue(i) = Max(z * (S * u ^ Max(i - j, 0) _
                    * d ^ Max(j - i, 0) - X), OptionValue(i))
            End If
        Next
     If j = 1 Then
           ReturnValue(1) = (OptionValue(2) - OptionValue(0)) / (S * u - S * d)
           ReturnValue(2) = ((OptionValue(2) - OptionValue(1)) / (S * u - S) _
                - (OptionValue(1) - OptionValue(0)) / (S - S * d)) / (0.5 * (S * u - S * d))
           ReturnValue(3) = OptionValue(1)
      End If
    Next
    
    ReturnValue(3) = (ReturnValue(3) - OptionValue(0)) / dt / 365
    
    ReturnValue(0) = OptionValue(0)
    If OutputFLag = "p" Then 'Option value
        TrinomialTree = ReturnValue(0)
    ElseIf OutputFLag = "d" Then 'Delta
        TrinomialTree = ReturnValue(1)
    ElseIf OutputFLag = "g" Then 'Gamma
        TrinomialTree = ReturnValue(2)
    ElseIf OutputFLag = "t" Then 'Theta
        TrinomialTree = ReturnValue(3)
   ElseIf OutputFLag = "a" Then ' All
        TrinomialTree = Application.Transpose(ReturnValue())
    End If
    
 End Function



Public Function EuropeanBinomialPlainVanilla(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, v As Double, n As Long) As Double
        
        Dim u As Double, d As Double, p As Double
        Dim Sum As Double, dt As Double, A As Double
        Dim j As Long
        
        dt = T / n
        u = Exp(v * Sqr(dt))
        d = 1 / u
        p = (Exp(b * dt) - d) / (u - d)
        A = Int(Log(X / (S * d ^ n)) / Log(u / d)) + 1

    Sum = 0
    If CallPutFlag = "c" Then
        For j = A To n
            Sum = Sum + Application.Combin(n, j) * p ^ j * (1 - p) ^ (n - j) * (S * u ^ j * d ^ (n - j) - X)
        Next
    ElseIf CallPutFlag = "p" Then
         For j = 0 To A - 1
            Sum = Sum + Application.Combin(n, j) * p ^ j * (1 - p) ^ (n - j) * (X - S * u ^ j * d ^ (n - j))
        Next
    End If
    EuropeanBinomialPlainVanilla = Exp(-r * T) * Sum
End Function



'// Cox-Ross-Rubinstein binomial tree
Public Function CRRBinomial(OutputFLag As String, AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, n As Integer) As Variant
                

    Dim OptionValue() As Double
    Dim u As Double, d As Double, p As Double
    Dim ReturnValue(4) As Double
    Dim dt As Double, Df As Double
    Dim i As Integer, j As Integer, z As Integer
    
    ReDim OptionValue(0 To n + 1)
    
    If CallPutFlag = "c" Then
        z = 1
        ElseIf CallPutFlag = "p" Then
        z = -1
    End If
    
    dt = T / n
    u = Exp(v * Sqr(dt))
    d = 1 / u
    p = (Exp(b * dt) - d) / (u - d)
    Df = Exp(-r * dt)
    
    For i = 0 To n
         OptionValue(i) = Max(0, z * (S * u ^ i * d ^ (n - i) - X))
    Next
    
    For j = n - 1 To 0 Step -1
        For i = 0 To j
            If AmeEurFlag = "e" Then
                OptionValue(i) = (p * OptionValue(i + 1) + (1 - p) * OptionValue(i)) * Df
            ElseIf AmeEurFlag = "a" Then
                OptionValue(i) = Max((z * (S * u ^ i * d ^ (j - i) - X)), _
                (p * OptionValue(i + 1) + (1 - p) * OptionValue(i)) * Df)
            End If
           
        Next
         If j = 2 Then
                ReturnValue(2) = ((OptionValue(2) - OptionValue(1)) / (S * u ^ 2 - S) _
                - (OptionValue(1) - OptionValue(0)) / (S - S * d ^ 2)) / (0.5 * (S * u ^ 2 - S * d ^ 2))
                ReturnValue(3) = OptionValue(1)
            End If
         If j = 1 Then
                ReturnValue(1) = (OptionValue(1) - OptionValue(0)) / (S * u - S * d)
            End If
    Next
    ReturnValue(3) = (ReturnValue(3) - OptionValue(0)) / (2 * dt) / 365
    ReturnValue(0) = OptionValue(0)
    If OutputFLag = "p" Then 'Option value
        CRRBinomial = ReturnValue(0)
    ElseIf OutputFLag = "d" Then 'Delta
        CRRBinomial = ReturnValue(1)
    ElseIf OutputFLag = "g" Then 'Gamma
        CRRBinomial = ReturnValue(2)
    ElseIf OutputFLag = "t" Then 'Theta
        CRRBinomial = ReturnValue(3)
   ElseIf OutputFLag = "a" Then
        CRRBinomial = Application.Transpose(ReturnValue())
    End If
    
End Function



'// Leisen-Reimer binomial tree
Public Function LeisenReimerBinomial(OutputFLag As String, AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, n As Integer) As Variant
                

    Dim OptionValue() As Double
    Dim ReturnValue(3) As Double
    Dim d1 As Double, d2 As Double
    Dim hd1 As Double, hd2 As Double
    Dim u As Double, d As Double, p As Double
    Dim dt As Double, Df As Double
    Dim i As Integer, j As Integer, z As Integer
    
    n = Application.Odd(n)
    
    ReDim OptionValue(0 To n)
    
    If CallPutFlag = "c" Then
        z = 1
        ElseIf CallPutFlag = "p" Then
        z = -1
    End If
    
    d1 = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T))
    d2 = d1 - v * Sqr(T)

    '// Using Preizer-Pratt inversion method 2
    hd1 = 0.5 + Sgn(d1) * (0.25 - 0.25 * Exp(-(d1 / (n + 1 / 3 + 0.1 / (n + 1))) ^ 2 * (n + 1 / 6))) ^ 0.5
    hd2 = 0.5 + Sgn(d2) * (0.25 - 0.25 * Exp(-(d2 / (n + 1 / 3 + 0.1 / (n + 1))) ^ 2 * (n + 1 / 6))) ^ 0.5
    
    dt = T / n
    p = hd2
    u = Exp(b * dt) * hd1 / hd2
    d = (Exp(b * dt) - p * u) / (1 - p)
    Df = Exp(-r * dt)
    For i = 0 To n
         OptionValue(i) = Max(0, z * (S * u ^ i * d ^ (n - i) - X))
    Next
    
    For j = n - 1 To 0 Step -1
        For i = 0 To j
            If AmeEurFlag = "e" Then
                OptionValue(i) = (p * OptionValue(i + 1) + (1 - p) * OptionValue(i)) * Df
            ElseIf AmeEurFlag = "a" Then
                OptionValue(i) = Max((z * (S * u ^ i * d ^ (j - i) - X)), _
                (p * OptionValue(i + 1) + (1 - p) * OptionValue(i)) * Df)
            End If
           
        Next
        If j = 2 Then
                ReturnValue(2) = ((OptionValue(2) - OptionValue(1)) / (S * u ^ 2 - S * u * d) _
                - (OptionValue(1) - OptionValue(0)) / (S * u * d - S * d ^ 2)) / (0.5 * (S * u ^ 2 - S * d ^ 2))
                ReturnValue(3) = OptionValue(1)
            End If
         If j = 1 Then
                ReturnValue(1) = (OptionValue(1) - OptionValue(0)) / (S * u - S * d)
            End If
    Next
    ReturnValue(0) = OptionValue(0)
    If OutputFLag = "p" Then 'Option value
        LeisenReimerBinomial = ReturnValue(0)
    ElseIf OutputFLag = "d" Then 'Delta
        LeisenReimerBinomial = ReturnValue(1)
    ElseIf OutputFLag = "g" Then 'Gamma
        LeisenReimerBinomial = ReturnValue(2)
   ElseIf OutputFLag = "a" Then
        LeisenReimerBinomial = Application.Transpose(ReturnValue())
    End If
End Function







'// Rendelman-Barter binomial tree
Public Function JarrowRuddBinomial(OutputFLag As String, AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, n As Integer) As Variant
                

    Dim OptionValue() As Double
    Dim u As Double, d As Double, p As Double
    Dim ReturnValue(4) As Double
    Dim dt As Double, Df As Double
    Dim i As Integer, j As Integer, z As Integer
    
    ReDim OptionValue(0 To n + 1)
    
    If CallPutFlag = "c" Then
        z = 1
        ElseIf CallPutFlag = "p" Then
        z = -1
    End If
    
    dt = T / n
    u = Exp((b - v ^ 2 / 2) * dt + v * Sqr(dt))
    d = Exp((b - v ^ 2 / 2) * dt - v * Sqr(dt))
    p = 0.5
    Df = Exp(-r * dt)
    
    For i = 0 To n
         OptionValue(i) = Max(0, z * (S * u ^ i * d ^ (n - i) - X))
    Next
    
    For j = n - 1 To 0 Step -1:
        For i = 0 To j
            If AmeEurFlag = "e" Then
                OptionValue(i) = (p * OptionValue(i + 1) + (1 - p) * OptionValue(i)) * Df
            ElseIf AmeEurFlag = "a" Then
                OptionValue(i) = Max((z * (S * u ^ i * d ^ (j - i) - X)), _
                (p * OptionValue(i + 1) + (1 - p) * OptionValue(i)) * Df)
            End If
           
        Next
         If j = 2 Then
                ReturnValue(2) = ((OptionValue(2) - OptionValue(1)) / (S * u ^ 2 - S * u * d) _
                - (OptionValue(1) - OptionValue(0)) / (S * u * d - S * d ^ 2)) / (0.5 * (S * u ^ 2 - S * d ^ 2))
            End If
         If j = 1 Then
                ReturnValue(1) = (OptionValue(1) - OptionValue(0)) / (S * u - S * d)
            End If
    Next
    ReturnValue(3) = (ReturnValue(3) - OptionValue(0)) / (2 * dt) / 365
    ReturnValue(0) = OptionValue(0)
    If OutputFLag = "p" Then 'Option value
        JarrowRuddBinomial = ReturnValue(0)
    ElseIf OutputFLag = "d" Then 'Delta
        JarrowRuddBinomial = ReturnValue(1)
    ElseIf OutputFLag = "g" Then 'Gamma
        JarrowRuddBinomial = ReturnValue(2)
     ElseIf OutputFLag = "a" Then
        JarrowRuddBinomial = Application.Transpose(ReturnValue())
    End If
End Function

