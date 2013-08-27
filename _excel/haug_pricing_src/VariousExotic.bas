Attribute VB_Name = "VariousExotic"
Option Explicit

' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug


'// Cox-Ross-Rubinstein binomial tree
Public Function BinomialCompoundOption(OutputFLag As String, CompoundEurAmeFlag As String, AmeEurFlag As String, CompoundTypeFlag As String, S As Double, X1 As Double, X2 As Double, t1 As Double, T2 As Double, _
                r As Double, b As Double, v As Double, n As Integer) As Variant
                

    Dim OptionValue() As Double, ReturnValue(3) As Double
    Dim u As Double, d As Double, p As Double
    Dim dt As Double, Df As Double
    Dim i As Integer, j As Integer, z As Integer, y As Integer, w As Integer
    
    ReDim OptionValue(0 To n + 1)
    
    If CompoundTypeFlag = "cc" Or CompoundTypeFlag = "pc" Then
        z = 1
    Else
        z = -1
    End If
    
    
    If CompoundTypeFlag = "cc" Or CompoundTypeFlag = "cp" Then
        y = 1
    Else
        y = -1
    End If
 
    w = 1
    dt = T2 / n
    u = Exp(v * Sqr(dt))
    d = 1 / u
    p = (Exp(b * dt) - d) / (u - d)
    Df = Exp(-r * dt)
    
    For i = 0 To n
         OptionValue(i) = Max(0, z * (S * u ^ i * d ^ (n - i) - X1))
    Next
    
    For j = n - 1 To 0 Step -1
        For i = 0 To j
          
            OptionValue(i) = (p * OptionValue(i + 1) + (1 - p) * OptionValue(i)) * Df
            
            If AmeEurFlag = "a" Then
                OptionValue(i) = Max((z * (S * u ^ i * d ^ (j - i) - X1)), OptionValue(i))
            End If
            
            If t1 >= dt * j And w = 1 Then
                    OptionValue(i) = Max(y * (OptionValue(i) - X2), 0)
                    If i = j Then
                        w = -1
                    End If
            End If
            If w = -1 And CompoundEurAmeFlag = "a" Then
                OptionValue(i) = Max(y * (OptionValue(i) - X2), OptionValue(i))
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
        BinomialCompoundOption = ReturnValue(0)
    ElseIf OutputFLag = "d" Then 'Delta
        BinomialCompoundOption = ReturnValue(1)
    ElseIf OutputFLag = "g" Then 'Gamma
        BinomialCompoundOption = ReturnValue(2)
    ElseIf OutputFLag = "t" Then 'Theta
        BinomialCompoundOption = ReturnValue(3)
   ElseIf OutputFLag = "a" Then
        BinomialCompoundOption = Application.Transpose(ReturnValue())
    End If
End Function


'// Cox-Ross-Rubinstein binomial tree
Public Function BinomialCompoundExoticOption(OutputFLag As String, AmeEurFlag As String, TypeFlag As String, UnderlyingOptionType As Integer, S As Double, X1 As Double, X2 As Double, t1 As Double, T2 As Double, _
                r As Double, b As Double, v As Double, pow As Double, cap As Double, n As Integer) As Variant
                

    Dim OptionValue() As Double
    Dim u As Double, d As Double, p As Double
    Dim ReturnValue(3) As Double
    Dim dt As Double, Df As Double, St As Double
    Dim i As Integer, j As Integer, z As Integer, y As Integer, w As Integer
    
    ReDim OptionValue(0 To n + 1)
    
    If TypeFlag = "cc" Or TypeFlag = "pc" Then
        z = 1
    Else
        z = -1
    End If
    
    If TypeFlag = "cc" Or TypeFlag = "cp" Then
        y = 1
    Else
        y = -1
    End If
 
    w = 1
    dt = T2 / n
    u = Exp(v * Sqr(dt))
    d = 1 / u
    p = (Exp(b * dt) - d) / (u - d)
    Df = Exp(-r * dt)
    
    For i = 0 To n
        St = S * u ^ i * d ^ (n - i)
         OptionValue(i) = BinomialPayoff(UnderlyingOptionType, z, St, X1, pow, cap)
    Next
    
    For j = n - 1 To 0 Step -1
        For i = 0 To j
          
            OptionValue(i) = (p * OptionValue(i + 1) + (1 - p) * OptionValue(i)) * Df
            
            If AmeEurFlag = "a" Then
                St = S * u ^ i * d ^ (j - i)
                OptionValue(i) = Max(BinomialPayoff(UnderlyingOptionType, z, St, X1, pow, cap), OptionValue(i))
            End If
            
            If t1 >= dt * j And w = 1 Then
                    OptionValue(i) = Max(y * (OptionValue(i) - X2), 0)
                    If i = j Then
                        w = -1
                    End If
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
        BinomialCompoundExoticOption = ReturnValue(0)
    ElseIf OutputFLag = "d" Then 'Delta
        BinomialCompoundExoticOption = ReturnValue(1)
    ElseIf OutputFLag = "g" Then 'Gamma
        BinomialCompoundExoticOption = ReturnValue(2)
    ElseIf OutputFLag = "t" Then 'Theta
        BinomialCompoundExoticOption = ReturnValue(3)
   ElseIf OutputFLag = "a" Then
        BinomialCompoundExoticOption = Application.Transpose(ReturnValue())
    End If
End Function

Public Function EuropeanBinomial(TypeFlag As Integer, CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, b As Double, v As Double, pow As Double, cap As Double, n As Long) As Double
        
    Dim u As Double, d As Double, p As Double
    Dim Sum As Double, dt As Double, Si As Double, z As Integer
    Dim j As Long
        
    z = 1
    If CallPutFlag = "p" Then z = -1
        
    dt = T / n
    u = Exp(v * Sqr(dt))
    d = 1 / u
    p = (Exp(b * dt) - d) / (u - d)

    Sum = 0
   
    For j = 0 To n
            Si = S * u ^ j * d ^ (n - j)
            Sum = Sum + Application.Combin(n, j) * p ^ j * (1 - p) ^ (n - j) * BinomialPayoff(TypeFlag, z, Si, X, pow, cap)
    Next
    
    EuropeanBinomial = Exp(-r * T) * Sum
    
End Function



Public Function BinomialPayoff(TypeFlag As Integer, z As Integer, S As Double, X As Double, pow As Double, cap As Double) As Double

    If TypeFlag = 1 Then 'Plain Vanilla
        BinomialPayoff = Max(z * (S - X), 0)
      ElseIf TypeFlag = 2 Then ' Power contract
        BinomialPayoff = S ^ pow
     ElseIf TypeFlag = 3 Then ' Capped Power contract
        BinomialPayoff = Application.Min(S ^ pow, cap)
     ElseIf TypeFlag = 4 Then ' Power contract
        BinomialPayoff = (S / X) ^ pow
    ElseIf TypeFlag = 5 Then ' Power contract
        BinomialPayoff = z * (S - X) ^ pow
    ElseIf TypeFlag = 6 Then 'Standard power option
        BinomialPayoff = Max(z * (S ^ pow - X), 0)
     ElseIf TypeFlag = 7 Then 'Capped power option
        BinomialPayoff = Application.Min(Max(z * (S ^ pow - X), 0), cap)
    ElseIf TypeFlag = 8 Then ' Powered option
            BinomialPayoff = Max((z * (S - X)), 0) ^ pow
    ElseIf TypeFlag = 9 Then ' Capped powered option
            BinomialPayoff = Application.Min(Max((z * (S - X)), 0) ^ pow, cap)
     ElseIf TypeFlag = 10 Then ' Sinus option
        BinomialPayoff = Max(z * (Sin(S) - X), 0)
     ElseIf TypeFlag = 11 Then ' Cosinus option
        BinomialPayoff = Max(z * (Cos(S) - X), 0)
      ElseIf TypeFlag = 12 Then ' Tangens option
        BinomialPayoff = Max(z * (Tan(S) - X), 0)
     ElseIf TypeFlag = 13 Then ' Log contract
        BinomialPayoff = Log(S)
     ElseIf TypeFlag = 14 Then ' Log contract
        BinomialPayoff = Log(S / X)
     ElseIf TypeFlag = 15 Then ' Log option
        BinomialPayoff = Max(Log(S / X), 0)
     ElseIf TypeFlag = 16 Then 'Square root contract
        BinomialPayoff = Sqr(S)
    ElseIf TypeFlag = 17 Then 'Square root contract
        BinomialPayoff = Sqr(S / X)
    ElseIf TypeFlag = 18 Then 'Square root option
        BinomialPayoff = Sqr(Max(z * (S - X), 0))
    End If
End Function
