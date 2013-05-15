Attribute VB_Name = "TreeModels"
Option Explicit
Option Base 0
Option Compare Text
Global Const Pi = 3.14159265358979

' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug

Public Function CRRBinomial(AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, n As Integer) As Double
                

    Dim OptionValue() As Double
    Dim U As Double, d As Double, P As Double
    Dim dt As Double, Df As Double
    Dim i As Integer, j As Integer, z As Integer
    
    ReDim OptionValue(0 To n + 1)
    
    'DefinererBinaryVariable
    If CallPutFlag = "c" Then
        z = 1
        ElseIf CallPutFlag = "p" Then
        z = -1
    End If
    
    dt = T / n
    U = Exp(v * Sqr(dt))
    d = 1 / U
    P = (Exp(b * dt) - d) / (U - d)
    Df = Exp(-r * dt)
    
    For i = 0 To n
         OptionValue(i) = Max(0, z * (S * U ^ i * d ^ (n - i) - X))
    Next
    
    For j = n - 1 To 0 Step -1:
        For i = 0 To j
            If AmeEurFlag = "e" Then
                OptionValue(i) = (P * OptionValue(i + 1) + (1 - P) * OptionValue(i)) * Df
            ElseIf AmeEurFlag = "a" Then
                OptionValue(i) = Max((z * (S * U ^ i * d ^ (Abs(i - j)) - X)), _
                (P * OptionValue(i + 1) + (1 - P) * OptionValue(i)) * Df)
            End If
        Next
    Next
    CRRBinomial = OptionValue(0)
End Function


Public Function DiscreteDividendYield(AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, v As Double, n As Integer, DividendTimes As Object, Dividends As Object) As Variant
    
    Dim ReturnValue() As Double
    Dim StepsDividend() As Double
    Dim St() As Double
    Dim OptionValue() As Double       'Option Value at each node
    Dim i As Integer, j As Integer, m As Integer, z As Integer
    Dim nDividends As Integer
    Dim Df As Double, dt As Double, P As Double
    Dim U As Double, uu As Double, d As Double, SumDividends As Double
    
    nDividends = Application.Count(DividendTimes)
    
    If nDividends = 0 Then
        DiscreteDividendYield = CRRBinomial(AmeEurFlag, CallPutFlag, S, X, T, r, r, v, n)
        Exit Function
    End If
    
    ReDim ReturnValue(0 To 3)
    ReDim StepsDividend(0 To nDividends)
    ReDim St(0 To n + 2)
    ReDim OptionValue(0 To n + 2)
    
    dt = T / n  '//Size of time step
    Df = Exp(-r * (T / n))  '// Discount factor
    U = Exp(v * Sqr(T / n))
     d = 1 / U
    uu = U ^ 2
    P = (Exp(r * dt) - d) / (U - d)  ' // Up probability

    z = 1 '// call
    If CallPutFlag = "p" Then
        z = -1 '// put
    End If
    
    SumDividends = 1
    
    For i = 0 To nDividends - 1
        StepsDividend(i) = Int(DividendTimes(i + 1) / T * n)
        SumDividends = SumDividends * (1 - Dividends(i + 1))
    Next
    
    For i = 0 To n '// Option value at expiry
        St(i) = S * U ^ i * d ^ (n - i) * SumDividends
        OptionValue(i) = Max(z * (St(i) - X), 0)
    Next

    For j = n - 1 To 0 Step -1
        For m = 0 To nDividends
            If j = StepsDividend(m) Then
                For i = 0 To j
                    St(i) = St(i) / (1 - Dividends(m + 1))
                Next i
            End If
        Next m
        For i = 0 To j
            St(i) = d * St(i + 1)
            OptionValue(i) = (P * OptionValue(i + 1) + (1 - P) * OptionValue(i)) * Df '//European value
           If AmeEurFlag = "a" Then '// American value
                OptionValue(i) = Max(OptionValue(i), z * (St(i) - X))
           End If
        Next i
        
          If j = 2 Then
          '// Gamma
                ReturnValue(2) = ((OptionValue(2) - OptionValue(1)) / (S * U ^ 2 - S) _
                        - (OptionValue(1) - OptionValue(0)) / (S - S * d ^ 2)) / (0.5 * (S * U ^ 2 - S * d ^ 2))
          '// Part of theta
               ReturnValue(3) = OptionValue(1)
            End If
         
         If j = 1 Then
         '// Delta
               ReturnValue(1) = (OptionValue(1) - OptionValue(0)) / (S * U - S * d)
            End If
            
    Next j
    
    ReturnValue(0) = OptionValue(0)
    ReturnValue(3) = (ReturnValue(3) - OptionValue(0)) / (2 * dt) / 365 ' // One day theta
   DiscreteDividendYield = Application.Transpose(ReturnValue())
   
End Function



'Non recombining binomial model with discrete cash dividends
Public Function BinomialDiscreteDividends(CallPutFlag As String, AmeEurFlag As String, S As Double, X As Double, T As Double, r As Double, v As Double, n As Integer, Optional CashDividends As Variant, Optional DividendTimes As Variant)
        
    Dim TmpDividendTimes() As Variant
    Dim TmpCashDividends() As Variant
    Dim StockPriceNode() As Double
    Dim OptionValueNode() As Double
    Dim NoOfDividends As Integer, Binary As Integer
    Dim Df As Double, dt As Double
    Dim U As Double, d As Double, uu As Double
    Dim P As Double, z As Double
    Dim i As Integer, j As Integer
    Dim StepsBeforeDividend As Integer
    Dim DividendAmount As Double
    Dim ValueNotExercising As Double
    
    If IsMissing(DividendTimes) Or IsEmpty(DividendTimes) Then
        NoOfDividends = 0
    Else
        ' // Counts the number of dividend payments
        NoOfDividends = Application.Count(DividendTimes)
    End If
            
    If NoOfDividends = 0 Then
        ' // If the number of dividends is zero use standard binomial model
        BinomialDiscreteDividends = CRRBinomial(AmeEurFlag, CallPutFlag, S, X, T, r, r, v, n)
        Exit Function
    End If
    
    ReDim TmpDividendTimes(1 To NoOfDividends) As Variant
    ReDim TmpCashDividends(1 To NoOfDividends) As Variant
         
    If CallPutFlag = "c" Then
        Binary = 1  ' // call option
    ElseIf CallPutFlag = "p" Then
        Binary = -1 ' // put option
    End If
    
    dt = T / n
    Df = Exp(-r * dt)
    U = Exp(v * Sqr(dt))
    d = 1 / U
    uu = U ^ 2
    P = (Exp(r * dt) - d) / (U - d)
    
    DividendAmount = CashDividends(1)
    
    For i = 1 To NoOfDividends - 1 Step 1
        TmpCashDividends(i) = CashDividends(i + 1)
        TmpDividendTimes(i) = DividendTimes(i + 1) - DividendTimes(1)
    Next
    
    StepsBeforeDividend = Int(DividendTimes(1) / T * n)
    
    ReDim StockPriceNode(1 To StepsBeforeDividend + 2) As Double
    ReDim OptionValueNode(1 To StepsBeforeDividend + 2) As Double
        
    StockPriceNode(1) = S * d ^ StepsBeforeDividend
    
    For i = 2 To StepsBeforeDividend + 1 Step 1
        StockPriceNode(i) = StockPriceNode(i - 1) * uu
    Next
    
    '// Calculate option values for all nodes at time step just before dividend payment
    For i = 1 To StepsBeforeDividend + 1 Step 1
        '// Because non recombining the model need to build a new binomial tree from every singel node at this time step
        ValueNotExercising = BinomialDiscreteDividends(CallPutFlag, AmeEurFlag, StockPriceNode(i) - DividendAmount, X, T - DividendTimes(1), r, v, n - StepsBeforeDividend, TmpCashDividends, TmpDividendTimes)
         If AmeEurFlag = "a" Then
            OptionValueNode(i) = Max(ValueNotExercising, Binary * (StockPriceNode(i) - X))
         ElseIf AmeEurFlag = "e" Then
            OptionValueNode(i) = ValueNotExercising
        End If
    Next

    '//Option values before dividend payment "standard binomial"
    For j = StepsBeforeDividend To 1 Step -1
        For i = 1 To j + 1 Step 1
            StockPriceNode(i) = d * StockPriceNode(i + 1)
            If AmeEurFlag = "a" Then
                OptionValueNode(i) = Max((P * OptionValueNode(i + 1) + (1 - P) * OptionValueNode(i)) * Df, Binary * (StockPriceNode(i) - X))
            ElseIf AmeEurFlag = "e" Then
                    OptionValueNode(i) = (P * OptionValueNode(i + 1) + (1 - P) * OptionValueNode(i)) * Df
            End If
        Next
    Next
    
    BinomialDiscreteDividends = OptionValueNode(1)
    
End Function
