Attribute VB_Name = "ImpliedTree"
Option Explicit         'Requires that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.

Option Base 0       'The "Option Base" statement allows to specify 0 or 1 as the
                            'default first index of arrays.
               
' Programmer Espen Gaarder Haug
' Copyright 2006 Espen Gaarder Haug
               
               
Public Function Max(X, y)
Attribute Max.VB_ProcData.VB_Invoke_Func = " \r14"
            Max = Application.Max(X, y)
End Function

Public Function Min(X, y)
Attribute Min.VB_ProcData.VB_Invoke_Func = " \r14"
            Min = Application.Min(X, y)
End Function


'// Trinomial tree
Public Function TrinomialTree(AmeEurFlag As String, CallPutFlag As String, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, n As Integer) As Double
Attribute TrinomialTree.VB_ProcData.VB_Invoke_Func = " \r14"
                

    Dim OptionValue() As Double
    Dim dt As Double, u As Double, d As Double
    Dim pu As Double, pd As Double, pm As Double
    Dim i As Integer, j As Integer, z As Integer
    Dim Df As Double
    
    ReDim OptionValue(0 To n * 2 + 1)
    
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
            If AmeEurFlag = "e" Then
                OptionValue(i) = (pu * OptionValue(i + 2) + pm * OptionValue(i + 1) + pd * OptionValue(i)) * Df
            ElseIf AmeEurFlag = "a" Then
                OptionValue(i) = Max((z * (S * u ^ Max(i - j, 0) * d ^ Max(j - i, 0) - X)), _
                (pu * OptionValue(i + 2) + pm * OptionValue(i + 1) + pd * OptionValue(i)) * Df)
            End If
        Next
    Next
    TrinomialTree = OptionValue(0)
    
End Function



'// Implied trinomial tree
Public Function ImpliedTrinomialTree(ReturnFlag As String, STEPn As Integer, STATEi As Integer, S As Double, X As Double, T As Double, _
                r As Double, b As Double, v As Double, Skew As Double, nSteps As Integer)
Attribute ImpliedTrinomialTree.VB_ProcData.VB_Invoke_Func = " \r14"
                
    Dim ArrowDebreu() As Double
    Dim LocalVolatility() As Double
    Dim UpProbability() As Double
    Dim DownProbability() As Double
    Dim OptionValueNode() As Double
    Dim dt As Double, u As Double, d As Double
    Dim Df As Double, pi As Double, qi As Double
    Dim Si1 As Double, Si As Double, Si2 As Double
    Dim vi As Double, Fj As Double, Fi As Double, Fo As Double
    Dim sum As Double, OptionValue As Double
    Dim i As Integer, j As Integer, n As Integer, z As Integer
   
    ReDim OptionValueNode(0 To nSteps * 2) As Double
    ReDim ArrowDebreu(0 To nSteps, 0 To nSteps * 2) As Double
    ReDim UpProbability(0 To nSteps - 1, 0 To nSteps * 2 - 2) As Double
    ReDim DownProbability(0 To nSteps - 1, 0 To nSteps * 2 - 2) As Double
    ReDim LocalVolatility(0 To nSteps - 1, 0 To nSteps * 2 - 2) As Double
   
    dt = T / nSteps
    u = Exp(v * Sqr(2 * dt))
    d = 1 / u
    Df = Exp(-r * dt)
    ArrowDebreu(0, 0) = 1
    For n = 0 To nSteps - 1
        For i = 0 To n * 2
            sum = 0
            Si1 = S * u ^ Max(i - n, 0) * d ^ Max(n * 2 - n - i, 0)
            Si = Si1 * d
            Si2 = Si1 * u
            Fi = Si1 * Exp(b * dt)
            vi = v + (S - Si1) * Skew
            If i < (n * 2) / 2 + 1 Then
                For j = 0 To i - 1
                    Fj = S * u ^ Max(j - n, 0) * d ^ Max(n * 2 - n - j, 0) * Exp(b * dt)
                    sum = sum + ArrowDebreu(n, j) * (Si1 - Fj)
                Next
                OptionValue = TrinomialTree("e", "p", S, Si1, (n + 1) * dt, r, b, vi, n + 1)
                qi = (Exp(r * dt) * OptionValue - sum) / (ArrowDebreu(n, i) * (Si1 - Si))
                pi = (Fi + qi * (Si1 - Si) - Si1) / (Si2 - Si1)
            Else
                OptionValue = TrinomialTree("e", "c", S, Si1, (n + 1) * dt, r, b, vi, n + 1)
                sum = 0
                For j = i + 1 To n * 2
                    Fj = S * u ^ Max(j - n, 0) * d ^ Max(n * 2 - n - j, 0) * Exp(b * dt)
                    sum = sum + ArrowDebreu(n, j) * (Fj - Si1)
                Next
                pi = (Exp(r * dt) * OptionValue - sum) / (ArrowDebreu(n, i) * (Si2 - Si1))
                qi = (Fi - pi * (Si2 - Si1) - Si1) / (Si - Si1)
            End If
            
            '// Replacing negative probabilities
            If pi < 0 Or pi > 1 Or qi < 0 Or qi > 1 Then
                If Fi > Si1 And Fi < Si2 Then
                    pi = 1 / 2 * ((Fi - Si1) / (Si2 - Si1) + (Fi - Si) / (Si2 - Si))
                    qi = 1 / 2 * ((Si2 - Fi) / (Si2 - Si))
                ElseIf Fi > Si And Fi < Si1 Then
                    pi = 1 / 2 * ((Fi - Si) / (Si2 - Si))
                    qi = 1 / 2 * ((Si2 - Fi) / (Si2 - Si) + (Si1 - Fi) / (Si1 - Si))
                End If
            End If
                DownProbability(n, i) = qi
                UpProbability(n, i) = pi
            '// Calculating local volatilities
                Fo = pi * Si2 + qi * Si + (1 - pi - qi) * Si1
                LocalVolatility(n, i) = Sqr((pi * (Si2 - Fo) ^ 2 + (1 - pi - qi) * _
                (Si1 - Fo) ^ 2 + qi * (Si - Fo) ^ 2) / (Fo ^ 2 * dt))

            '// Calculating Arrow-Debreu prices
            If n = 0 Then
                ArrowDebreu(n + 1, i) = qi * ArrowDebreu(n, i) * Df
                ArrowDebreu(n + 1, i + 1) = (1 - pi - qi) * ArrowDebreu(n, i) * Df
                ArrowDebreu(n + 1, i + 2) = pi * ArrowDebreu(n, i) * Df
            ElseIf n > 0 And i = 0 Then
                ArrowDebreu(n + 1, i) = qi * ArrowDebreu(n, i) * Df
            ElseIf n > 0 And i = n * 2 Then
                ArrowDebreu(n + 1, i) = UpProbability(n, i - 2) * ArrowDebreu(n, i - 2) * Df _
                                + (1 - UpProbability(n, i - 1) - DownProbability(n, i - 1)) * ArrowDebreu(n, i - 1) * Df _
                                + qi * ArrowDebreu(n, i) * Df
                ArrowDebreu(n + 1, i + 1) = UpProbability(n, i - 1) * ArrowDebreu(n, i - 1) * Df + (1 - pi - qi) * ArrowDebreu(n, i) * Df
                ArrowDebreu(n + 1, i + 2) = pi * ArrowDebreu(n, i) * Df
            ElseIf n > 0 And i = 1 Then
                ArrowDebreu(n + 1, i) = (1 - UpProbability(n, i - 1) - DownProbability(n, i - 1)) * ArrowDebreu(n, i - 1) * Df _
                                + qi * ArrowDebreu(n, i) * Df
            Else
            ArrowDebreu(n + 1, i) = UpProbability(n, i - 2) * ArrowDebreu(n, i - 2) * Df _
                                + (1 - UpProbability(n, i - 1) - DownProbability(n, i - 1)) * ArrowDebreu(n, i - 1) * Df _
                                + qi * ArrowDebreu(n, i) * Df
        End If
        Next
    Next
        
    If ReturnFlag = "DPM" Then
        ImpliedTrinomialTree = Application.Transpose(DownProbability)
    ElseIf ReturnFlag = "UPM" Then
        ImpliedTrinomialTree = Application.Transpose(UpProbability)
    ElseIf ReturnFlag = "DPni" Then
        ImpliedTrinomialTree = (DownProbability(STEPn, STATEi))
    ElseIf ReturnFlag = "UPni" Then
      ImpliedTrinomialTree = (UpProbability(STEPn, STATEi))
    ElseIf ReturnFlag = "ADM" Then
        ImpliedTrinomialTree = Application.Transpose(ArrowDebreu)
    ElseIf ReturnFlag = "LVM" Then
        ImpliedTrinomialTree = Application.Transpose(LocalVolatility)
    ElseIf ReturnFlag = "LVni" Then
        ImpliedTrinomialTree = Application.Transpose(LocalVolatility(STEPn, STATEi))
    ElseIf ReturnFlag = "ADni" Then
        ImpliedTrinomialTree = (ArrowDebreu(STEPn, STATEi))
    Else
    
    '// Calculation of option price using the implied trinomial tree
        If ReturnFlag = "c" Then
            z = 1
        ElseIf ReturnFlag = "p" Then
            z = -1
        End If
        For i = 0 To (2 * nSteps)
                OptionValueNode(i) = Max(0, z * (S * u ^ Max(i - nSteps, 0) * d ^ Max(nSteps - i, 0) - X))
        Next
         For n = nSteps - 1 To 0 Step -1
            For i = 0 To (n * 2)
                OptionValueNode(i) = (UpProbability(n, i) * OptionValueNode(i + 2) + (1 - UpProbability(n, i) - DownProbability(n, i)) * OptionValueNode(i + 1) + DownProbability(n, i) * OptionValueNode(i)) * Df
            Next
        Next
        ImpliedTrinomialTree = OptionValueNode(0)
    End If
End Function
