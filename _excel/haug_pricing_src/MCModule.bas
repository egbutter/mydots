Attribute VB_Name = "MCModule"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.

Function Max(m1 As Double, m2 As Double) As Double
    Max = Application.WorksheetFunction.Max(m1, m2)
End Function
Function NormInv(n1 As Double, n2 As Double, n3 As Double) As Double
    NormInv = Application.NormInv(n1, n2, n3)
End Function

' Monte Carlo plain vanilla American option, Broadie and Glasserman (1997)
Public Function BroadieGlasserman(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, _
        b As Double, Sig As Double, m As Integer, Branches As Integer, nSimulations As Integer) As Double
    
    '  Based on codes supplied by Silvan G.R. Meier
        
    Dim Drift As Double, SigSqrdt As Double, Discdt As Double, z As Integer
    Dim i As Integer, j As Integer, i1 As Integer, i2 As Integer, Estimator As Integer, Simulation As Integer
    Dim EstimatorSum As Double, Sum1 As Double, Sum2 As Double
    Dim w() As Integer, v() As Double, Estimators() As Double
    ReDim w(1 To m) As Integer, v(1 To Branches, 1 To m) As Double, Estimators(1 To 2) As Double
    
    z = 1
    If CallPutFlag = "Put" Then z = -1
    
    Drift = (b - Sig ^ 2 / 2) * T / (m - 1)
    SigSqrdt = Sig * Sqr(T / (m - 1))
    Discdt = Exp(-r * T / (m - 1))
    
    For Estimator = 1 To 2
        EstimatorSum = 0
        For Simulation = 1 To nSimulations
            v(1, 1) = S
            w(1) = 1
            For j = 2 To m
                v(1, j) = v(1, j - 1) * Exp(Drift + SigSqrdt * NormInv(Rnd, 0, 1))
                w(j) = 1
            Next j
            j = m
            Do While j > 0
                If j = m Then
                    v(w(j), j) = Max(z * (v(w(j), j) - X), 0)
                    If w(j) < Branches Then
                        v(w(j) + 1, j) = v(w(j - 1), j - 1) * Exp(Drift + SigSqrdt * NormInv(Rnd, 0, 1))
                        w(j) = w(j) + 1
                    ElseIf w(j) = Branches Then
                        w(j) = 0
                        j = j - 1
                    End If
                ElseIf j < m Then
                    If Estimator = 1 Then 'the high estimator
                        Sum1 = 0
                        For i1 = 1 To Branches
                            Sum1 = Sum1 + Discdt * v(i1, j + 1)
                        Next i1
                        v(w(j), j) = Max(Max(z * (v(w(j), j) - X), 0), Sum1 / Branches)
                    ElseIf Estimator = 2 Then 'the low estimator
                        Sum1 = 0 'value determinant
                        For i1 = 1 To Branches
                            Sum2 = 0 'decision determinant
                            For i2 = 1 To Branches
                                If i1 <> i2 Then Sum2 = Sum2 + Discdt * v(i2, j + 1)
                            Next i2
                            If Max(z * (v(w(j), j) - X), 0) >= Sum2 / (Branches - 1) Then
                                Sum1 = Sum1 + Max(z * (v(w(j), j) - X), 0)
                            Else
                                Sum1 = Sum1 + Discdt * v(i1, j + 1)
                            End If
                        Next i1
                        v(w(j), j) = Sum1 / Branches
                    End If
                    If w(j) < Branches Then
                        If j > 1 Then
                            v(w(j) + 1, j) = v(w(j - 1), j - 1) * Exp(Drift + SigSqrdt * NormInv(Rnd, 0, 1))
                            w(j) = w(j) + 1
                            For i = j + 1 To m
                                v(1, i) = v(w(j), j) * Exp(Drift + SigSqrdt * NormInv(Rnd, 0, 1))
                                w(i) = 1
                            Next i
                            j = m
                        Else
                            j = 0
                        End If
                     ElseIf w(j) = Branches Then
                        w(j) = 0
                        j = j - 1
                    End If
                End If
            Loop
            EstimatorSum = EstimatorSum + v(1, 1)
        Next Simulation
        Estimators(Estimator) = EstimatorSum / nSimulations
    Next Estimator
    BroadieGlasserman = 0.5 * Max(Max(z * (S - X), 0), Estimators(2)) + 0.5 * Estimators(1)
End Function


' Monte Carlo plain vanilla American option, Broadie and Glasserman (1997)
Function BroadieGlassermanOriginal(CallPutFlag As String, S As Double, X As Double, T As Double, r As Double, _
        Del As Double, Sig As Double, d As Integer, b As Integer, nSimulations As Integer) As Double
        
    Dim Drift As Double, SigSqrdt As Double, Discdt As Double, z As Integer
    Dim i As Integer, j As Integer, i1 As Integer, i2 As Integer, Estimator As Integer, Simulation As Integer
    Dim EstimatorSum As Double, Sum1 As Double, Sum2 As Double
    Dim w() As Integer, v() As Double, Estimators() As Double
    ReDim w(1 To d) As Integer, v(1 To b, 1 To d) As Double, Estimators(1 To 2) As Double
    If CallPutFlag = "Put" Then z = -1 Else z = 1
    Drift = (r - Del - Sig ^ 2 / 2) * T / (d - 1)
    SigSqrdt = Sig * Sqr(T / (d - 1))
    Discdt = Exp(-r * T / (d - 1))
    For Estimator = 1 To 2
        EstimatorSum = 0
        For Simulation = 1 To nSimulations
            v(1, 1) = S: w(1) = 1
            For j = 2 To d
                v(1, j) = v(1, j - 1) * Exp(Drift + SigSqrdt * NormInv(Rnd, 0, 1))
                w(j) = 1
            Next
            j = d
            Do While j > 0
                If j = d Then
                    v(w(j), j) = Max(z * (v(w(j), j) - X), 0)
                    If w(j) < b Then
                        v(w(j) + 1, j) = v(w(j - 1), j - 1) * Exp(Drift + SigSqrdt * NormInv(Rnd, 0, 1))
                        w(j) = w(j) + 1
                    ElseIf w(j) = b Then
                        w(j) = 0
                        j = j - 1
                    End If
                ElseIf j < d Then
                    If Estimator = 1 Then 'the high estimator
                        Sum1 = 0: For i1 = 1 To b: Sum1 = Sum1 + Discdt * v(i1, j + 1): Next i1
                        v(w(j), j) = Max(Max(z * (v(w(j), j) - X), 0), Sum1 / b)
                    ElseIf Estimator = 2 Then 'the low estimator
                        Sum1 = 0 'value determinant
                        For i1 = 1 To b
                            Sum2 = 0 'decision determinant
                            For i2 = 1 To b
                                If i1 <> i2 Then Sum2 = Sum2 + Discdt * v(i2, j + 1)
                            Next i2
                            If Max(z * (v(w(j), j) - X), 0) >= Sum2 / (b - 1) Then
                                Sum1 = Sum1 + Max(z * (v(w(j), j) - X), 0)
                            Else
                                Sum1 = Sum1 + Discdt * v(i1, j + 1)
                            End If
                        Next i1
                        v(w(j), j) = Sum1 / b
                    End If
                    If w(j) < b Then
                        If j > 1 Then
                            v(w(j) + 1, j) = v(w(j - 1), j - 1) * Exp(Drift + SigSqrdt * NormInv(Rnd, 0, 1))
                            w(j) = w(j) + 1
                            For i = j + 1 To d
                                v(1, i) = v(w(j), j) * Exp(Drift + SigSqrdt * NormInv(Rnd, 0, 1))
                                w(i) = 1
                            Next
                            j = d
                        Else
                            j = 0
                        End If
                     ElseIf w(j) = b Then
                        w(j) = 0
                        j = j - 1
                    End If
                End If
            Loop
            EstimatorSum = EstimatorSum + v(1, 1)
        Next Simulation
        Estimators(Estimator) = EstimatorSum / nSimulations
    Next Estimator
    BroadieGlassermanOriginal = 0.5 * Max(Max(z * (S - X), 0), Estimators(2)) + 0.5 * Estimators(1)
End Function
