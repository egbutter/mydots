Attribute VB_Name = "Interpolation"
Option Explicit
Option Base 0


' Programmer Espen Gaarder Haug
' Copyright 2006, Espen Gaarder Haug

'//Cubic Interpolation Function
Private Function CubicInterpolation(XVector As Object, YVector As Object, X As Double) As Double
    
    Dim k As Double
    Dim ytmp As Double
    Dim Y As Double
    Dim i As Integer, j As Integer
    
    k = Application.Match(X, XVector)
    If k < 2 Then
         k = 2
    ElseIf k > XVector.Rows.Count - 2 Then
        k = XVector.Rows.Count - 2
    End If
    Y = 0
    For i = k - 1 To k + 2
        ytmp = 1
        For j = k - 1 To k + 2
            If i <> j Then
                ytmp = ytmp * (X - Application.Index(XVector, j)) / (Application.Index(XVector, i) - Application.Index(XVector, j))
            End If
        Next
        Y = Y + ytmp * Application.Index(YVector, i)
    Next
    CubicInterpolation = Y
End Function

Public Function EInterpolation(XVector As Object, YVector As Object, X As Double, InterpolationMethod As Integer) As Double
        
    Dim X1 As Double, X2 As Double
    Dim r1 As Double, r2 As Double
    Dim T1 As Double, T2 As Double
    
    X1 = Application.Match(X, XVector)
        
        X2 = X1 + 1
    T1 = Application.Index(XVector, X1)
    T2 = Application.Index(XVector, X2)
    r1 = Application.Index(YVector, X1)
    r2 = Application.Index(YVector, X2)
   
    
    If InterpolationMethod = 1 Then
        EInterpolation = (r2 - r1) * (X - T1) / (T2 - T1) + r1
    ElseIf InterpolationMethod = 2 Then
        EInterpolation = (r2 / r1) ^ ((X - T1) / (T2 - T1)) * r1
    ElseIf InterpolationMethod = 3 Then
        EInterpolation = CubicInterpolation(XVector, YVector, X)
    End If
End Function


Function CubicSpline(XArray As Variant, YArray As Variant, X As Variant)
    
    Dim nRates As Integer, nn As Integer
    Dim i As Integer, j As Integer
    Dim ArrayNo As Integer
    Dim ti As Double, y1 As Double, ai As Double
    Dim bi As Double, ci As Double

    nRates = Application.Count(XArray) - 1
    
    Dim M() As Variant, N() As Variant
    Dim Alfa() As Variant, Beta() As Variant, Delta() As Variant
    Dim Q() As Variant
    Dim A() As Variant, B() As Variant, C() As Variant
    
    ReDim M(0 To nRates + 1)
    ReDim N(0 To nRates + 1)
    ReDim Alfa(0 To nRates + 1)
    ReDim Beta(0 To nRates + 1)
    ReDim Delta(0 To nRates + 1)
    ReDim Q(0 To nRates + 1)
    ReDim A(0 To nRates + 1)
    ReDim B(0 To nRates + 1)
    ReDim C(0 To nRates + 1)
     
    For i = 0 To nRates - 1
        M(i) = XArray(i + 2) - XArray(i + 1)
        N(i) = YArray(i + 2) - YArray(i + 1)
    Next
    For i = 1 To nRates - 1
        Q(i) = 3 * (N(i) / M(i) - N(i - 1) / M(i - 1))
    Next
        Alfa(0) = 1
        Beta(0) = 0
        Delta(0) = 0
    For i = 1 To nRates - 1
        Alfa(i) = 2 * (M(i - 1) + M(i)) - M(i - 1) * Beta(i - 1)
        Beta(i) = M(i) / Alfa(i)
        Delta(i) = (Q(i) - M(i - 1) * Delta(i - 1)) / Alfa(i)
    Next
    Alfa(nRates) = 0
    B(nRates) = 0
    Delta(nRates) = 0
    For j = (nRates - 1) To 0 Step -1
        B(j) = Delta(j) - Beta(j) * B(j + 1)
        A(j) = N(j) / M(j) - M(j) / 3 * (B(j + 1) + 2 * B(j))
        C(j) = (B(j + 1) - B(j)) / (3 * M(j))
    Next
    nn = Application.Count(X)
    Dim z() As Double
    ReDim z(0 To nn + 1)
    For i = 1 To nn
        ArrayNo = Application.Match(X(i), XArray)
        ti = Application.Index(XArray, ArrayNo)
        y1 = Application.Index(YArray, ArrayNo)
        ai = Application.Index(A(), ArrayNo)
        bi = Application.Index(B(), ArrayNo)
        ci = Application.Index(C(), ArrayNo)
        z(i - 1) = y1 + ai * (X(i) - ti) + bi * (X(i) - ti) ^ 2 + ci * (X(i) - ti) ^ 3
    Next
    CubicSpline = Application.Transpose(z())
End Function
