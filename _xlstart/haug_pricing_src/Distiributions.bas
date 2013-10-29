Attribute VB_Name = "Distiributions"
Global Const Pi = 3.14159265358979

Option Explicit     'Requirs that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.

Option Base 0       'The "Option Base" statment alowws to specify 0 or 1 as the
                             'default first index of arrays.

' Programmer Espen Gaarder Haug, Copyright 2006

'// Convert rate to countinuous compounding rate
Public Function ConvertingToCCRate(r As Double, Compoundings As Double) As Double
    
    If Compoundings = 0 Then
        ConvertingToCCRate = r
    Else
        ConvertingToCCRate = Compoundings * Log(1 + r / Compoundings)
    End If
End Function

'// Inverse cummulative normal distribution function
Public Function CNDEV(U As Double) As Double
    
    Dim x As Double, r As Double
    Dim A As Variant, b As Variant, c As Variant
    
    A = Array(2.50662823884, -18.61500062529, 41.39119773534, -25.44106049637)
    b = Array(-8.4735109309, 23.08336743743, -21.06224101826, 3.13082909833)
    c = Array(0.337475482272615, 0.976169019091719, 0.160797971491821, 2.76438810333863E-02, 3.8405729373609E-03, 3.951896511919E-04, 3.21767881767818E-05, 2.888167364E-07, 3.960315187E-07)

    x = U - 0.5
    If Abs(x) < 0.92 Then
        r = x * x
        r = x * (((A(3) * r + A(2)) * r + A(1)) * r + A(0)) _
        / ((((b(3) * r + b(2)) * r + b(1)) * r + b(0)) * r + 1)
        CNDEV = r
        Exit Function
    End If
    r = U
    If x >= 0 Then r = 1 - U
    r = Log(-Log(r))
    r = c(0) + r * (c(1) + r * (c(2) + r * (c(3) + r + (c(4) + _
        r * (c(5) + r * (c(6) + r * (c(7) + r * c(8))))))))
    If x < 0 Then r = -r
    CNDEV = r
    
End Function
                    
'// The cumulative bivariate normal distribution function
'// Drezner-Wesolowsky 1990 simple algorithm
Public Function CBND2(A As Double, b As Double, rho As Double) As Double

   Dim g As Double, P As Double, x, y, sum As Double
   Dim i As Integer
    
    x = Array(0.018854042, 0.038088059, 0.0452707394, 0.038088059, 0.018854042)
    y = Array(0.04691008, 0.23076534, 0.5, 0.76923466, 0.95308992)
    
    sum = 0
    For i = 0 To 4
        P = y(i) * rho
        g = 1 - P ^ 2
        sum = sum + x(i) * Exp((2 * A * b * P - A ^ 2 - b ^ 2) / g / 2) / Sqr(g)
    Next
    CBND2 = rho * sum + CND(A) * CND(b)
End Function

Public Function Max(x, y)
            Max = Application.Max(x, y)
End Function

Public Function Min(x, y)
            Min = Application.Min(x, y)
End Function
                                                                                                         
'// The normal distribution function
Public Function ND(x As Double) As Double
    ND = 1 / Sqr(2 * Pi) * Exp(-x ^ 2 / 2)
End Function


'// Cummulative double precision algorithm based on Hart 1968
'// Based on implementation by Graeme West
Function CND(x As Double) As Double
    Dim y As Double, Exponential As Double, SumA As Double, SumB As Double
    
    y = Abs(x)
    If y > 37 Then
        CND = 0
    Else
        Exponential = Exp(-y ^ 2 / 2)
        If y < 7.07106781186547 Then
            SumA = 3.52624965998911E-02 * y + 0.700383064443688
            SumA = SumA * y + 6.37396220353165
            SumA = SumA * y + 33.912866078383
            SumA = SumA * y + 112.079291497871
            SumA = SumA * y + 221.213596169931
            SumA = SumA * y + 220.206867912376
            SumB = 8.83883476483184E-02 * y + 1.75566716318264
            SumB = SumB * y + 16.064177579207
            SumB = SumB * y + 86.7807322029461
            SumB = SumB * y + 296.564248779674
            SumB = SumB * y + 637.333633378831
            SumB = SumB * y + 793.826512519948
            SumB = SumB * y + 440.413735824752
            CND = Exponential * SumA / SumB
        Else
            SumA = y + 0.65
            SumA = y + 4 / SumA
            SumA = y + 3 / SumA
            SumA = y + 2 / SumA
            SumA = y + 1 / SumA
            CND = Exponential / (SumA * 2.506628274631)
        End If
  End If
  
  If x > 0 Then CND = 1 - CND

End Function


'// The cumulative normal distribution function
Public Function CND2(x As Double) As Double
    
    If x = 0 Then
     CND2 = 0.5
   Else
    
        Dim L As Double, k As Double
        Const a1 = 0.31938153:  Const a2 = -0.356563782: Const a3 = 1.781477937:
        Const a4 = -1.821255978:  Const a5 = 1.330274429
    
        L = Abs(x)
        k = 1 / (1 + 0.2316419 * L)
        CND2 = 1 - 1 / Sqr(2 * Pi) * Exp(-L ^ 2 / 2) * (a1 * k + a2 * k ^ 2 + a3 * k ^ 3 + a4 * k ^ 4 + a5 * k ^ 5)
    
        If x < 0 Then
            CND2 = 1 - CND2
        End If
    End If
    
End Function


Public Function CBND4(A As Double, b As Double, rho As Double) As Double
'modified/corrected from the second function in Drez & Wes paper pg. 105
'0/0 case resolved by l'H rule

  Dim i As Integer
  Dim x As Variant, W As Variant
  Dim h1 As Double, h2 As Double
  Dim LH As Double, h12 As Double, h3 As Double, h5 As Double, h6 As Double, h7 As Double, h8 As Double
  Dim r1 As Double, r2 As Double, r3 As Double, rr As Double
  Dim AA As Double, ab As Double
  
  x = Array(0.04691008, 0.23076534, 0.5, 0.76923466, 0.95308992)
  W = Array(0.018854042, 0.038088059, 0.0452707394, 0.038088059, 0.018854042)
  
  h1 = A
  h2 = b
  h12 = (h1 * h1 + h2 * h2) / 2
  
  If Abs(rho) >= 0.7 Then
    r2 = 1 - rho * rho
    r3 = Sqr(r2)
    If rho < 0 Then h2 = -h2
    h3 = h1 * h2
    h7 = Exp(-h3 / 2)
    If Abs(rho) < 1 Then
      h6 = Abs(h1 - h2)
      h5 = h6 * h6 / 2
      h6 = h6 / r3
      AA = 0.5 - h3 / 8
      ab = 3 - 2 * AA * h5
      LH = 0.13298076 * h6 * ab * (1 - CND(h6)) - Exp(-h5 / r2) * (ab + AA * r2) * 0.053051647
      For i = 0 To 4
        r1 = r3 * x(i)
        rr = r1 * r1
        r2 = Sqr(1 - rr)
        If h7 = 0 Then
          h8 = 0
        Else
          h8 = Exp(-h3 / (1 + r2)) / r2 / h7
        End If
        LH = LH - W(i) * Exp(-h5 / rr) * (h8 - 1 - AA * rr)
      Next i
    End If
    CBND4 = LH * r3 * h7 + CND(Min(h1, h2))
    If rho < 0 Then
      CBND4 = CND(h1) - CBND4
    End If
  Else
    h3 = h1 * h2
    If rho <> 0 Then
      For i = 0 To 4
        r1 = rho * x(i)
        r2 = 1 - r1 * r1
        LH = LH + W(i) * Exp((r1 * h3 - h12) / r2) / Sqr(r2)
      Next i
    End If
    CBND4 = CND(h1) * CND(h2) + rho * LH
  End If
     
End Function

Public Function CBNDGeneral(TypeFlag As Integer, x As Double, y As Double, rho As Double) As Double

    If TypeFlag = 1 Then 'Drezner-78
        CBNDGeneral = CBND2(x, y, rho)
    ElseIf TypeFlag = 2 Then 'Drezner-Weso-90a
        CBNDGeneral = CBND3(x, y, rho)
     ElseIf TypeFlag = 3 Then 'Drezner-Weso-90a
        CBNDGeneral = CBND4(x, y, rho)
     ElseIf TypeFlag = 4 Then ' Genze
        CBNDGeneral = CBND(x, y, rho)
    End If
    
End Function
'// The  bivariate normal distribution function
Public Function BND(x As Double, y As Double, rho As Double) As Double
    
    BND = 1 / (2 * Pi * Sqr(1 - rho ^ 2)) * Exp(-1 / (2 * (1 - rho ^ 2)) * (x ^ 2 + y ^ 2 - 2 * x * y * rho))

End Function

'// The cumulative bivariate normal distribution function
Public Function CBND(x As Double, y As Double, rho As Double) As Double
'
'     A function for computing bivariate normal probabilities.
'
'       Alan Genz
'       Department of Mathematics
'       Washington State University
'       Pullman, WA 99164-3113
'       Email : alangenz@wsu.edu
'
'    This function is based on the method described by
'        Drezner, Z and G.O. Wesolowsky, (1990),
'        On the computation of the bivariate normal integral,
'        Journal of Statist. Comput. Simul. 35, pp. 101-107,
'    with major modifications for double precision, and for |R| close to 1.
'   This code was originally transelated into VBA by Graeme West

Dim i As Integer, ISs As Integer, LG As Integer, NG As Integer
Dim XX(10, 3) As Double, W(10, 3) As Double
Dim h As Double, k As Double, hk As Double, hs As Double, BVN As Double, Ass As Double, asr As Double, sn As Double
Dim A As Double, b As Double, bs As Double, c As Double, d As Double
Dim xs As Double, rs As Double

W(1, 1) = 0.17132449237917
XX(1, 1) = -0.932469514203152
W(2, 1) = 0.360761573048138
XX(2, 1) = -0.661209386466265
W(3, 1) = 0.46791393457269
XX(3, 1) = -0.238619186083197

W(1, 2) = 4.71753363865118E-02
XX(1, 2) = -0.981560634246719
W(2, 2) = 0.106939325995318
XX(2, 2) = -0.904117256370475
W(3, 2) = 0.160078328543346
XX(3, 2) = -0.769902674194305
W(4, 2) = 0.203167426723066
XX(4, 2) = -0.587317954286617
W(5, 2) = 0.233492536538355
XX(5, 2) = -0.36783149899818
W(6, 2) = 0.249147045813403
XX(6, 2) = -0.125233408511469

W(1, 3) = 1.76140071391521E-02
XX(1, 3) = -0.993128599185095
W(2, 3) = 4.06014298003869E-02
XX(2, 3) = -0.963971927277914
W(3, 3) = 6.26720483341091E-02
XX(3, 3) = -0.912234428251326
W(4, 3) = 8.32767415767048E-02
XX(4, 3) = -0.839116971822219
W(5, 3) = 0.10193011981724
XX(5, 3) = -0.746331906460151
W(6, 3) = 0.118194531961518
XX(6, 3) = -0.636053680726515
W(7, 3) = 0.131688638449177
XX(7, 3) = -0.510867001950827
W(8, 3) = 0.142096109318382
XX(8, 3) = -0.37370608871542
W(9, 3) = 0.149172986472604
XX(9, 3) = -0.227785851141645
W(10, 3) = 0.152753387130726
XX(10, 3) = -7.65265211334973E-02
      
If Abs(rho) < 0.3 Then
  NG = 1
  LG = 3
ElseIf Abs(rho) < 0.75 Then
  NG = 2
  LG = 6
Else
  NG = 3
  LG = 10
End If
      
h = -x
k = -y
hk = h * k
BVN = 0
      
If Abs(rho) < 0.925 Then
  If Abs(rho) > 0 Then
    hs = (h * h + k * k) / 2
    asr = ArcSin(rho)
    For i = 1 To LG
      For ISs = -1 To 1 Step 2
        sn = Sin(asr * (ISs * XX(i, NG) + 1) / 2)
        BVN = BVN + W(i, NG) * Exp((sn * hk - hs) / (1 - sn * sn))
      Next ISs
    Next i
    BVN = BVN * asr / (4 * Pi)
  End If
  BVN = BVN + CND(-h) * CND(-k)
Else
  If rho < 0 Then
    k = -k
    hk = -hk
  End If
  If Abs(rho) < 1 Then
    Ass = (1 - rho) * (1 + rho)
    A = Sqr(Ass)
    bs = (h - k) ^ 2
    c = (4 - hk) / 8
    d = (12 - hk) / 16
    asr = -(bs / Ass + hk) / 2
    If asr > -100 Then BVN = A * Exp(asr) * (1 - c * (bs - Ass) * (1 - d * bs / 5) / 3 + c * d * Ass * Ass / 5)
    If -hk < 100 Then
      b = Sqr(bs)
      BVN = BVN - Exp(-hk / 2) * Sqr(2 * Pi) * CND(-b / A) * b * (1 - c * bs * (1 - d * bs / 5) / 3)
    End If
    A = A / 2
    For i = 1 To LG
      For ISs = -1 To 1 Step 2
        xs = (A * (ISs * XX(i, NG) + 1)) ^ 2
        rs = Sqr(1 - xs)
        asr = -(bs / xs + hk) / 2
        If asr > -100 Then
           BVN = BVN + A * W(i, NG) * Exp(asr) * (Exp(-hk * (1 - rs) / (2 * (1 + rs))) / rs - (1 + c * xs * (1 + d * xs)))
        End If
      Next ISs
    Next i
    BVN = -BVN / (2 * Pi)
  End If
  If rho > 0 Then
    BVN = BVN + CND(-Max(h, k))
  Else
    BVN = -BVN
    If k > h Then BVN = BVN + CND(k) - CND(h)
  End If
End If
CBND = BVN

End Function

Private Function ArcSin(x As Double) As Double
  If Abs(x) = 1 Then
    ArcSin = Sgn(x) * Pi / 2
  Else
    ArcSin = Atn(x / Sqr(1 - x ^ 2))
  End If
End Function


'// The cumulative bivariate normal distribution function
'// Based on Drezner-1978
Public Function CBND3(A As Double, b As Double, rho As Double) As Double

    Dim x As Variant, y As Variant
    Dim rho1 As Double, rho2 As Double, delta As Double
    Dim a1 As Double, b1 As Double, sum As Double
    Dim i As Integer, j As Integer
    
    x = Array(0.24840615, 0.39233107, 0.21141819, 0.03324666, 0.00082485334)
    y = Array(0.10024215, 0.48281397, 1.0609498, 1.7797294, 2.6697604)
    a1 = A / Sqr(2 * (1 - rho ^ 2))
    b1 = b / Sqr(2 * (1 - rho ^ 2))
    
    If A <= 0 And b <= 0 And rho <= 0 Then
        sum = 0
        For i = 0 To 4
            For j = 0 To 4
                sum = sum + x(i) * x(j) * Exp(a1 * (2 * y(i) - a1) _
                + b1 * (2 * y(j) - b1) + 2 * rho * (y(i) - a1) * (y(j) - b1))
            Next
        Next
        CBND3 = Sqr(1 - rho ^ 2) / Pi * sum
    ElseIf A <= 0 And b >= 0 And rho >= 0 Then
        CBND3 = CND(A) - CBND3(A, -b, -rho)
    ElseIf A >= 0 And b <= 0 And rho >= 0 Then
        CBND3 = CND(b) - CBND3(-A, b, -rho)
    ElseIf A >= 0 And b >= 0 And rho <= 0 Then
        CBND3 = CND(A) + CND(b) - 1 + CBND3(-A, -b, rho)
    ElseIf A * b * rho > 0 Then
        rho1 = (rho * A - b) * Sgn(A) / Sqr(A ^ 2 - 2 * rho * A * b + b ^ 2)
        rho2 = (rho * b - A) * Sgn(b) / Sqr(A ^ 2 - 2 * rho * A * b + b ^ 2)
        delta = (1 - Sgn(A) * Sgn(b)) / 4
        CBND3 = CBND3(A, 0, rho1) + CBND3(b, 0, rho2) - delta
    End If
End Function








