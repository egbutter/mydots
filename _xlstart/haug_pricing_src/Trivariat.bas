Attribute VB_Name = "Trivariat"
Option Explicit

Global Const Pi = 3.14159265358979


'Appendix: Fortran Source Code


Function CTND(LIMIT1 As Double, LIMIT2 As Double, LIMIT3 As Double, _
            SIGMA1 As Double, SIGMA2 As Double, SIGMA3 As Double) As Double
'
'     A function for computing trivariate normal probabilities.
'     This function uses an algorithm given in the paper
'        "Numerical Computation of Bivariate and
'             Trivariate Normal Probabilities",
'     by
'       Alan Genz
'       Department of Mathematics
'       Washington State University
'       Pullman, WA 99164-3113
'       Email : alangenz@wsu.edu
'
' CTND calculates the probability that X(I) < LIMIT(I), I = 1, 2, 3.
'
' Parameters
'
'   LIMIT  DOUBLE PRECISION array of three upper integration limits.
'   SIGMA  DOUBLE PRECISION array of three correlation coefficients,
'          SIGMA should contain the lower left portion of the
'          correlation matrix R.
'          SIGMA(1) = R(2,1), SIGMA(2) = R(3,1), SIGMA(3) = R(3,2).
'
'    CTND cuts the outer integral over -infinity to B1 to
'      an integral from -8.5 to B1 and then uses an adaptive
'      integration method to compute the integral of a bivariate
'      normal distribution function.
'
Dim TAIL As Boolean
'
'     Bivariate normal distribution function CBND is required.
'
Dim SQ21 As Double, SQ31 As Double, rho As Double
Dim b1 As Double, B2 As Double, B3 As Double, b2p As Double, b3p As Double
Dim RHO21 As Double, RHO31 As Double, RHO32 As Double
Const SQTWPI = 2.506628274631
Const XCUT = -8.5
Const EPS = 5E-16

      'COMMON /TRVBKD/B2P, B3P, RHO21, RHO31, RHO
      
      b1 = LIMIT1
      B2 = LIMIT2
      B3 = LIMIT3
      RHO21 = SIGMA1
      RHO31 = SIGMA2
      RHO32 = SIGMA3
      If Abs(B2) > Max(Abs(b1), Abs(B3)) Then
         b1 = B2
         B2 = LIMIT1
         RHO31 = RHO32
         RHO32 = SIGMA2
      ElseIf Abs(B3) > Max(Abs(b1), Abs(B2)) Then
         b1 = B3
         B3 = LIMIT1
         RHO21 = RHO32
         RHO32 = SIGMA1
      End If
    
    TAIL = False
      
      If b1 > 0 Then
         TAIL = True
         b1 = -b1
         RHO21 = -RHO21
         RHO31 = -RHO31
      End If
      
      If b1 > XCUT Then
         If 2 * Abs(RHO21) < 1 Then
            SQ21 = Sqr(1 - RHO21 ^ 2)
         Else
            SQ21 = Sqr((1 - RHO21) * (1 + RHO21))
         End If
         
         If 2 * Abs(RHO31) < 1 Then
            SQ31 = Sqr(1 - RHO31 ^ 2)
         Else
            SQ31 = Sqr((1 - RHO31) * (1 + RHO31))
         End If
         
         rho = (RHO32 - RHO21 * RHO31) / (SQ21 * SQ31)
         b2p = B2 / SQ21
         RHO21 = RHO21 / SQ21
         b3p = B3 / SQ31
         RHO31 = RHO31 / SQ31
         CTND = ADONED(XCUT, b1, EPS, b2p, b3p, RHO21, RHO31, rho) / SQTWPI
      Else
         CTND = 0
      End If
      
      If TAIL = True Then
        CTND = CBND(B2, B3, RHO32) - CTND
      End If

End Function

Public Function Max(x, y)
            Max = Application.Max(x, y)
End Function
    
Function TRVFND(T As Double, _
            B2 As Double, B3 As Double, RHO21 As Double, RHO31 As Double, rho As Double)

      TRVFND = Exp(-T * T / 2) * CBND(-T * RHO21 + B2, -T * RHO31 + B3, rho)
End Function

Function ADONED(A As Double, b As Double, TOL As Double, _
            b2p As Double, b3p As Double, RHO21 As Double, RHO31 As Double, rho As Double) As Double
'
'     One Dimensional Globally Adaptive Integration Function
'

Dim i As Integer, IM As Integer, IP As Integer
Const NL = 100
Dim EI(NL) As Double, AI(NL) As Double, BI(NL) As Double, FI(NL) As Double
Dim FIN As Double, ERR As Double
      
      IP = 1
      AI(IP) = A
      BI(IP) = b
      FI(IP) = KRNRDD(AI(IP), BI(IP), EI(IP), b2p, b3p, RHO21, RHO31, rho)
      IM = 1
10    IM = IM + 1
      BI(IM) = BI(IP)
      AI(IM) = (AI(IP) + BI(IP)) / 2
      BI(IP) = AI(IM)
      FIN = FI(IP)
      FI(IP) = KRNRDD(AI(IP), BI(IP), EI(IP), b2p, b3p, RHO21, RHO31, rho)
      FI(IM) = KRNRDD(AI(IM), BI(IM), EI(IM), b2p, b3p, RHO21, RHO31, rho)
      ERR = Abs(FIN - FI(IP) - FI(IM)) / 2
      EI(IP) = EI(IP) + ERR
      EI(IM) = EI(IM) + ERR
      IP = 1
      ERR = 0
      FIN = 0
      For i = 1 To IM
         If EI(i) > EI(IP) Then
            IP = i
         End If
         
         FIN = FIN + FI(i)
         ERR = ERR + EI(i)
      Next i
      If ERR > TOL And IM < NL Then
        GoTo 10
      End If
      
      ADONED = FIN

End Function

Function KRNRDD(A As Double, b As Double, ABSERR As Double, _
            b2p As Double, b3p As Double, RHO21 As Double, RHO31 As Double, rho As Double) As Double
'
'     Kronrod Rule
'

Dim ABSCIS As Double, CENTER As Double, FC As Double, FUNSUM As Double, HFLGTH As Double
Dim RESLTG As Double, RESLTK As Double
'
'           The abscissae and weights are given for the interval (-1,1)
'           because of symmetry only the positive abscisae and their
'           corresponding weights are given.
'
'           XGK    - abscissae of the 2N+1-point Kronrod rule:
'                    XGK(2), XGK(4), ...  N-point Gauss rule abscissae;
'                    XGK(1), XGK(3), ...  abscissae optimally added
'                    to the N-point Gauss rule.
'
'           WGK    - weights of the 2N+1-point Kronrod rule.
'
'           WG     - weights of the N-point Gauss rule.
'
Dim j As Integer
Const N = 11
'
Dim WG(0 To (N + 1) / 2) As Double, WGK(0 To N) As Double, XGK(0 To N) As Double

       WG(1) = 5.56685671161745E-02
       WG(2) = 0.125580369464905
       WG(3) = 0.186290210927735
       WG(4) = 0.233193764591991
       WG(5) = 0.262804544510248
       WG(0) = 0.272925086777901
'
       XGK(1) = 0.996369613889543
       XGK(2) = 0.978228658146057
       XGK(3) = 0.941677108578068
       XGK(4) = 0.887062599768095
       XGK(5) = 0.816057456656221
       XGK(6) = 0.730152005574049
       XGK(7) = 0.630599520161965
       XGK(8) = 0.519096129206812
       XGK(9) = 0.397944140952378
       XGK(10) = 0.269543155952345
       XGK(11) = 0.136113000799362
       XGK(0) = 0#
'
       WGK(1) = 9.76544104596129E-03
       WGK(2) = 2.71565546821044E-02
       WGK(3) = 4.58293785644267E-02
       WGK(4) = 6.30974247503748E-02
       WGK(5) = 7.86645719322276E-02
       WGK(6) = 9.29530985969007E-02
       WGK(7) = 0.105872074481389
       WGK(8) = 0.116739502461047
       WGK(9) = 0.125158799100319
       WGK(10) = 0.131280684229806
       WGK(11) = 0.135193572799885
       WGK(0) = 0.136577794711118
'
'
'           List of major variables
'
'           CENTER  - mid point of the interval
'           HFLGTH  - half-length of the interval
'           ABSCIS   - abscissae
'           RESLTG   - result of the N-point Gauss formula
'           RESLTK   - result of the 2N+1-point Kronrod formula
'
'
      HFLGTH = (b - A) / 2
      CENTER = (b + A) / 2
'
'           compute the 2N+1-point Kronrod approximation to
'           the integral, and estimate the absolute error.
'
      FC = TRVFND(CENTER, b2p, b3p, RHO21, RHO31, rho)
      RESLTG = FC * WG(0)
      RESLTK = FC * WGK(0)
      For j = 1 To N
         ABSCIS = HFLGTH * XGK(j)
         FUNSUM = TRVFND(CENTER - ABSCIS, b2p, b3p, RHO21, RHO31, rho) + _
                    TRVFND(CENTER + ABSCIS, b2p, b3p, RHO21, RHO31, rho)
         RESLTK = RESLTK + WGK(j) * FUNSUM
         If j Mod 2 = 0 Then
            RESLTG = RESLTG + WG(j / 2) * FUNSUM
         End If
      Next j
      KRNRDD = RESLTK * HFLGTH
      ABSERR = 3 * Abs((RESLTK - RESLTG) * HFLGTH)
      
End Function



