Attribute VB_Name = "TwoDInterpolation"

Option Explicit     'Requirs that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.

Option Base 0      'The "Option Base" statment alowws to specify 0 or 1 as the


' Programmer Espen Gaarder Haug
' Copyright 2006, Espen Gaarder Haug

Public Function TwoDimensionalInterpolation(Strike As Double, Maturity As Double, _
   Strikes As Object, Expieries As Object, VolatilityMatrix As Object) As Double
   
    Dim Bond1 As Double, Bond2 As Double
    Dim Option1 As Double, Option2 As Double
    Dim TB1 As Double, TB2 As Double
    Dim TO1 As Double, TO2 As Double
    Dim Vb1o1 As Double, Vb1o2 As Double
    Dim Vb2o1 As Double, Vb2o2 As Double

    
    Bond1 = Application.Index(Strikes, 1, Application.Match(Strike, Strikes, 1))
    Bond2 = Application.Index(Strikes, 1, Application.Match(Strike, Strikes, 1) + 1)
    Option1 = Application.Index(Expieries, Application.Match(Maturity, Expieries, 1), 1)
    Option2 = Application.Index(Expieries, Application.Match(Maturity, Expieries, 1) + 1, 1)
    TB1 = Application.Match(Bond1, Strikes, 1)
    TB2 = Application.Match(Bond2, Strikes, 1)
    TO1 = Application.Match(Maturity, Expieries, 1)
    TO2 = Application.Match(Maturity, Expieries, 1) + 1
    Vb1o1 = Application.Index(VolatilityMatrix, TO1, TB1)
    Vb1o2 = Application.Index(VolatilityMatrix, TO2, TB1)
    Vb2o1 = Application.Index(VolatilityMatrix, TO1, TB2)
    Vb2o2 = Application.Index(VolatilityMatrix, TO2, TB2)
    
    TwoDimensionalInterpolation = Vb1o1 + (Strike - Bond1) * (Vb2o1 - Vb1o1) / _
    (Bond2 - Bond1) + (Maturity - Option1) * ((Bond2 - Strike) / (Bond2 - Bond1) * (Vb1o2 - Vb1o1) / (Option2 - Option1) + (Strike - Bond1) / (Bond2 - Bond1) * (Vb2o2 - Vb2o1) / (Option2 - Option1))

    
End Function
