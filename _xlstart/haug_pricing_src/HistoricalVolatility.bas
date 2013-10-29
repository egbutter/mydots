Attribute VB_Name = "HistoricalVolatility"
Option Explicit     'Requirs that all variables to be declared explicitly.
Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.

Option Base 1       'The "Option Base" statment alowws to specify 0 or 1 as the
                            'default first index of arrays.
                            
' Programmer Espen Gaarder Haug, Copyright 2006

'// Correlation coefficient of logarithmic changes between two assets
Public Function HistoricalCorrelation(PricesAsset1 As Object, PricesAsset2 As Object) As Double
    HistoricalCorrelation = Application.Correl(LogChange(PricesAsset1), LogChange(PricesAsset2))
End Function
                            
 '// Pearson Kurtosis of logarithmic changes of a vector
Public Function HistoricalKurtosis(HistoricalData As Object) As Double
    HistoricalKurtosis = Application.Kurt(LogChange(HistoricalData)) + 3
End Function

'// Skewness of logarithmic changes of a vector
Public Function HistoricalSkewness(HistoricalData As Object) As Double
    HistoricalSkewness = Application.Skew(LogChange(HistoricalData))
End Function
                            
Public Function CloseVolatility(ClosePrices As Object, Optional DataPerYear As Variant) As Double
    '
    ' Volatility (= standard deviation of logarithmic changes of a vector)
    ' Default adjustment is based on data for 252 days a year (Calendar day volatility)
    '
    If IsMissing(DataPerYear) Then
        DataPerYear = 252
    End If
    CloseVolatility = Application.StDev(LogChange(ClosePrices)) * Sqr(DataPerYear)
    
End Function

Public Function HighLowVolatility(HighPrices As Object, LowPrices As Object, Optional DataPerYear As Variant) As Double
    
    Dim n As Integer
    
    n = HighPrices.Rows.Count
     If IsMissing(DataPerYear) Then
        DataPerYear = 252
    End If
    HighLowVolatility = 1 / (2 * n * Sqr(Log(2))) * Application.sum(LogHighLow(HighPrices, LowPrices)) * Sqr(DataPerYear)

End Function


Public Function HighLowCloseVolatility(HighPrices As Object, LowPrices As Object, ClosePrices As Object, Optional DataPerYear As Variant) As Double
    
    Dim n As Integer
    
    n = HighPrices.Rows.Count
     If IsMissing(DataPerYear) Then
        DataPerYear = 252
    End If
    HighLowCloseVolatility = Sqr(1 / n * 1 / 2 * Application.SumSq(LogHighLow(HighPrices, LowPrices)) _
    - 1 / n * (2 * Log(2) - 1) * Application.SumSq(LogChange(ClosePrices))) * Sqr(DataPerYear)

End Function


Public Function ExponentiallyWeightedVol(PriceVector As Object, Lambda As Double, Optional DataPerYear As Variant) As Double
    
    Dim ExpVol As Double
    Dim nRow As Integer
    Dim nCol As Integer
    Dim nVec As Integer
    Dim Element As Integer
    
    If IsMissing(DataPerYear) Then
        DataPerYear = 252
    End If
    
    ExpVol = Log(PriceVector(2) / PriceVector(1)) ^ 2
    
    
    If PriceVector.Areas.Count <> 1 Then
        ' Multiple selections not allowed
        ExponentiallyWeightedVol = CVErr(xlErrValue)
    Else
        nRow = PriceVector.Rows.Count
        nCol = PriceVector.Columns.Count
        If (nRow = 1 And nCol >= 2) Or (nCol = 1 And nRow >= 2) Then
            nVec = Application.Max(nRow, nCol)
            For Element = 2 To nVec - 1
               ExpVol = Lambda * ExpVol + (1 - Lambda) * Log(PriceVector(Element + 1) / PriceVector(Element)) ^ 2
            Next Element
            ExponentiallyWeightedVol = Sqr(ExpVol * DataPerYear)
        Else
            ' DataVector is not a vector
            ExponentiallyWeightedVol = CVErr(xlErrValue)
        End If
    End If
    
End Function

Private Function LogChange(DataVector As Object)
    '
    ' Returns the natural logarithm of the changes in DataVectortor
    '
    Dim nRow As Integer
    Dim nCol As Integer
    Dim nVec As Integer
    Dim Element As Integer
    Dim TmpVec() As Double
    If DataVector.Areas.Count <> 1 Then
        ' Multiple selections not allowed
        LogChange = CVErr(xlErrValue)
    Else
        nRow = DataVector.Rows.Count
        nCol = DataVector.Columns.Count
        If (nRow = 1 And nCol >= 2) Or (nCol = 1 And nRow >= 2) Then
            nVec = Application.Max(nRow, nCol)
            ReDim TmpVec(nVec - 1)
            For Element = 1 To nVec - 1
                TmpVec(Element) = Log(DataVector(Element + 1) / DataVector(Element))
            Next Element
            LogChange = Application.Transpose(TmpVec)
        Else
            ' DataVector is not a vector
            LogChange = CVErr(xlErrValue)
        End If
    End If
End Function

Private Function LogHighLow(HighPrices As Object, LowPrices As Object)
  
  
    Dim nRow As Integer
    Dim nCol As Integer
    Dim nVec As Integer
    Dim Element As Integer
    Dim TmpVec() As Double
    If HighPrices.Areas.Count <> 1 Or LowPrices.Areas.Count <> 1 Then
        ' Multiple selections not allowed
        LogHighLow = CVErr(xlErrValue)
    Else
        nRow = HighPrices.Rows.Count
        nCol = HighPrices.Columns.Count
        If (nRow = 1 And nCol >= 2) Or (nCol = 1 And nRow >= 2) Then
            nVec = Application.Max(nRow, nCol)
            ReDim TmpVec(nVec)
            For Element = 1 To nVec
                TmpVec(Element) = Log(HighPrices(Element) / LowPrices(Element))
            Next Element
            LogHighLow = Application.Transpose(TmpVec)
        Else
        
            LogHighLow = CVErr(xlErrValue)
        End If
    End If
End Function


Public Function VolatilityCone(DataVec As Object, VolPeriod As Integer, Optional DataPerYear As Variant)
    '
    ' Volatility cone calculation with a volatility period,
    ' Default adjustment is based on data for 252 days a year (trading days).
    ' Returns a vector which contains minimum, maximum, average and last volatility
    '
    Dim nRow As Integer, nCol As Integer, nVec As Integer, Elem As Integer
    Dim i As Integer, j As Integer
    Dim VolVec(4) As Double
    Dim TmpVec() As Double
    
    If IsMissing(DataPerYear) Then
        DataPerYear = 252
    End If
    If DataVec.Areas.Count <> 1 Then
        ' Multiple selections not allowed
        VolatilityCone = CVErr(xlErrValue)
    Else
        nRow = DataVec.Rows.Count
        nCol = DataVec.Columns.Count
        If (nRow = 1 And nCol >= 2) Or (nCol = 1 And nRow >= 2) Then
            nVec = Application.Max(nRow, nCol)
            If VolPeriod <= nVec - 1 Then
                ReDim TmpVec(nVec - VolPeriod)
                For j = 1 To nVec - VolPeriod
                    TmpVec(j) = CloseVolatility(DataVec.Range(Cells(j, 1), Cells(j + VolPeriod, 1)), DataPerYear)
                Next j
                VolVec(1) = Application.Min(TmpVec)
                VolVec(2) = Application.Max(TmpVec)
                VolVec(3) = Application.Average(TmpVec)
                VolVec(4) = TmpVec(nVec - VolPeriod)
            Else
                ' Not enough data for a volatility period this long
                VolVec(1) = CVErr(xlErrValue)
            End If
            VolatilityCone = Application.Transpose(VolVec)
        Else
            VolatilityCone = CVErr(xlErrValue)
        End If
    End If
End Function


Public Function ConfidenceIntervalVolatility(alfa As Double, n As Integer, VolatilityEstimate As Double, _
UpperLower As String)

    'UpperLower     ="L" gives the lower cofidence interval
    '               ="U" gives the upper cofidence interval
    'n: number of observations
    If UpperLower = "L" Then
        ConfidenceIntervalVolatility = VolatilityEstimate * Sqr((n - 1) / (Application.ChiInv(alfa / 2, n - 1)))
    ElseIf UpperLower = "U" Then
        ConfidenceIntervalVolatility = VolatilityEstimate * Sqr((n - 1) / (Application.ChiInv(1 - alfa / 2, n - 1)))
    End If
    
End Function
