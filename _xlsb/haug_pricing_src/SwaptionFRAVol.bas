Attribute VB_Name = "Module1"
Option Base 1

'programmer Espen Gaarder Haug, Copyright 2006


Public Function SwaptionVol(SwapStart As Integer, SwapTenor As Integer, Zeros As Variant, Vols As Variant, CorrelationMatrix As Variant) As Double
        
        Dim Weight() As Double
        no_weights = Application.Count(Vols)
        ReDim Weight(no_weights) As Double
    
        
        For i = SwapStart + 1 To SwapStart + SwapTenor
            Weight(i) = 1 / (1 + Zeros(i)) ^ i
            SumDiscountFactors = SumDiscountFactors + Weight(i)
        Next
        
        For i = SwapStart + 1 To SwapStart + SwapTenor
            Weight(i) = Weight(i) / SumDiscountFactors
        Next
        
        
        For i = SwapStart + 1 To SwapStart + SwapTenor
                Sum = Sum + Weight(i) ^ 2 * Vols(i) ^ 2
            For j = i + 1 To SwapStart + SwapTenor
                Sum = Sum + 2 * Weight(i) * Weight(j) * Vols(i) * Vols(j) * CorrelationMatrix(i, j)
            Next
        Next
    
        SwaptionVol = Sqr(Sum)
        
End Function
