Attribute VB_Name = "VolSwapVBA"
Option Explicit


' Programmer Espen Gaarder Haug
' Copyright Espen Gaarder Haug 2006

 'F(v,T) for problem A
Public Function FvTProblemA(v, T, kappa, Theta)
    FvTProblemA = Theta + Exp(-kappa * T) * (v - Theta)
End Function



' F(v,I,t) for problem B
Public Function FvIt(v, i, T, t1, kappa, Theta)
        
            FvIt = Theta * (T - t1 + (Exp(-kappa * (T - t1)) - 1) / kappa) _
            + 1 / kappa * (1 - Exp(-kappa * (T - t1))) * v + i
        
End Function

 'G(v,T) for problem A
Public Function GvtProblemA(v, T, kappa, Theta, gamma)

   GvtProblemA = 2 * kappa ^ 2 * Theta ^ 2 / (gamma ^ 2 - kappa) * ((Exp((gamma ^ 2 - 2 * kappa) * T) - 1) / (gamma ^ 2 - 2 * kappa) + (Exp(-kappa * T) - 1) / kappa) _
    + 2 * kappa * Theta / (gamma ^ 2 - kappa) * (Exp((gamma ^ 2 - 2 * kappa) * T) - Exp(-kappa * T)) * v + Exp((gamma ^ 2 - 2 * kappa) * T) * v ^ 2

End Function


' G(v,I,t) for problem B
Public Function GvIt(v, i, T, t1, kappa, Theta, gamma)
        Dim ft As Double, gt As Double, ht As Double, lt As Double, nt As Double
        
         ft = Theta ^ 2 * (T - t1) ^ 2 - 4 * Theta ^ 2 * (gamma ^ 2 - kappa) / (kappa * (gamma ^ 2 - 2 * kappa)) _
            * (T - t1 + (Exp(-kappa * (T - t1)) - 1) / kappa) _
            - 4 * Theta ^ 2 * kappa ^ 2 / ((gamma ^ 2 - kappa) ^ 2 * (gamma ^ 2 - 2 * kappa)) _
            * ((1 - Exp((gamma ^ 2 - 2 * kappa) * (T - t1))) / (gamma ^ 2 - 2 * kappa) + (1 - Exp(-kappa * (T - t1))) / kappa) _
            - 2 * Theta ^ 2 * (gamma ^ 2 + kappa) / (gamma ^ 2 - kappa) _
            * (Exp(-kappa * (T - t1)) * (T - t1) / kappa + 1 / kappa ^ 2 * (Exp(-kappa * (T - t1)) - 1))
            
             gt = 2 * Theta / kappa * (T - t1) - 4 * Theta * (gamma ^ 2 - kappa) / (kappa ^ 2 * (gamma ^ 2 - 2 * kappa)) * (1 - Exp(-kappa * (T - t1))) _
           + 4 * Theta * kappa / ((gamma ^ 2 - kappa) ^ 2 * (gamma ^ 2 - 2 * kappa)) _
           * (Exp((gamma ^ 2 - 2 * kappa) * (T - t1)) - Exp(-kappa * (T - t1))) _
           + 2 * Theta * (gamma ^ 2 + kappa) / (kappa * (gamma ^ 2 - kappa)) * (T - t1) * Exp(-kappa * (T - t1))
           
           ht = 2 / (kappa * (gamma ^ 2 - 2 * kappa)) * (Exp((gamma ^ 2 - 2 * kappa) * (T - t1)) - 1) _
         - 2 / (kappa * (gamma ^ 2 - kappa)) * (Exp((gamma ^ 2 - 2 * kappa) * (T - t1)) - Exp(-kappa * (T - t1)))
         
         lt = 2 * Theta * (T - t1 + (Exp(-kappa * (T - t1)) - 1) / kappa)
            
        nt = 2 / kappa * (1 - Exp(-kappa * (T - t1)))
        
        GvIt = (ft + gt * v + ht * v ^ 2 + lt * i + nt * i + i ^ 2) / T ^ 2
        
End Function

Public Function ft(v, i, T, t1, kappa, Theta, gamma)
        
            ft = Theta ^ 2 * (T - t1) ^ 2 - 4 * Theta ^ 2 * (gamma ^ 2 - kappa) / (kappa * (gamma ^ 2 - 2 * kappa)) _
            * (T - t1 + (Exp(-kappa * (T - t1)) - 1) / kappa) _
            - 4 * Theta ^ 2 * kappa ^ 2 / ((gamma ^ 2 - kappa) ^ 2 * (gamma ^ 2 - 2 * kappa)) _
            * ((1 - Exp((gamma ^ 2 - 2 * kappa) * (T - t1))) / (gamma ^ 2 - 2 * kappa) + (1 - Exp(-kappa * (T - t1))) / kappa) _
            - 2 * Theta ^ 2 * (gamma ^ 2 + kappa) / (gamma ^ 2 - kappa) _
            * (Exp(-kappa * (T - t1)) * (T - t1) / kappa + 1 / kappa ^ 2 * (Exp(-kappa * (T - t1)) - 1))
            
        
End Function


Public Function gt(v, i, T, t1, kappa, Theta, gamma)
        
           gt = 2 * Theta / kappa * (T - t1) - 4 * Theta * (gamma ^ 2 - kappa) / (kappa ^ 2 * (gamma ^ 2 - 2 * kappa)) * (1 - Exp(-kappa * (T - t1))) _
           + 4 * Theta * kappa / ((gamma ^ 2 - kappa) ^ 2 * (gamma ^ 2 - 2 * kappa)) _
           * (Exp((gamma ^ 2 - 2 * kappa) * (T - t1)) - Exp(-kappa * (T - t1))) _
           + 2 * Theta * (gamma ^ 2 + kappa) / (kappa * (gamma ^ 2 - kappa)) * (T - t1) * Exp(-kappa * (T - t1))
            
        
End Function

Public Function ht(v, i, T, t1, kappa, Theta, gamma)
        
         ht = 2 / (kappa * (gamma ^ 2 - 2 * kappa)) * (Exp((gamma ^ 2 - 2 * kappa) * (T - t1)) - 1) _
         - 2 / (kappa * (gamma ^ 2 - kappa)) * (Exp((gamma ^ 2 - 2 * kappa) * (T - t1)) - Exp(-kappa * (T - t1)))
            
        
End Function
