'returns the friction factor by taking in the reynolds number and relative roughness
'It works for Laminar, transition, and turbulent flow


Function Moody(R As Double, K As Double) As Double

Dim X1 As Double, X2 As Double, F As Double, E As Double
X1 = K * R * 0.123968186335418
X2 = Log(R) - 0.779397488455682
F = X2 - 0.2
E = (Log(X1 + F) + F - X2) / (1 + X1 + F)
F = F - (1 + X1 + F + 0.5 * E) * E * (X1 + F) / (1 + X1 + F + E * (1 + E / 3))
E = (Log(X1 + F) + F - X2) / (1 + X1 + F)
F = F - (1 + X1 + F + 0.5 * E) * E * (X1 + F) / (1 + X1 + F + E * (1 + E / 3))
F = 1.15129254649702 / F

If R >= 4000 Then
    Moody = F * F
    Else
    If R <= 2300 And R > 0 Then
        Moody = 64 / R
        Else
        If R > 2300 And R < 4000 Then
        Moody = (F * F + 64 / R) / 2
        End If
    End If
End If


End Function
