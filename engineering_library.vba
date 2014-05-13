'This module contains all of my custom functions.
'circ_segment
'moody
'density_water
'cp_water
'cp_glycol
'last updated 5-13-14 by Jacqui Nelson, Jacqnelson@gmail.com

'circ_segment
'used to calculate a partially filled circle
'Last updated 5-5-14

Function circ_segment(output, radius, partial_area)
i = 0
a = 0
b = 2 * radius * 3.141592654
s = radius * 3.141592654

Do While i <= 100
    
    Phi = s / radius
    
    area = (radius * radius * (Phi - Sin(Phi))) / 2

    If area >= partial_area Then
        b = s
        s = a + Abs(b - a) / 2
        Else
        a = s
        s = b - Abs(b - a) / 2
    End If
  
    i = i + 1
Loop

If output = "h" Then
    circ_segment = radius - radius * Cos(s / (2 * radius))
    Else
    If output = "s" Then
    circ_segment = s
    End If
End If
   
End Function

'Moody
'returns the friction factor by taking in the reynolds number (R) and relative roughness (K)
'It works for Laminar, transition, and turbulent flow
'last updated 5-5-14

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

'Density_water
'Density of water in imperial units or metric
'met takes in celcius and gives kg/m3
'imp takes in farenheit and gives lb/ft3

Function Density_water(temp, unit)
If unit = "met" Then
    Density_water = -0.0034 * temp * temp + 0.0288 * temp + 993.87
    ElseIf unit = "imp" Then
    Density_water = -0.00007 * temp * temp + 0.0052 * temp + 61.938
End If
End Function

'cp_water
'the specific heat of water
'give it the temp in F and it will give you the specific heat of water in BTu/lbmF
'I got the coefficients from charting all the numbers in excel and making a formula
'range 50F to 500F

Function cp_water(temp)
cp_water = 0.000000002 * temp * temp * temp - 0.0000005 * temp * temp + 0.00006 * temp + 0.9961
End Function

'cp_glycol
'calculates the specific heat of glycol for a given temperature and % using the data from the DOW chemcial book
'specific heat = A + BT +CT^2
'for percents between (like 55) it goes down to the nearest percent (like 50)

Function cp_glycol(percent, tempF)

Dim a() As Variant

TempC = (tempF - 32) * 5 / 9

List = Int(percent / 10)

a = Array(1.0054, 0.96705, 0.9249, 0.88012, 0.83229, 0.78229, 0.722, 0.66688, 0.60393, 0.53888, 0.4861)
b = Array(-0.00027286, -0.000027144, 0.00020429, 0.00043, 0.00062286, 0.00079286, 0.00094, 0.0010871, 0.0012043, 0.00128, 0.0013929)
c = Array(0.0000029143, 0.0000024952, 0.0000024524, 0.0000016952, 0.0000013714, 0.0000010857, 0.0000008, 0.0000004762, 0.00000028571, 0.00000019048, -0.00000005714)

cp_glycol = a(List) + b(List) * TempC + c(List) * TempC * TempC

End Function
