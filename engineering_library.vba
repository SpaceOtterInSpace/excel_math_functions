'This module contains all of my custom functions.
'circ_segment
'moody
'density_water
'cp_water
'cp_glycol
'unit_conversion
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

'unit_conversion
'Last updated 5-5-14


Function unit_convert(number, unit_type, from_unit, to_unit) 'This function works by using the top number to convert to the first variable and the bottom number to convert to the unit you want
   
 x = 0 'this is for when the conversion involves addition
    'pressure
    If unit_type = "pressure" Then
        u1 = "psi"
        v1 = 1
        u2 = "bar"
        v2 = 14.5033
        u3 = "Pa"
        v3 = 0.000145037738
        u4 = "kPa"
        v4 = 0.145037738
        u5 = "inH2O"
        v5 = 0.0361396333
        u6 = "ftH2O"
        v6 = 0.4335275
    Else
            'specific heat
        If unit_type = "specific heat (cp)" Then
            u1 = "Btu/lbmF"
            v1 = 1
            u2 = "kJ/kgK"
            v2 = 1 / 4.1868
            u3 = "kcal/kgC"
            v3 = 1
            
            Else
                'thermal conductivity
                If unit_type = "thermal conductivity" Then
                    u1 = "Btu-in/hr-ft2-F"
                    v1 = 1
                    u2 = "Btu-ft/hr-ft2-F"
                    v2 = 12
                    u3 = "Btu-in/s-ft2-F"
                    v3 = 3600
                    u4 = "Btu/hr-ft-F"
                    v4 = 12
                    u5 = "W/mK"
                    v5 = 6.933471799
                    
                    Else
                    'density
                    If unit_type = "density" Then
                        u1 = "lb/ft3"
                        v1 = 1
                        u2 = "SG_H2O"
                        v2 = 62.4
                        u3 = "kg/m3"
                        v3 = 0.0624279606
                        u4 = "g/cm3"
                        v4 = 62.4279606
                        If from_unit = "API*" Then 'This is just an estimate, the story of API is very complicated. This becomes more correct the closer you are to 60F
                            number = 141.5 / (number + 131.5)
                            from_unit = "SG_H2O"
                            u2 = "SG_H2O"
                            v2 = 62.4
                        End If
   
                    End If
                        
                        
                End If
        End If
    End If

    
    
    If from_unit = u1 Then
        Top = v1
    Else
        If from_unit = u2 Then
            Top = v2
        Else
            If from_unit = u3 Then
                Top = v3
            Else
                If from_unit = u4 Then
                Top = v4
                Else
                    If from_unit = u5 Then
                    Top = v5
                    Else
                        If from_unit = u6 Then
                        Top = v6
                        Else
                            If from_unit = u7 Then
                            Top = v7
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If to_unit = u1 Then
        bottom = v1
        Else
        If to_unit = u2 Then
            bottom = v2
        Else
            If to_unit = u3 Then
            bottom = v3
            Else
                If to_unit = u4 Then
                bottom = v4
                Else
                    If to_unit = u5 Then
                    bottom = v5
                    Else
                        If to_unit = u6 Then
                        bottom = v6
                        Else
                            If to_unit = u7 Then
                            bottom = v7
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
   If to_unit <> "API*" Then
    unit_convert = number * Top / bottom
   End If
 
 If to_unit = "API*" Then
    unit_convert = 141.5 / (number * Top / 62.4) - 131.5
 End If

      
      
End Function
