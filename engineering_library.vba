'This module contains some engineering custom functions.
'circ_segment
'moody
'density_water
'cp_water
'cp_glycol
'nps_data
'convert_pressure
'convert_specific_heat
'convert_thermal_conductivity
'convert_density
'vessel_pressure_withOD
'vessel_thickness_withOD
'vessel_pressure_withID
'vessel_thickness_withID
'last updated 5-27-14 by Jacqui Nelson, Jacqnelson@gmail.com

'circ_segment
'used to calculate a partially filled circle
'Last updated 5-5-14

Function circ_segment(radius, partial_area, pickOutput_height_arcLength)
i = 0
a = 0
b = 2 * radius * 3.141592654
arcLength = radius * 3.141592654

Do While i <= 100
    
    Phi = arcLength / radius
    
    area = (radius * radius * (Phi - Sin(Phi))) / 2

    If area >= partial_area Then
        b = arcLength
        arcLength = a + Abs(b - a) / 2
        Else
        a = arcLength
        arcLength = b - Abs(b - a) / 2
    End If
  
    i = i + 1
Loop

If pickOutput_height_arcLength = "height" Then
    circ_segment = radius - radius * Cos(arcLength / (2 * radius))
    Else
    If pickOutput_height_arcLength = "arcLength" Then
    circ_segment = arcLength
    End If
End If
End Function

'Moody
'returns the friction factor by taking in the reynolds number (R) and relative roughness (K)
'It works for Laminar, transition, and turbulent flow
'last updated 5-5-14

Function Moody(Reynolds As Double, rel_roughness As Double) As Double

Dim X1 As Double, X2 As Double, f As Double, E As Double
X1 = rel_roughness * Reynolds * 0.123968186335418
X2 = Log(Reynolds) - 0.779397488455682
f = X2 - 0.2
E = (Log(X1 + f) + f - X2) / (1 + X1 + f)
f = f - (1 + X1 + f + 0.5 * E) * E * (X1 + f) / (1 + X1 + f + E * (1 + E / 3))
E = (Log(X1 + f) + f - X2) / (1 + X1 + f)
f = f - (1 + X1 + f + 0.5 * E) * E * (X1 + f) / (1 + X1 + f + E * (1 + E / 3))
f = 1.15129254649702 / f

If Reynolds >= 4000 Then
    Moody = f * f
    Else
    If Reynolds <= 2300 And Reynolds > 0 Then
        Moody = 64 / Reynolds
        Else
        If Reynolds > 2300 And Reynolds < 4000 Then
        Moody = (f * f + 64 / Reynolds) / 2
        End If
    End If
End If
End Function

'Density_water
'Density of water in imperial units or metric
'met takes in celcius and gives kg/m3
'imp takes in farenheit and gives lb/ft3
'last updated 5-5-14

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
'last updated 5-5-14

Function cp_water(temp_F)
cp_water = 0.000000002 * temp_F * temp_F * temp_F - 0.0000005 * temp_F * temp_F + 0.00006 * temp_F + 0.9961
End Function

'cp_glycol
'calculates the specific heat of glycol for a given temperature and % using the data from the DOW chemcial book
'specific heat = A + BT +CT^2
'for percents between (like 55) it goes down to the nearest percent (like 50)
'last updated 5-5-14

Function cp_glycol(percent, TempF)

Dim a() As Variant

TempC = (TempF - 32) * 5 / 9

List = Int(percent / 10)

a = Array(1.0054, 0.96705, 0.9249, 0.88012, 0.83229, 0.78229, 0.722, 0.66688, 0.60393, 0.53888, 0.4861)
b = Array(-0.00027286, -0.000027144, 0.00020429, 0.00043, 0.00062286, 0.00079286, 0.00094, 0.0010871, 0.0012043, 0.00128, 0.0013929)
c = Array(0.0000029143, 0.0000024952, 0.0000024524, 0.0000016952, 0.0000013714, 0.0000010857, 0.0000008, 0.0000004762, 0.00000028571, 0.00000019048, -0.00000005714)

cp_glycol = a(List) + b(List) * TempC + c(List) * TempC * TempC

End Function

'nps_data
'takes the nps and the schedule and will give you either OD, ID or thickness, depending on what you pick for output type.
'data taken from edgen murry nps chart
'last updated 5/14/14 by jacqui nelson

Function nps_data(nps, schedule, output_thickness_OD_ID)

If nps = 1 / 8 Then
    Thickness = Array(0.405, "x", "x", "x", 0.068, 0.068, "x", 0.095, 0.095, "x", "x", "x", "x", "x")
    Else
    If nps = 1 / 4 Then
        Thickness = Array(0.54, "x", "x", "x", 0.088, 0.088, "x", 0.119, 0.119, "x", "x", "x", "x", "x")
        Else
        If nps = 3 / 8 Then
            Thickness = Array(0.675, "x", "x", "x", 0.091, 0.091, "x", 0.126, 0.126, "x", "x", "x", "x", "x")
            Else
            If nps = 1 / 2 Then
                Thickness = Array(0.84, "x", "x", "x", 0.109, 0.109, "x", 0.147, 0.147, "x", "x", "x", 0.188, 0.294)
                Else
                If nps = 3 / 4 Then
                    Thickness = Array(1.05, "x", "x", "x", 0.113, 0.113, "x", 0.154, 0.154, "x", "x", "x", 0.219, 0.308)
                    Else
                    If nps = 1 Then
                        Thickness = Array(1.315, "x", "x", "x", 0.133, 0.133, "x", 0.179, 0.179, "x", "x", "x", 0.25, 0.358)
                        Else
                        If nps = 1 + 1 / 4 Then
                            Thickness = Array(1.66, "x", "x", "x", 0.14, 0.14, "x", 0.191, 0.191, "x", "x", "x", 0.25, 0.382)
                            Else
                            If nps = 1 + 1 / 2 Then
                                Thickness = Array(1.9, "x", "x", "x", 0.145, 0.145, "x", 0.2, 0.2, "x", "x", "x", 0.281, 0.4)
                                Else
                                If nps = 2 Then
                                    Thickness = Array(2.375, "x", "x", "x", 0.154, 0.154, "x", 0.218, 0.218, "x", "x", "x", 0.344, 0.436)
                                    Else
                                    If nps = 2 + 1 / 2 Then
                                        Thickness = Array(2.875, "x", "x", "x", 0.203, 0.203, "x", 0.276, 0.276, "x", "x", "x", 0.375, 0.552)
                                        Else
                                        If nps = 3 Then
                                            Thickness = Array(3.5, "x", "x", "x", 0.216, 0.216, "x", 0.3, 0.3, "x", "x", "x", 0.438, 0.6)
                                            Else
                                            If nps = 3 + 1 / 2 Then
                                                Thickness = Array(4, "x", "x", "x", 0.226, 0.226, "x", 0.318, 0.318, "x", "x", "x", "x", "x")
                                                Else
                                                If nps = 4 Then
                                                    Thickness = Array(4.5, "x", "x", "x", 0.237, 0.237, "x", 0.337, 0.337, "x", 0.438, "x", 0.531, 0.674)
                                                    Else
                                                    If nps = 5 Then
                                                        Thickness = Array(5.563, "x", "x", "x", 0.258, 0.258, "x", 0.375, 0.375, "x", 0.5, "x", 0.625, 0.75)
                                                        Else
                                                        If nps = 6 Then
                                                            Thickness = Array(6.625, "x", "x", "x", 0.28, 0.28, "x", 0.432, 0.432, "x", 0.562, "x", 0.719, 0.864)
                                                            Else
                                                            If nps = 8 Then
                                                                Thickness = Array(8.625, "x", 0.25, 0.277, 0.322, 0.322, 0.406, 0.5, 0.5, 0.594, 0.719, 0.812, 0.906, 0.875)
                                                                Else
                                                                If nps = 10 Then
                                                                    Thickness = Array(10.75, "x", 0.25, 0.307, 0.365, 0.365, 0.5, 0.5, 0.594, 0.719, 0.844, 1, 1.125, 1)
                                                                    Else
                                                                    If nps = 12 Then
                                                                        Thickness = Array(12.75, "x", 0.25, 0.33, 0.375, 0.406, 0.562, 0.5, 0.688, 0.844, 1, 1.125, 1.312, 1)
                                                                        Else
                                                                        If nps = 14 Then
                                                                            Thickness = Array(14, 0.25, 0.312, 0.375, 0.375, 0.438, 0.594, 0.5, 0.75, 0.938, 1.094, 1.25, 1.406, "x")
                                                                            Else
                                                                            If nps = 16 Then
                                                                                Thickness = Array(16, 0.25, 0.312, 0.375, 0.375, 0.5, 0.656, 0.5, 0.844, 1.031, 1.219, 1.438, 1.594, "x")
                                                                                Else
                                                                                If nps = 18 Then
                                                                                    Thickness = Array(18, 0.25, 0.312, 0.438, 0.375, 0.562, 0.75, 0.5, 0.938, 1.156, 1.375, 1.562, 1.781, "x")
                                                                                    Else
                                                                                    If nps = 20 Then
                                                                                        Thickness = Array(20, 0.25, 0.375, 0.5, 0.375, 0.594, 0.812, 0.5, 1.031, 1.28, 1.5, 1.75, 1.968, "x")
                                                                                        Else
                                                                                        If nps = 22 Then
                                                                                            Thickness = Array(22, 0.25, 0.375, 0.5, 0.375, "x", 0.875, 0.5, 1.125, 1.375, 1.625, 1.875, 2.125, "x")
                                                                                            Else
                                                                                            If nps = 24 Then
                                                                                                Thickness = Array(24, 0.25, 0.375, 0.562, 0.375, 0.688, 0.989, 0.5, 1.219, 1.531, 1.812, 2.062, 2.344, "x")
                                                                                                Else
                                                                                                If nps = 30 Then
                                                                                                    Thickness = Array(30, 0.312, 0.5, 0.625, 0.375, "x", "x", 0.5, "x", "x", "x", "x", "x", "x")
                                                                                                    Else
                                                                                                    If nps = 32 Then
                                                                                                        Thickness = Array(32, 0.312, 0.5, 0.625, 0.375, 0.688, "x", "x", "x", "x", "x", "x", "x", "x")
                                                                                                        Else
                                                                                                        If nps = 34 Then
                                                                                                            Thickness = Array(34, 0.312, 0.5, 0.625, 0.375, 0.688, "x", "x", "x", "x", "x", "x", "x", "x")
                                                                                                            Else
                                                                                                            If nps = 36 Then
                                                                                                                Thickness = Array(36, 0.312, 0.5, 0.625, 0.375, 0.75, "x", "x", "x", "x", "x", "x", "x", "x")
                                                                                                                Else
                                                                                                                If nps = 42 Then
                                                                                                                    Thickness = Array(42, "x", 0.5, 0.625, 0.375, 0.75, "x", "x", "x", "x", "x", "x", "x", "x")
                                                                                                                    Else
                                                                                                                    Thickness = "troz"

End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
i = 0
If schedule = 10 Then
    i = 1
    Else
    If schedule = 20 Then
    i = 2
    Else
        If schedule = 30 Then
        i = 3
        Else
        If schedule = "std" Then
        i = 4
        Else
            If schedule = 40 Then
            i = 5
            Else
                If schedule = 60 Then
                i = 6
                Else
                    If schedule = "xs" Then
                    i = 7
                    Else
                        If schedule = 80 Then
                        i = 8
                        Else
                            If schedule = 100 Then
                            i = 9
                            Else
                                If schedule = 120 Then
                                i = 10
                                Else
                                    If schedule = 140 Then
                                    i = 11
                                    Else
                                        If schedule = 160 Then
                                        i = 12
                                        Else
                                            If schedule = "xxs" Then
                                            i = 13
                                            Else
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        End If
    End If
End If
    
If output_thickness_OD_ID = "thickness" Then
    nps_data = Thickness(i)
    Else
    If output_thickness_OD_ID = "OD" And Thickness(i) < 100 And Thickness(i) > 0 Then 'the test on thickness is just to make sure it is a number and not the letter x
        nps_data = Thickness(0)
        Else
        If output_thickness_OD_ID = "ID" And Thickness(i) < 100 And Thickness(i) > 0 Then
            nps_data = Thickness(0) - 2 * Thickness(i)
            Else
            nps_data = "doesn't exist"
        End If
    End If
End If

End Function

'convert_pressure
'This function works by using the top number to convert to the first variable and the bottom number to convert to the unit you want
'last updated 5/14

Function convert_pressure(number, pickFromUnit_psi_bar_Pa_kPa_inH2O_ftH2O, pickToUnit_psi_bar_Pa_kPa_inH2O_ftH2O)
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

    If pickFromUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u1 Then
        Top = v1
    Else
        If pickFromUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u2 Then
            Top = v2
        Else
            If pickFromUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u3 Then
                Top = v3
            Else
                If pickFromUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u4 Then
                Top = v4
                Else
                    If pickFromUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u5 Then
                    Top = v5
                    Else
                        If pickFromUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u6 Then
                        Top = v6
                        Else
                            If pickFromUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u7 Then
                            Top = v7
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If pickToUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u1 Then
        bottom = v1
        Else
        If pickToUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u2 Then
            bottom = v2
        Else
            If pickToUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u3 Then
            bottom = v3
            Else
                If pickToUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u4 Then
                bottom = v4
                Else
                    If pickToUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u5 Then
                    bottom = v5
                    Else
                        If pickToUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u6 Then
                        bottom = v6
                        Else
                            If pickToUnit_psi_bar_Pa_kPa_inH2O_ftH2O = u7 Then
                            bottom = v7
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    convert_pressure = number * Top / bottom
         
End Function

'specific heat
'This function works by using the top number to convert to the first variable and the bottom number to convert to the unit you want
'last updated 5-14-14

Function convert_specific_heat(number, pickFromUnit_BtuperlbmF_kJperkgK_kcalperkgC, pickToUnit_BtuperlbmF_kJperkgK_kcalperkgC)
   
          

            u1 = "Btu/lbmF"
            v1 = 1
            u2 = "kJ/kgK"
            v2 = 1 / 4.1868
            u3 = "kcal/kgC"
            v3 = 1

    If pickFromUnit_BtuperlbmF_kJperkgK_kcalperkgC = u1 Then
        Top = v1
    Else
        If pickFromUnit_BtuperlbmF_kJperkgK_kcalperkgC = u2 Then
            Top = v2
        Else
            If pickFromUnit_BtuperlbmF_kJperkgK_kcalperkgC = u3 Then
                Top = v3
            Else
                If pickFromUnit_BtuperlbmF_kJperkgK_kcalperkgC = u4 Then
                Top = v4
                Else
                    If pickFromUnit_BtuperlbmF_kJperkgK_kcalperkgC = u5 Then
                    Top = v5
                    Else
                        If pickFromUnit_BtuperlbmF_kJperkgK_kcalperkgC = u6 Then
                        Top = v6
                        Else
                            If pickFromUnit_BtuperlbmF_kJperkgK_kcalperkgC = u7 Then
                            Top = v7
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If pickToUnit_BtuperlbmF_kJperkgK_kcalperkgC = u1 Then
        bottom = v1
        Else
        If pickToUnit_BtuperlbmF_kJperkgK_kcalperkgC = u2 Then
            bottom = v2
        Else
            If pickToUnit_BtuperlbmF_kJperkgK_kcalperkgC = u3 Then
            bottom = v3
            Else
                If pickToUnit_BtuperlbmF_kJperkgK_kcalperkgC = u4 Then
                bottom = v4
                Else
                    If pickToUnit_BtuperlbmF_kJperkgK_kcalperkgC = u5 Then
                    bottom = v5
                    Else
                        If pickToUnit_BtuperlbmF_kJperkgK_kcalperkgC = u6 Then
                        bottom = v6
                        Else
                            If pickToUnit_BtuperlbmF_kJperkgK_kcalperkgC = u7 Then
                            bottom = v7
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
convert_specific_heat = number * Top / bottom
      
End Function


 'thermal conductivity
 'last updated 5-14-14
 
Function convert_thermal_conductivity(number, pickFromUnit_BTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK, pickToUnitBTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK) 'This function works by using the top number to convert to the first variable and the bottom number to convert to the unit you want

        
               

                    u1 = "Btuin/hrft2F"
                    v1 = 1
                    u2 = "Btuft/hrft2F"
                    v2 = 12
                    u3 = "Btuin/sft2F"
                    v3 = 3600
                    u4 = "Btu/hrftF"
                    v4 = 12
                    u5 = "W/mK"
                    v5 = 6.933471799
    
    
    If pickFromUnit_BTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u1 Then
        Top = v1
    Else
        If pickFromUnit_BTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u2 Then
            Top = v2
        Else
            If pickFromUnit_BTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u3 Then
                Top = v3
            Else
                If pickFromUnit_BTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u4 Then
                Top = v4
                Else
                    If pickFromUnit_BTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u5 Then
                    Top = v5
                    Else
                        If pickFromUnit_BTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u6 Then
                        Top = v6
                        Else
                            If pickFromUnit_BTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u7 Then
                            Top = v7
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If pickToUnitBTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u1 Then
        bottom = v1
        Else
        If pickToUnitBTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u2 Then
            bottom = v2
        Else
            If pickToUnitBTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u3 Then
            bottom = v3
            Else
                If pickToUnitBTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u4 Then
                bottom = v4
                Else
                    If pickToUnitBTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u5 Then
                    bottom = v5
                    Else
                        If pickToUnitBTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u6 Then
                        bottom = v6
                        Else
                            If pickToUnitBTUinperhrft2F_Btuftperhrft2F_Btuinpersft2F_BtuperhrftF_WpermK = u7 Then
                            bottom = v7
                            End If
                        End If
                    End If
                End If
       End If
        End If
    End If
    
  
  convert_thermal_conductivity = number * Top / bottom
 
      
      
End Function


'density
'This function works by using the top number to convert to the first variable and the bottom number to convert to the unit you want
'last updated 5-16-14

Function convert_density(number, pickFromUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR, pickToUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR)
   
 x = 0 'this is for when the conversion involves addition


                        u1 = "lb/ft3"
                        v1 = 1
                        u2 = "SG_H2O"
                        v2 = 62.4
                        u3 = "kg/m3"
                        v3 = 0.0624279606
                        u4 = "g/cm3"
                        v4 = 62.4279606
                        If pickFromUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = "API" Then 'This is just an estimate, the story of API is very complicated. This becomes more correct the closer you are to 60F
                            number = 141.5 / (number + 131.5)
                            pickFromUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = "SG_H2O"
                            u2 = "SG_H2O"
                            v2 = 62.4
                        End If
                        u5 = "SG_AIR"
                        v5 = 0.0752256925
    
    If pickFromUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u1 Then
        Top = v1
    Else
        If pickFromUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u2 Then
            Top = v2
        Else
            If pickFromUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u3 Then
                Top = v3
            Else
                If pickFromUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u4 Then
                Top = v4
                Else
                    If pickFromUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u5 Then
                    Top = v5
                    Else
                        If pickFromUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u6 Then
                        Top = v6
                        Else
                            If pickFromUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u7 Then
                            Top = v7
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If pickToUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u1 Then
        bottom = v1
        Else
        If pickToUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u2 Then
            bottom = v2
        Else
            If pickToUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u3 Then
            bottom = v3
            Else
                If pickToUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u4 Then
                bottom = v4
                Else
                    If pickToUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u5 Then
                    bottom = v5
                    Else
                        If pickToUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u6 Then
                        bottom = v6
                        Else
                            If pickToUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = u7 Then
                            bottom = v7
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
   If pickToUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR <> "API" Then
  convert_density = number * Top / bottom
   End If
 
 If pickToUnit_lbperft3_SG_H2O_kgperm3_gpercm3_API_SG_AIR = "API" Then
  convert_density = 141.5 / (number * Top / 62.4) - 131.5
 End If

      
      
End Function

'vessel_pressure_withOD
'last updated 5/27/14

Function vessel_pressure_withOD(thicknessin, ODin, allowableStressPsi, jointEfficiency, pickShape_shell_sphere_hemi_ellipse, pick_div1_div2)

D = ODin
S = allowableStressPsi
E = jointEfficiency
Shape = pickShape_shell_sphere_hemi_ellipse
R = D / 2
t = thicknessin
division = pick_div1_div2
pressure = "error"


'these formulas were obtained from the pressure vessel handbook 13th edition
'div1
shellODP = (S * E * t) / (R - 0.4 * t)
sphereODP = 2 * S * E * t / (R - 0.8 * t)
ellipODP = 2 * S * E * t / (D - 1.8 * t)

'div2
shellODP_div2 = S * E * Log(2 * t / (D - 2 * t) + 1)
sphereODP_div2 = S * E / 0.5 * Log(2 * t / (D - 2 * t) + 1)
ellipODP_div2 = "further calculations required"


'shell
If Shape = "shell" And division = "div1" Then
 pressure = shellODP
End If
If Shape = "shell" And division = "div2" Then
 pressure = shellODP_div2
End If

'sphere
If Shape = "sphere" And division = "div1" Then
 pressure = sphereODP
End If
If Shape = "sphere" And division = "div2" Then
 pressure = sphereODP_div2
End If

'hemi
'the equations for sphere and hemi are the same
If Shape = "hemi" And division = "div1" Then
 pressure = sphereODP
End If
If Shape = "hemi" And division = "div2" Then
 pressure = sphereODP_div2
End If

'ellipse
If Shape = "ellipse" And division = "div1" Then
 pressure = ellipODP
End If
If Shape = "ellipse" And division = "div2" Then
 pressure = ellipODP_div2
End If

vessel_pressure_withOD = pressure

End Function

'finds the thickness using ASME div1 rules
'the joint efficiency is usually 1
'updated 5-27-14

Function vessel_thickness_withOD(PressurePsi, ODin, CorrosionAllowancein, allowableStressPsi, jointEfficiency, pickShape_shell_sphere_hemi_ellipse, pick_div1_div2)

D = ODin
CA = CorrosionAllowancein
S = allowableStressPsi
E = jointEfficiency
Shape = pickShape_shell_sphere_hemi_ellipse
R = D / 2
p = PressurePsi
division = pick_div1_div2

'these formulas were obtained from the pressure vessel handbook 13th edition
'div1
shellODt = p * R / (S * E + 0.4 * p) + CA
sphereODt = (p * R) / (2 * S * E + 0.8 * p) + CA
ellipODt = p * D / (2 * S * E + 1.8 * p) + CA

'div2
shellODt_div2 = D / (2 * 2.7182818 ^ (p / (S * E))) * (2.7182818 ^ (p / (S * E)) - 1) + CA
sphereODt_div2 = D / (2 * 2.7182818 ^ (0.5 * p / (S * E))) * (2.7182818 ^ (0.5 * p / (S * E)) - 1)
ellipODt_div2 = "further calculations required"

'shell
If Shape = "shell" And division = "div1" Then
 Thickness = shellODt
End If
If Shape = "shell" And division = "div2" Then
 Thickness = shellODt_div2
End If

'sphere
If Shape = "sphere" And division = "div1" Then
 Thickness = sphereODt
End If
If Shape = "sphere" And division = "div2" Then
 Thickness = sphereODt_div2
End If

'hemi
'the equations for sphere and hemi are the same
If Shape = "hemi" And division = "div1" Then
 Thickness = sphereODt
End If
If Shape = "hemi" And division = "div2" Then
 Thickness = sphereODt_div2
End If

'ellipse
If Shape = "ellipse" And division = "div1" Then
 Thickness = ellipODt
End If
If Shape = "ellipse" And division = "div2" Then
 Thickness = ellipODt_div2
End If

vessel_thickness_withOD = Thickness

End Function

'if you give it the pressure it will find you thickness and visa versa
'the joint efficiency is usually 1
'created 5-27-14

Function vessel_pressure_withID(thicknessin, IDin, allowableStressPsi, jointEfficiency, pickShape_shell_sphere_hemi_ellipse, pick_div1_div2)

D = IDin
S = allowableStressPsi
E = jointEfficiency
Shape = pickShape_shell_sphere_hemi_ellipse
R = D / 2
t = thicknessin
pressure = "error"
division = pick_div1_div2

'these formulas were obtained from the pressure vessel handbook 13th edition
'div1
shellIDP = S * E * t / (R + 0.6 * t)
sphereIDP = 2 * S * E * t / (R + 0.2 * t)
ellipIDP = 2 * S * E * t / (D + 0.2 * t)

'div2
shellIDP_div2 = S * E * Log(2 * t / D + 1)
sphereIDP_div2 = S * E / 0.5 * Log(2 * t / D + 1)
ellipIDP_div2 = "further calculations required"

'shell
If Shape = "shell" And division = "div1" Then
 pressure = shellIDP
End If
If Shape = "shell" And division = "div2" Then
 pressure = shellIDP_div2
End If

'sphere
If Shape = "sphere" And division = "div1" Then
 pressure = sphereIDP
End If
If Shape = "sphere" And division = "div2" Then
 pressure = sphereIDP_div2
End If

'hemi
'the equations for sphere and hemi are the same
If Shape = "hemi" And division = "div1" Then
 pressure = sphereIDP
End If
If Shape = "hemi" And division = "div2" Then
 pressure = sphereIDP_div2
End If

'ellipse
If Shape = "ellipse" And division = "div1" Then
 pressure = ellipIDP
End If
If Shape = "ellipse" And division = "div2" Then
 pressure = ellipIDP_div2
End If

vessel_pressure_withID = pressure

End Function

'finds the thickness using ASME div1 rules
'the joint efficiency is usually 1
'updated 5-27-14

Function vessel_thickness_withID(PressurePsi, IDin, CorrosionAllowancein, allowableStressPsi, jointEfficiency, pickShape_shell_sphere_hemi_ellipse, pick_div1_div2)

D = IDin
CA = CorrosionAllowancein
S = allowableStressPsi
E = jointEfficiency
Shape = pickShape_shell_sphere_hemi_ellipse
R = D / 2
p = PressurePsi
Thickness = "error"
division = pick_div1_div2

'these formulas were obtained from the pressure vessel handbook 13th edition
'div1
shellIDt = p * R / (S * E - 0.6 * p) + CA
sphereIDt = (p * R) / (2 * S * E - 0.2 * p) + CA
ellipIDt = p * D / (2 * S * E - 0.2 * p) + CA

'div2
shellIDt_div2 = D / 2 * (2.7182818 ^ (p / (S * E)) - 1)
sphereIDt_div2 = D / 2 * (2.7182818 ^ (0.5 * p / (S * E)) - 1)
ellipIDt_div2 = "further calculations required"

'shell
If Shape = "shell" And division = "div1" Then
 Thickness = shellIDt
End If
If Shape = "shell" And division = "div2" Then
    Thickness = shellIDt_div2
End If

'sphere
If Shape = "sphere" And division = "div1" Then
 Thickness = sphereIDt
End If
If Shape = "sphere" And division = "div2" Then
 Thickness = sphereIDt_div2
End If

'hemi
'the equations for sphere and hemi are the same
If Shape = "hemi" And division = "div1" Then
 Thickness = sphereIDt
End If
If Shape = "hemi" And division = "div2" Then
 Thickness = sphereIDt_div2
End If

'ellipse
If Shape = "ellipse" And division = "div1" Then
Thickness = ellipIDt
End If
If Shape = "ellipse" And division = "div2" Then
Thickness = ellipIDt_div2
End If
vessel_thickness_withID = Thickness

End Function
