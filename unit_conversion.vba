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
                        If from_unit = "API" Then
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
    
   If to_unit <> "API" Then
    unit_convert = number * Top / bottom
   End If
 
 If to_unit = "API" Then
    unit_convert = 141.5 / (number * Top / 62.4) - 131.5
 End If

      
      
End Function

