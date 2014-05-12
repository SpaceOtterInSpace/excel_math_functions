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
