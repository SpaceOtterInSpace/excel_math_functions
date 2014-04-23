'used to calculate a partially filled circle


Function circ_segment(output, radius, partial_area)
'the output will be either "h" for height of liquid or "s" for arc segment that is wet
' the radius is the radius of the circle
' the partial_area is the amount of area covered in water

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
