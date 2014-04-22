'used to calculate a partially filled circle


Function pheight(radius, partial_area)
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

pheight = radius - radius * Cos(s / (2 * radius))


End Function
