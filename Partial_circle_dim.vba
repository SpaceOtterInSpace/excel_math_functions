'used to calculate a partially filled circle

Function pheight(radius, partial_area)
i = 0
a = 0
b = 2 * radius * 3.141592654
tol = 0.1
s = 0


Do While i <= 100
s = a + (b - a) / 2

'area = (radius * radius * (2 * (Application.Acos((radius - p) / radius)) - Application.Sin(2 * (Application.Acos((radius - p) / radius))))) / 2
Phi = s / radius
area = radius * radius * (Phi - Sin(Phi)) / 2

    If partial_area < area Then
        b = a + (b - a) / 2
        Else
        a = b - (b - a) / 2
    End If
i = i + 1



Loop

pheight = radius - radius * Cos(s / (2 * radius))


'FA = (radius * radius * (2 * (Acos((radius - a) / radius)) - Sin(2 * (Acos((radius - a) / radius))))) / 2
'FB = (radius * radius * (2 * (Acos((radius - b) / radius)) - Sin(2 * (Acos((radius - b) / radius))))) / 2

'If FA * area > 0 Then
   ' a = p
    'FA = area

End Function
