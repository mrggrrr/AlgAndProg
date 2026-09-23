Attribute VB_Name = "task_6"
Option Explicit

Sub FigCalc()

Dim Figure As Integer
Dim a1 As Integer, b1 As Integer, c1 As Integer
Dim a2 As Integer, b2 As Integer
Dim a3 As Integer
Dim r As Integer
Dim a5 As Integer, b5 As Integer, c5 As Integer, d5 As Integer, h As Integer
Dim a6 As Integer, angle As Integer
Dim P As Double, S As Double, Pp As Double


   Const Pi As Double = 3.14159
    
    Figure = InputBox("Выберите одну из 6 фигур: " & Chr(10) & "1) Треугольник" & Chr(10) & "2) Прямоугольник" & Chr(10) & "3) Квадрат" & Chr(10) & "4) Круг" & Chr(10) & "5) Трапеция" & Chr(10) & "6) Ромб")
    
     If Figure < 1 Or Figure > 6 Then
        MsgBox "Больше фигур НЕТ" & Chr(10) & "Введите число из списка!!!"
     
     
     ElseIf Figure = 1 Then
        a1 = InputBox("Введите сторону треугольника a1 = ")
        b1 = InputBox("Введите сторону треугольника b1 = ")
        c1 = InputBox("Введите сторону треугольника c1 = ")
        
        P = a1 + b1 + c1
        Pp = (a1 + b1 + c1) / 2
        S = (Pp * (Pp - a1) * (Pp - b1) * (Pp - c1)) * 1 / 2
     
        If (a1 + b1) > c1 And (a1 + c1) > b1 And (b1 + c1) > a1 Then
           MsgBox "Веденные значения: a1 = " & a1 & " b1 = " & b1 & "c1 = " & c1 & Chr(10) & "Треугольник ВОЗМОЖЕН"
           MsgBox "1) Треугольник" & Chr(10) & "Периметр = " & P & " Площадь = " & S
        Else
           MsgBox "Веденные значения: a1 = " & a1 & " b1 = " & b1 & "c1 = " & c1 & Chr(10) & "Треугольник НЕ ВОЗМОЖЕН"
        End If
     
     ElseIf Figure = 2 Then
        a2 = InputBox("Введите сторону прямоугольника a2 = ")
        b2 = InputBox("Введите сторону прямоугольника b2 = ")
        
        P = 2 * (a2 + b2)
        S = a2 * b2
        
        MsgBox "Веденные значения: a2 = " & a2 & " b2 = " & b2 & Chr(10) & "2) Прямоугольник" & Chr(10) & "Периметр = " & P & " Площадь = " & S
     
     
     ElseIf Figure = 3 Then
        a3 = InputBox("Введите сторону квадрата a3 = ")
        
        P = 4 * a3
        S = a3 ^ 2
        MsgBox "Веденные значения: a3 = " & a3 & Chr(10) & "3) Квадрат" & Chr(10) & "Периметр = " & P & " Площадь = " & S
     
     
     ElseIf Figure = 4 Then
        r = InputBox("Введите радиус r = ")
        
        P = 2 * Pi * r
        S = Pi * r ^ 2
        
        MsgBox "Веденные значения: радиус r = " & r & Chr(10) & "4) Круг" & Chr(10) & "Периметр = " & P & " Площадь = " & S
     
     ElseIf Figure = 5 Then
        a5 = InputBox("Введите основание трапеции a5 = ")
        b5 = InputBox("Введите боковую сторону трапеции b5 = ")
        c5 = InputBox("Введите основание трапеции с5 = ")
        d5 = InputBox("Введите боковую сторону трапеции d5 = ")
        
        P = a5 + b5 + c5 + d5
        Pp = (Abs(a5 - b5) + c5 + d5) / 2
        h = (2 / Abs(a5 - b5)) * (Pp * (Pp - Abs(a5 - b5) * (Pp - c5) * (Pp - d5))) * 1 / 2
        S = ((a5 + b5) / 2) * h
        
        MsgBox "Веденные значения: a5 = " & a5 & " b5 = " & b5 & " c5 = " & c5 & " d5 = " & d5 & Chr(10) & "5) Трапеция" & Chr(10) & "Периметр = " & P & " Площадь = " & S
        
    ElseIf Figure = 6 Then
        a6 = InputBox("Введите сторону ромба a6 = ")
        angle = InputBox("Введите угол ромба = ")
        
        S = a6 ^ 2 * Sin(angle)
        P = 4 * a6
        
        MsgBox "Веденные значения: a6 = " & a6 & " угол = " & angle & Chr(10) & "6) Ромб" & Chr(10) & "Периметр = " & P & " Площадь = " & S

     End If
End Sub

