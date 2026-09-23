Attribute VB_Name = "task_2"
Option Explicit

Sub Mean()
Dim a As Integer, b As Integer, c As Integer, d As Integer, Mean

    a = InputBox("Введите A")
    b = InputBox("Введите B")
    c = InputBox("Введите C")
    d = InputBox("Введите D")
    
    If a * b > c * d Then
        Mean = (a + b + c + d) / 4
        MsgBox "Среднее арифмитическое " & Mean
    Else
        Mean = (a * b * c * d) ^ (1 / 4)
        MsgBox "Среднее геометрическое " & Mean
    End If

End Sub

