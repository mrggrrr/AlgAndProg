Attribute VB_Name = "task_2"
Option Explicit

Sub Mean()
Dim a As Integer, b As Integer, c As Integer, d As Integer, Mean

    a = InputBox("¬ведите A")
    b = InputBox("¬ведите B")
    c = InputBox("¬ведите C")
    d = InputBox("¬ведите D")
    
    If a * b > c * d Then
        Mean = (a + b + c + d) / 4
        MsgBox "—реднее арифмитическое " & Mean
    Else
        Mean = (a * b * c * d) ^ (1 / 4)
        MsgBox "—реднее геометрическое " & Mean
    End If

End Sub

