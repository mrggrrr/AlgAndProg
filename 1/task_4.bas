Attribute VB_Name = "task_4"
Option Explicit

Sub Distance_A1A2()
    Dim x1, y1, x2, y2, A1A2
    x1 = Val(InputBox("Введите x1: "))
    y1 = Val(InputBox("Введите y1: "))
    x2 = Val(InputBox("Введите x2: "))
    y2 = Val(InputBox("Введите y2: "))
    A1A2 = Val((((x2 - x1) ^ 2) + ((y2 - y1) ^ 2)) ^ (1 / 2))
    MsgBox "Расстояние между двумя точками: " & A1A2
End Sub
