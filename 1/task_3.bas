Attribute VB_Name = "task_3"
Option Explicit

Sub Circles()
    Dim L, W, R, D, AmountCircles, SCircles, SRectangle, SWaste
    Const p = 3.14
    L = 12
    W = 1.4
    R = 0.15
    D = R * 2
    AmountCircles = Val((L / D) * (W / D))
    SCircles = p * R ^ 2
    SRectangle = L * W
    SWaste = SRectangle - (AmountCircles * SCircles)
    MsgBox "Площадь отходов: " & SWaste & "     Кол-во заготовок: " & AmountCircles
End Sub

