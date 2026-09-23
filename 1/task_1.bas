Attribute VB_Name = "task_1"
Option Explicit

Sub Cube()
    Dim L, V, S As Integer
    L = InputBox("Введите длину ребра:", "Длина ребра")
    V = L ^ 3
    S = L ^ 2 * 6
    MsgBox "Объем куба: " & V & "       Площадь поверхности: " & S
End Sub
