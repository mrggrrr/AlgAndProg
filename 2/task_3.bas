Attribute VB_Name = "task_3"
Option Explicit

Sub Triangle()
Dim a As Integer, b As Integer, c As Integer
a = InputBox("Введите значение a = ")
b = InputBox("Введите значение b = ")
c = InputBox("Введите значение c = ")
If (a + b) > c And (a + c) > b And (b + c) > a Then
MsgBox "Веденные значения: a = " & a & " b = " & b & "c = " & c & Chr(10) & "Треугольник ВОЗМОЖЕН"
Else
MsgBox "Веденные значения: a = " & a & " b = " & b & "c = " & c & Chr(10) & "Треугольник НЕ ВОЗМОЖЕН"
End If
End Sub
