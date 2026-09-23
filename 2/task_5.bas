Attribute VB_Name = "task_5"
Option Explicit

Sub Module()
Dim a As Integer, b As Integer, c As Integer
a = InputBox("Введите значение a = ")
b = InputBox("Введите значение b = ")
c = InputBox("Введите значение c = ")
If Abs(a) > Abs(b) And Abs(a) > Abs(c) Then
MsgBox "Веденные значения: a = " & a & " b = " & b & " c = " & c & Chr(10) & "Большее из трех чисел a = " & a
ElseIf Abs(b) > Abs(a) And Abs(b) > Abs(c) Then
MsgBox "Веденные значения: a = " & a & " b = " & b & " c = " & c & Chr(10) & "Большее из трех чисел b = " & b
Else
MsgBox "Веденные значения: a = " & a & " b = " & b & " c = " & c & Chr(10) & "Большее из трех чисел c = " & c
End If
End Sub
