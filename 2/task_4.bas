Attribute VB_Name = "task_4"
Option Explicit

Sub Max()
Randomize
Dim a1 As Integer, a2 As Integer, a3 As Integer, a4 As Integer, a5 As Integer, Max As Integer

    a1 = InputBox("Введите a1 = ")
    a2 = InputBox("Введите a2 = ")
    a3 = InputBox("Введите a3 = ")
    
    If a1 > a2 And a1 > a3 Then
    Max = a1
    ElseIf a2 > a1 And a2 > a3 Then
    Max = a2
    Else
    Max = a3
    End If
    MsgBox "Введеные значения a1 = " & a1 & " a2 = " & a2 & " a3 = " & a3 & Chr(10) & "Максимум из 3 чисел = " & Max
    
    a4 = Int(100 * Rnd + 1)
    MsgBox "+ 1 РАНДОМНОЕ число" & Chr(10) & "a4 = " & a4
    
    If a1 > a2 And a1 > a3 And a1 > a4 Then
    Max = a1
    ElseIf a2 > a1 And a2 > a3 And a2 > a4 Then
    Max = a2
    ElseIf a3 > a1 And a3 > a2 And a3 > a4 Then
    Max = a3
    Else
    Max = a4
    End If
    MsgBox "Введеные значения a1 = " & a1 & " a2 = " & a2 & " a3 = " & a3 & " a4 = " & a4 & Chr(10) & "Максимум из 4 чисел = " & Max
    
    a5 = Int(100 * Rnd + 1)
    MsgBox "+ 1 РАНДОМНОЕ число" & Chr(10) & "a5 = " & a5
    
    If a1 > a2 And a1 > a3 And a1 > a4 And a1 > a5 Then
    Max = a1
    ElseIf a2 > a1 And a2 > a3 And a2 > a4 And a2 > a5 Then
    Max = a2
    ElseIf a3 > a1 And a3 > a2 And a3 > a4 And a3 > a5 Then
    Max = a3
    ElseIf a4 > a1 And a3 > a2 And a3 > a4 And a4 > a5 Then
    Max = a4
    Else
    Max = a5
    End If
    MsgBox "Введеные значения a1 = " & a1 & " a2 = " & a2 & " a3 = " & a3 & " a4 = " & a4 & " a5 = " & a5 & Chr(10) & "Максимум из 5 чисел = " & Max
End Sub
