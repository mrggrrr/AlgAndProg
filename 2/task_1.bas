Attribute VB_Name = "task_1"
Option Explicit

Sub CondSwap()
Dim a As Integer, b As Integer, c As Integer

    a = InputBox("¬ведите A")
    b = InputBox("¬ведите B")
    
    If a > b Then
        c = a
        a = b
        b = c
    End If

    MsgBox "A = " & a & " B= " & b
End Sub
