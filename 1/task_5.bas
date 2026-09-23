Attribute VB_Name = "task_5"
Option Explicit

Sub Proportion()
    Dim A, B, C, proportion_A, proportion_B, proportion_C
    A = Val(InputBox("Введите A, A НЕ РАВНО 0: "))
    B = Val(InputBox("Введите B, B НЕ РАВНО 0: "))
    C = Val(InputBox("Введите C, C НЕ РАВНО 0: "))
    proportion_A = (A / (A + (B - C))) * 100
    proportion_B = (B / (A + (B - C))) * 100
    proportion_C = (C / (A + (B - C))) * 100
    MsgBox "A= " & proportion_A & "%" & "  B= " & proportion_B & "%" & "  C=" & proportion_C & "%"
End Sub
