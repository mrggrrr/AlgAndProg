Attribute VB_Name = "Task_1"
Option Explicit

Sub SumRow()
    Dim n As Integer
    Dim i As Integer
    Dim S As Double
    
        n = InputBox("Введите количество членов ряда n = :")
        i = 1
        S = 0
            Do While i <= n
                S = S + 1 / i ^ i
                i = i + 1
            Loop
        MsgBox "Сумма ряда S = 1 + 1 / i ^ i" & Chr(10) & S
End Sub
