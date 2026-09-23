Attribute VB_Name = "task_2"
Option Explicit

Sub NumberDigit()

Dim x

    x = InputBox("Введите число x =")

    Select Case x

    Case Is < 10
        MsgBox "Одноразрядное число"

    Case Is < 100
        MsgBox "Двухразрядное число"

    Case Is < 1000
        MsgBox "Трехразрядное число"

    Case Is < 10000
        MsgBox "Четырехразрядное число"

    Case Is < 100000
        MsgBox "Пятиразрядное число"

    Case Else
        MsgBox "ОШИБКА :((("

End Select

End Sub
