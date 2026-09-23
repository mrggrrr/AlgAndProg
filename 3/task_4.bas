Attribute VB_Name = "task_4"
Option Explicit

Sub OpDetect()
Dim phone, opcode
    
    phone = InputBox("Введите номер")
    
    If Len(phone) < 10 Or Len(phone) > 10 Then
        MsgBox "Неправильный номер"
        Exit Sub
    End If
    
    
    opcode = Mid(phone, 1, 3)

    Select Case opcode
    
    Case "900"
        MsgBox "TELE2"
    Case "902" Or "926"
        MsgBox "Мегафон"
    Case "902"
        MsgBox "Билайн"
    Case "999"
        MsgBox "Yota"
    Case Default
        MsgBox "Низвестный оператора"
    End Select
    
End Sub

