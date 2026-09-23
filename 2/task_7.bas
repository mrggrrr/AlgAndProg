Attribute VB_Name = "task_7"
Option Explicit
Sub Seasons()
    Dim Month As Integer, Seasons As String, Day As String
        
        Month = InputBox("Введите номер месяца: ")
        
        Select Case Month
        Case 12, 1, 2
        Seasons = "Время года - ЗИМА"
        Case 3, 4, 5
        Seasons = "Время года - ВЕСНА"
        Case 6, 7, 8
        Seasons = "Время года - ЛЕТО"
        Case 9, 10, 11
        Seasons = "Время года - ОСЕНЬ"
        Case Else
        Seasons = "Такого времени года НЕ СУЩЕСТВУЕТ"
        End Select

        Select Case Month
        Case 1, 3, 5, 7, 8, 10, 12
        Day = "Количество дней - 31"
        Case 4, 6, 9, 11
        Day = "Количество дней - 30"
        Case 2
        Day = "Количество дней - 28"
        Case Else
        Day = "Месяца с таким количеством дней НЕ СУЩЕСТВУЕТ"
        
        End Select

MsgBox Seasons & Chr(10) & Day
End Sub
