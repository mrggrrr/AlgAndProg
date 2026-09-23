Attribute VB_Name = "task_2"
Option Explicit

Sub YarnCal()
    Dim Yarn, Waste, Losses
    Const Amount = 12000
    Yarn = 0.93
    Waste = 0.06
    Losses = 0.01
    Yarn = Yarn * Amount
    Waste = Waste * Amount
    Losses = Losses * Amount
    MsgBox "Количество Пряжи: " & Yarn & "     Отходов: " & Waste & "     Потерь: " & Losses
End Sub
