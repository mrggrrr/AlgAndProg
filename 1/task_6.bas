Attribute VB_Name = "task_6"
Option Explicit

Sub MachineProductivity()
    Dim T1, T6, AmountDishes, AmountMachine
    T1 = 60
    T6 = T1 * 6
    AmountDishes = 7
    AmountDishes = AmountDishes * T6
    AmountMachine = 3
    AmountDishes = AmountDishes * AmountMachine
    MsgBox "Кол-во тарелок тремя машинами за 6 часов: " & AmountDishes
End Sub
