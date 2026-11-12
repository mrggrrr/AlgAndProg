Attribute VB_Name = "task_2"
Option Explicit

Sub Task_2()
    
    Dim i As Integer
    Dim j As Integer
    Dim M As Integer
    Dim N As Integer
    Dim Row As String
    Dim a As Integer
    Dim kZero As Integer
    
    N = InputBox("Введите кол-во СТРОК матрицы: ")
    M = InputBox("Введите кол-во СТОЛБЦОВ матрицы")
    
    kZero = 0
    
    For i = 1 To N
        For j = 1 To M
            a = Int(Rnd * 11) - 5
            Row = Row & a & Chr(10)
            
            If a = 0 Then
                kZero = kZero + 1
            End If
        Next j
            Row = Row & Chr(10)
    Next i
        
        MsgBox "Матрица: " & Row & Chr(10) & "Количество нулей: " & kZero
            
End Sub
