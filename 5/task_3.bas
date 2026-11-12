Attribute VB_Name = "task_3"
Option Explicit

Sub Task_3()
    
    Dim i As Integer
    Dim j As Integer
    Dim M As Integer
    Dim N As Integer
    Dim Row As String
    Dim a As Integer
    Dim Max As Integer
    Dim Min As Integer
    
    N = InputBox("Введите кол-во СТРОК матрицы: ")
    M = InputBox("Введите кол-во СТОЛБЦОВ матрицы")
    
    
    For i = 1 To N
        For j = 1 To M
            a = Int(Rnd * 101) - 50
            Row = Row & a & Chr(10)
            
            If i = 1 And j = 1 Then
                Max = a
                Min = a
            Else
                If a > Max Then Max = a
                If a < Min Then Min = a
            End If
        Next j
            Row = Row & Chr(10)
    Next i
        
        MsgBox "Матрица: " & Row & Chr(10) & "Максимум: " & Max & Chr(10) & "Минимум: " & Min
            
End Sub
