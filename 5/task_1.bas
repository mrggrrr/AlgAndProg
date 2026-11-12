Attribute VB_Name = "task_1"
Option Explicit

Sub Task_1()
    
    Dim i As Integer
    Dim j As Integer
    Dim Row As String
        
        For i = 2 To 9
            For j = 1 To 10
                Row = Row & i & " * " & j & " = " & i * j & Chr(10)
            Next j
                Row = Row & Chr(10)
        Next i
        
        MsgBox Row
            
End Sub
