Attribute VB_Name = "task_5"
Option Explicit

Sub Task_5()
    Dim n As Integer
    Dim i As Integer
    Dim r As Integer
    Dim Row As String
    Dim S As Integer
    Dim Mean As Double
    Dim Max As Integer
    Dim Min As Integer
    Dim MaxDeviation As Double
    
        n = Int(Rnd * 90) + 10
                
        Min = 9999
        
        i = 1
        S = 0
        
        MaxDeviation = 0
        
            Do
                r = Int(Rnd * 90) + 10
                Row = Row & r & " "
                
                S = S + r
                
                If r > Max Then
                    Max = r
                ElseIf r < Min Then
                    Min = r
                End If
                
                i = i + 1
            Loop Until i > n
            
            Mean = S / n
            
            MaxDeviation = Abs(Max - Mean)
            
            If Abs(Min - Mean) > MaxDeviation Then
                MaxDeviation = Abs(Min - Mean)
            End If
            

            
        MsgBox "Последовательность:" & Chr(10) & Row & Chr(10) & "Сумма: " & S & Chr(10) & "Среднее арифметическое: " & Mean & Chr(10) & "Максимальное отклонение от среднего: " & MaxDeviation & Chr(10) & "Максимум: " & Max & Chr(10) & "Минимум: " & Min
End Sub

