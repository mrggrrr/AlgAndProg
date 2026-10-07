Attribute VB_Name = "task_6"
Option Explicit

Sub Task_6()
    Dim n As Integer
    Dim i As Integer
    Dim r As Integer
    Dim Row As String
    Dim S As Integer
    Dim Mean As Double
    Dim K_Even As Integer
    Dim K_Odd As Double
    
        n = Int(Rnd * 90) + 10
        i = 1
        S = 0
        K_Even = 0
        K_Odd = 0
        
            Do
                r = Int(Rnd * 90) + 10
                Row = Row & r & " "
                
                S = S + r
                
                If r Mod 2 = 0 Then
                    K_Even = K_Even + 1
                Else
                    K_Odd = K_Odd + 1
                End If
                
                i = i + 1
            Loop Until i > n
            
        Mean = S / n
        
        If K_Even > K_Odd Then
            Mean = Mean * 5
            MsgBox "Последовательность:" & Chr(10) & Row & Chr(10) & "Кол-во четных: " & K_Even & Chr(10) & "Кол-во нечетных: " & K_Odd & Chr(10) & "Количество ЧЕТНЫХ БОЛЬШЕ, умножаем на 5: " & Mean
        Else
            Mean = Mean ^ (1 / 5)
            MsgBox "Последовательность:" & Chr(10) & Row & Chr(10) & "Кол-во четных: " & K_Even & Chr(10) & "Кол-во нечетных: " & K_Odd & Chr(10) & "Количество ЧЕТНЫХ МЕНЬШЕ, возводим в степень: " & Mean
        End If
            
End Sub

