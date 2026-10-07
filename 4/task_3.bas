Attribute VB_Name = "task_3"
Option Explicit

Sub Task_3()
    Dim n As Integer
    Dim i As Integer
    Dim Sum3 As Integer
    Dim Sum4 As Integer
    Dim r As Integer
    Dim Posled As String
    Dim Mean As Double
    
        n = Int(Rnd * 90) + 10
        i = 1
        Sum3 = 0
        Sum4 = 0
        
            Do While i <= n
            
                r = Int(Rnd * 90) + 10
                
                Posled = Posled & r & " "
                
                    If r Mod 3 = 0 Then
                        Sum3 = Sum3 + r
                    End If
                    
                    If r Mod 4 = 0 Then
                        Sum4 = Sum4 + r
                    End If
                    
                i = i + 1
            Loop
            
            If Sum3 > Sum4 Then
                Mean = (Sum3 + Sum4) / 2
                
                MsgBox "Последовательность:" & Chr(10) & Posled & Chr(10)
                MsgBox "Сумма чисел, кратных 3: " & Sum3 & Chr(10)
                MsgBox "Сумма чисел, кратных 4: " & Sum4 & Chr(10)
                MsgBox "Сумма чисел, кратных 3 БОЛЬШЕ." & "Выводим СРЕДНЕЕ АРИФМЕТИЧЕСКОЕ СУММ:" & Chr(10) & Mean
            Else
                Mean = (Sum3 + Sum4) * 1 / 2
                
                MsgBox "Последовательность:" & Chr(10) & Posled & Chr(10)
                MsgBox "Сумма чисел, кратных 3: " & Sum3 & Chr(10)
                MsgBox "Сумма чисел, кратных 4: " & Sum4 & Chr(10)
                MsgBox "Сумма чисел, кратных 3 МЕНЬШЕ." & "Выводим СРЕДНЕЕ ГЕОМЕТРИЧЕСКОЕ СУММ:" & Chr(10) & Mean
            End If
End Sub

