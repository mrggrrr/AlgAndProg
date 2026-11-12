Attribute VB_Name = "task_1"
Option Explicit


Sub task_1()
    Dim A(1 To 15) As Integer
    Dim i As Integer
    Dim k_two As Integer
    Dim k_even As Integer
    Dim k_odd As Integer
    Dim odd_sum As Integer
    Dim mean As Double
    Dim out_str_arrFirst As String
    Dim out_str_arrEnd As String
    
    out_str_arrFirst = "Начальный массив: " & Chr(10)
    
    For i = 1 To 15
        A(i) = Int(Rnd * 26)
        out_str_arrFirst = out_str_arrFirst & A(i) & " "
        
        If A(i) = 2 Then
            k_two = k_two + 1
        End If
    Next i
    
    out_str_arrEnd = "Конечный массив: " & Chr(10)
    
    If k_two > 0 Then
        For i = 1 To 15
            If A(i) Mod 2 = 0 Then
                A(i) = 0
                k_even = k_even + 1
            End If
            out_str_arrEnd = out_str_arrEnd & A(i) & " "
        Next i
        
        MsgBox out_str_arrFirst & Chr(10) & "Количество четных: " & k_even & Chr(10) & out_str_arrEnd
    
    
    Else
        For i = 1 To 15
            If A(i) Mod 2 <> 0 Then
                k_odd = k_odd + 1
                odd_sum = odd_sum + A(i)
            End If
            out_str_arrEnd = out_str_arrEnd & A(i) & " "
        Next i
        If k_odd > 0 Then
            mean = odd_sum / k_odd
            MsgBox out_str_arrFirst & Chr(10) & "Среднее арифметическое: " & mean & Chr(10) & out_str_arrEnd
        Else
            
            MsgBox out_str_arrFirst & Chr(10) & out_str_arrEnd & Chr(10) & "Нет нечетных элементов"
        End If
        
    End If
    
    
End Sub
