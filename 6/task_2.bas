Attribute VB_Name = "task_2"
Option Explicit


Sub task_2()
    Dim M(1 To 15) As Integer
    Dim i As Integer
    Dim sum As Integer
    Dim operation As Double
    Dim mean_a As Double
    Dim mean_g As Double
    Dim diff As Integer
    Dim sum_sqr As Integer
    Dim out_str_arrFirst As String
    Dim out_str_arrEnd As String
    
    out_str_arrFirst = "Начальный массив: " & Chr(10)
    
    For i = 1 To 15
        M(i) = Int(Rnd * 20) + 1
        out_str_arrFirst = out_str_arrFirst & M(i) & " "
    Next i
    
    For i = 1 To 15
        sum = sum + M(i)
    Next i
    mean_a = sum / 15
    
    operation = 1
    
    For i = 1 To 15
        operation = operation * M(i)
    Next i
    mean_g = operation ^ (1 / 15)
    
    diff = Abs(mean_a - mean_g)
    
    out_str_arrEnd = "Конечный массив: " & Chr(10)
    
    If diff Mod 2 = 0 Then
        For i = 1 To 15
            If i Mod 2 = 0 Then
                M(i) = 0
            End If
            out_str_arrEnd = out_str_arrEnd & M(i) & " "
        Next i
        
        MsgBox out_str_arrFirst & Chr(10) & "Разность кратна 2: " & diff & Chr(10) & out_str_arrEnd
    
    
    Else
        For i = 1 To 15
            M(i) = M(i) ^ 2
            sum_sqr = sum_sqr + M(i)
            
            out_str_arrEnd = out_str_arrEnd & M(i) & " "
        Next i
        
            MsgBox out_str_arrFirst & Chr(10) & "Разность среднего: " & diff & Chr(10) & out_str_arrEnd & Chr(10) & "Сумма квадратов массива: " & sum_sqr
        
    End If
    
    
End Sub

