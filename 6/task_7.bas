Attribute VB_Name = "task_7"
Option Explicit

Sub task_7()
    Dim C(1 To 15) As Integer
    Dim i As Integer
    Dim operation As Double
    Dim mean_g As Double
    Dim deviation As Double
    Dim max_deviation As Integer
    Dim out_str_arrFirst As String
    Dim out_str_arrEnd As String

    out_str_arrFirst = "Начальный массив:" & Chr(10)

    For i = 1 To 15
        C(i) = Int(Rnd * 20) + 1
        out_str_arrFirst = out_str_arrFirst & C(i) & " "
    Next i

    operation = 1
    For i = 1 To 15
        operation = operation * C(i)
    Next i
    mean_g = operation ^ (1 / 15)


    max_deviation = C(1) - mean_g

    For i = 2 To 15
        deviation = C(i) - mean_g

        If deviation > max_deviation Then
            max_deviation = deviation
        End If
    Next i

    
    If max_deviation Mod 2 = 0 Then
        For i = 1 To 15
            If i Mod 3 = 0 Then
                C(i) = max_deviation
            End If
        Next i
    Else
        For i = 1 To 15
            If i Mod 5 = 0 Then
                C(i) = C(i) ^ 2
            End If
        Next i
    End If

   
    out_str_arrEnd = "Конечный массив:" & Chr(10)

    For i = 1 To 15
        out_str_arrEnd = out_str_arrEnd & C(i) & " "
    Next i

    MsgBox out_str_arrFirst & Chr(10) & _
           "Среднее геометрическое: " & Round(mean_g, 4) & Chr(10) & _
           "Наибольшее отклонение: " & Round(max_deviation, 4) & Chr(10) & _
           out_str_arrEnd
End Sub
