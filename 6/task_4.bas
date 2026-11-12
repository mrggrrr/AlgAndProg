Attribute VB_Name = "task_4"
Option Explicit

Sub task_4()
    Dim A(1 To 12) As Integer
    Dim i As Integer
    Dim sum As Integer
    Dim mean_a As Integer
    Dim out_str_arrFirst As String
    Dim out_str_arrEnd As String

    out_str_arrFirst = "Начальный массив:" & Chr(10)

    For i = 1 To 12
        A(i) = Int(Rnd * 20) + 1
        out_str_arrFirst = out_str_arrFirst & A(i) & " "
        sum = sum + A(i)
    Next i

    mean_a = sum / 12

    For i = 1 To 12
        If A(i) < mean_a Then
            A(i) = 0
        ElseIf A(i) > mean_a Then
            A(i) = 1
        End If
    Next i

    out_str_arrEnd = "Конечный массив:" & Chr(10)

    For i = 1 To 12
        out_str_arrEnd = out_str_arrEnd & A(i) & " "
    Next i

    MsgBox out_str_arrFirst & Chr(10) & _
           "Среднее арифметическое: " & mean_a & Chr(10) & _
           out_str_arrEnd
End Sub

