Attribute VB_Name = "task_3"
Option Explicit

Sub task_3()
    Dim C(1 To 15) As Integer
    Dim i As Integer
    Dim sum As Integer
    Dim mean As Double
    Dim deviation As Double
    Dim max_deviation As Double
    Dim out_str As String

    out_str = "Массив C:" & Chr(10)

    For i = 1 To 15
        C(i) = Int(Rnd * 20) + 1
        out_str = out_str & C(i) & " "
        sum = sum + C(i)
    Next i

    mean = sum / 15

    max_deviation = C(1) - mean

    For i = 2 To 15
        deviation = C(i) - mean

        If deviation > max_deviation Then
            max_deviation = deviation
        End If
    Next i

    MsgBox out_str & Chr(10) & _
           "Среднее арифметическое: " & mean & Chr(10) & _
           "Наибольшее отклонение: " & max_deviation
End Sub

