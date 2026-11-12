Attribute VB_Name = "task_9"
Option Explicit

Sub task_9()
    Dim K(1 To 8) As Integer
    Dim i As Integer
    Dim j As Integer
    Dim temp As Integer
    Dim out_str_arrFirst As String
    Dim out_str_arrEnd As String

    out_str_arrFirst = "Начальный массив:" & Chr(10)

    For i = 1 To 8
        K(i) = Int(Rnd * 20) + 1
        out_str_arrFirst = out_str_arrFirst & K(i) & " "
    Next i

    For i = 1 To 7
        For j = i + 1 To 8
            If K(i) < K(j) Then
                temp = K(i)
                K(i) = K(j)
                K(j) = temp
            End If
        Next j
    Next i

    out_str_arrEnd = "Массив по убыванию:" & Chr(10)

    For i = 1 To 8
        out_str_arrEnd = out_str_arrEnd & K(i) & " "
    Next i

    MsgBox out_str_arrFirst & Chr(10) & out_str_arrEnd
End Sub
