Attribute VB_Name = "task_6"
Option Explicit

Sub task_6()
    Dim M(1 To 15) As Integer
    Dim i As Integer
    Dim K As Integer
    Dim sum As Integer
    Dim out_str_arrFirst As String
    Dim out_str_arrEnd As String

    out_str_arrFirst = "Начальный массив:" & Chr(10)

    For i = 1 To 15
        M(i) = Int(Rnd * 16)
        out_str_arrFirst = out_str_arrFirst & M(i) & " "
        
        If M(i) >= 3 And M(i) <= 11 Then
            K = K + 1
        End If
    Next i

    If K = 0 Then
        out_str_arrEnd = "Количество элементов от 3 до 11 равно 0."
    ElseIf K < 4 Then
        For i = 1 To 15
            If M(i) Mod K = 0 Then
                sum = sum + M(i)
            End If
        Next i

    out_str_arrEnd = "Сумма элементов, кратных количеству: " & sum
    
    Else
        For i = 1 To 15
            If M(i) Mod K <= 2 Then
                M(i) = M(i) ^ 2
            End If
        Next i

    out_str_arrEnd = "Конечный массив:" & Chr(10)

        For i = 1 To 15
            out_str_arrEnd = out_str_arrEnd & M(i) & " "
        Next i
    End If
    
    MsgBox out_str_arrFirst & Chr(10) & _
           "Количество элементов от 3 до 11: " & K & Chr(10) & _
           out_str_arrEnd
End Sub

