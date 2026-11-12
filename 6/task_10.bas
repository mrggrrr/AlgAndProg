Attribute VB_Name = "task_10"
Option Explicit

Sub task_10()
    Dim B(1 To 15) As Integer
    Dim i As Integer
    Dim first_even As Integer
    Dim second_even As Integer
    Dim temp As Integer
    Dim count As Integer
    Dim out_str_arrFirst As String
    Dim out_str_arrEnd As String

    out_str_arrFirst = "Начальный массив:" & Chr(10)

    For i = 1 To 15
        B(i) = Int(Rnd * 21)
        out_str_arrFirst = out_str_arrFirst & B(i) & " "
    Next i

    first_even = 0
    second_even = 0

    For i = 1 To 15
        If B(i) Mod 2 = 0 Then
            If first_even = 0 Then
                first_even = i
            Else
                second_even = i
                Exit For
            End If
        End If
    Next i

    If first_even <> 0 And second_even <> 0 Then
        temp = B(first_even)
        B(first_even) = B(second_even)
        B(second_even) = temp
    End If

 
    count = 0

    For i = 1 To 14
        If B(i) Mod 2 = 0 And B(i + 1) Mod 2 = 0 Then
            count = count + 1
        End If
    Next i

   
    For i = 1 To 15 Step 2
        B(i) = count
    Next i

   
    out_str_arrEnd = "Конечный массив:" & Chr(10)

    For i = 1 To 15
        out_str_arrEnd = out_str_arrEnd & B(i) & " "
    Next i

    MsgBox out_str_arrFirst & Chr(10) & _
           "Количество соседств четных элементов: " & count & Chr(10) & _
           out_str_arrEnd
End Sub
