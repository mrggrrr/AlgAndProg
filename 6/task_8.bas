Attribute VB_Name = "task_8"
Option Explicit

Sub task_8()
    Dim C(1 To 20) As Integer
    Dim i As Integer
    Dim n As Integer
    Dim P As Double
    Dim temp As Double
    Dim out_str_arrFirst As String
    Dim out_str_arrEnd As String

    n = 10
    P = Val(InputBox("Введите число P:"))

    out_str_arrFirst = "Начальный массив:" & Chr(10)

    For i = 1 To n
        C(i) = Rnd() * 50 - 25
        out_str_arrFirst = out_str_arrFirst & _
            C(i) & " "
    Next i

    i = n - 1

    Do While i >= 1
        If C(i) * C(i + 1) > 0 Then
            
            For temp = n To i + 1 Step -1
                C(temp + 1) = C(temp)
            Next temp

            C(i + 1) = P
            n = n + 1
        End If

        i = i - 1
    Loop

    out_str_arrEnd = "Конечный массив:" & Chr(10)

    For i = 1 To n
        out_str_arrEnd = out_str_arrEnd & _
            C(i) & " "
    Next i

    MsgBox out_str_arrFirst & Chr(10) & out_str_arrEnd
End Sub
