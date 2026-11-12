Attribute VB_Name = "task_5"
Option Explicit

Sub task_5()
    Dim C(1 To 8) As Integer
    Dim B(1 To 8) As Integer
    Dim i As Integer
    Dim out_str_arrFirst As String
    Dim out_str_arrEnd As String

    out_str_arrFirst = "Начальные массивы:" & Chr(10)
    out_str_arrEnd = "Конечные массивы:" & Chr(10)

    For i = 1 To 8
        C(i) = Int(Rnd * 20) + 1
        B(i) = Int(Rnd * 20) + 1

        out_str_arrFirst = out_str_arrFirst & _
            "C(" & i & ") = " & C(i) & "    B(" & i & ") = " & B(i) & Chr(10)
    Next i

    For i = 1 To 8
        If C(i) Mod 5 = 0 Then
            B(i) = B(i) * 5
        Else
            B(i) = 0
        End If

        out_str_arrEnd = out_str_arrEnd & _
            "C(" & i & ") = " & C(i) & "    B(" & i & ") = " & B(i) & Chr(10)
    Next i

    MsgBox out_str_arrFirst & Chr(10) & out_str_arrEnd
End Sub
