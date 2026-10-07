Attribute VB_Name = "task_2"
Sub Factorial()
    Dim n As Integer
    Dim i As Integer
    Dim F As Double
    
        n = InputBox("Введите число n = :")
        i = 1
        F = 1
            Do While i <= n
                F = F * i
                i = i + 1
            Loop
        MsgBox "ФАКТОРИАЛ ЧИСЛА n:" & Chr(10) & F
End Sub
