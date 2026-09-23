Attribute VB_Name = "task_1"
Option Explicit



Function Eq1(x As Double) As Double
    If (x < 0 Or x > 1) Then
        MsgBox "X не входит в множество определения функции"
        Eq1 = 0
        Exit Function
    End If
    
    Select Case x
    Case Is < 0.2
        Eq1 = 1 + WorksheetFunction.Ln(1 + x)
    Case 0.2 To 0.8
        Eq1 = (1 + x ^ (1 / 2)) / (1 + x)
    Case Is > 0.8
        Eq1 = 2 * Exp(-2 * x)
    End Select
    
End Function

Function Eq2(x As Double) As Double
    If (x < -2 Or x > 2) Then
        MsgBox "X не входит в множество определения функции"
        Eq2 = 0
        Exit Function
    End If
    
    Select Case x
    Case Is <= -1
        Eq2 = (1 + Abs(x)) / ((1 + x + x ^ 2) ^ (1 / 3))
        
    Case -1 To 0
        Eq2 = 2 * WorksheetFunction.Ln(1 + x ^ 2) + (1 + Cos(x) ^ 4) / (2 + x)

    Case Is >= 0
        Eq2 = (1 + x) ^ (3 / 5)
    End Select
    
End Function

Function Eq3(x As Double) As Double
    If (x < -2 Or x > 2) Then
        MsgBox "X не входит в множество определения функции"
        Eq3 = 0
        Exit Function
    End If
    
    Select Case x
    Case Is <= 0
        Eq3 = (1 + x) / ((1 + x ^ 2) ^ (1 / 3))
        
    Case 0 To 1
        Eq3 = -x + 2 * Exp(-2 * x)
        
    Case Is >= 1
        Eq3 = Abs(2 - x) ^ (1 / 3)
    End Select
    
End Function

Function Eq4(x As Double) As Double
    If (x < -2 Or x > 2) Then
        MsgBox "X не входит в множество определения функции"
        Eq4 = 0
        Exit Function
    End If
    
    Select Case x
    Case Is < 0
        Eq4 = 3 * x + Sqr(1 + x ^ 2)
        
    Case 0 To 1
        Eq4 = 2 * Cos(x) * Exp(-2 * x)
        
    Case Is > 1
        Eq4 = 2 * Sin(3 * x)
    End Select
    
End Function

Function Eq5(x As Double) As Double
    If (x < -1.7 Or x > 1.5) Then
        MsgBox "X не входит в множество определения функции"
        Eq5 = 0
        Exit Function
    End If
    
    ' При x = -1 знаменатель равен нулю
    If x = -1 Then
        MsgBox "Функция не определена при X = -1"
        Eq5 = 0
        Exit Function
    End If
    
    Select Case x
    Case Is < 0
        Eq5 = (1 + x + x ^ 2) / (1 + x ^ 3)
        
    Case 0 To 1
        Eq5 = Sqr(1 + (2 * x) / (1 + x ^ 2))

    Case Is >= 1
        Eq5 = 2 * Abs(0.5 + Sin(x))
    End Select
    
End Function

Function Eq6(x As Double) As Double
    If (x < -1.5 Or x > 1.8) Then
        MsgBox "X не входит в множество определения функции"
        Eq6 = 0
        Exit Function
    End If
    
    Select Case x
    Case Is < 0
        Eq6 = 1 + (3 + x) / (1 + x ^ 2)
        
    Case 0 To 1
        Eq6 = Sqr(1 + (1 - x) ^ 2)
        
    Case Is >= 1
        Eq6 = (1 + x) / (1 + Cos(x) ^ 2)
    End Select
    
End Function

Function Eq7(x As Double) As Double
    If (x < -1.4 Or x > 1.9) Then
        MsgBox "X не входит в множество определения функции"
        Eq7 = 0
        Exit Function
    End If
    
    Select Case x
    Case Is < 0
        Eq7 = (1 + 2 * x) / (1 + x ^ 2)
        
    Case 0 To 1
        Eq7 = Sin(x) ^ 2 * Sqr(1 + x)
        
    Case Is >= 1
        Eq7 = Sin(x) ^ 2 * Exp(0.2 * x)
    End Select
    
End Function

Function Eq8(x As Double) As Double
    If (x < 0 Or x > 1) Then
        MsgBox "X не входит в мн-во определения функции"
        Eq8 = 0
        Exit Function
    End If
    
    Select Case x
    Case Is < 0.2
        Eq8 = 1 + WorksheetFunction.Ln(1 + x)
    Case 0.2 To 0.8
        Eq8 = (1 + x ^ (1 / 2)) / (1 + x)
    Case Is > 0.8
        Eq8 = 2 * Exp(-2 * x)
    End Select
    
End Function



Sub EqSolve()

Dim x As Double, choice As Integer

    choice = InputBox("Выберите уравнение")
    x = InputBox("Введите число x =")

    Select Case choice
    Case 1
        MsgBox "Решение " & Eq1(x)
    Case 2
        MsgBox "Решение " & Eq2(x)
    Case 3
        MsgBox "Решение " & Eq3(x)
    Case 4
        MsgBox "Решение " & Eq4(x)
    Case 5
        MsgBox "Решение " & Eq5(x)
    Case 6
        MsgBox "Решение " & Eq6(x)
    Case 7
        MsgBox "Решение " & Eq7(x)
    Case 8
        MsgBox "Решение " & Eq8(x)
   
    Case Else
        MsgBox "ОШИБКА :((("

End Select

End Sub

