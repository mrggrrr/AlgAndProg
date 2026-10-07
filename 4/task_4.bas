Attribute VB_Name = "task_4"
Option Explicit

Sub Task_4()
    Dim n As Integer
    Dim i As Integer
    Dim Max As Integer
    Dim Min As Integer
    Dim r As Integer
    Dim S As Integer
    Dim Row As String
    Dim Mean As Double
    
        n = Int(Rnd * 90) + 10
        i = 1
        S = 0
        
            Do
            
                r = Int(Rnd * 90) + 10
                
                S = S + r
                
                Row = Row & S & " "
                
                    If i = 1 Then
                        Max = S
                        Min = S
                    Else
                        If S > Max Then
                            Max = S
                        End If
                        
                        If S < Min Then
                            Min = S
                        End If
                    End If
                    
                i = i + 1
                
            Loop Until i > n
            
        MsgBox "Накопленный ряд: " & Chr(10) & Row & Chr(10) & "Максимум = " & Max & Chr(10) & "Минимум = " & Min
            
            
End Sub

