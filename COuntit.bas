Attribute VB_Name = "Module3"
Sub countPerm()

Dim countD2 As Long, countD1 As Long, countD As Long

For i = 14 To 7789
        
        countD2 = Cells(6, 8).Value
        
        countD1 = Cells(6, 9).Value
        
        countD = Cells(6, 10).Value
        
        If Cells(i, 7).Value = 1 And Cells(i, 8).Value = 1 Then
            
            countd2fin = countD2 + 1
            
            Cells(6, 8).Value = countd2fin
            
        ElseIf Cells(i, 7).Value = "" And Cells(i, 8) = 1 Then
        
            countD1fin = countD1 + 1
            
            Cells(6, 9).Value = countD1fin
            
        ElseIf Cells(i, 7).Value = 1 And Cells(i, 8) = "" Then
        
            countD1fin = countD1 + 1
            
            Cells(6, 9).Value = countD1fin
            
        ElseIf Cells(i, 7).Value = "" And Cells(i, 8).Value = "" Then
        
            countDfin = countD + 1
            
            Cells(6, 10).Value = countDfin
        
        End If
            
Next i

End Sub
