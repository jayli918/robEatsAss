Attribute VB_Name = "Module2"
Sub riskRules()
    
    
    
    Dim DhighestValue As Double, DsecondHighestValue As Double, AhighestValue As Double, AsecondHighestValue As Double
    
    For i = 14 To 7789
        
        DhighestValue = 0
        
        DsecondHighestValue = 0
        
        AhighestValue = 0
        
        AsecondHighestValue = 0
    
            For c = 1 To 2
                
                If Cells(i, c).Value > DhighestValue Then
                    
                    DhighestValue = Cells(i, c).Value
                    
                ElseIf Cells(i, c).Value = DhighestValue Then
                    
                    DhighestValue = Cells(i, c).Value
                
                End If
                
                If c = 1 Then
                
                    DsecondHighestValue = Cells(i, 2).Value
                
                ElseIf c = 2 Then
                    
                    DsecondHighestValue = Cells(i, 1).Value
                
                End If
        
            Next c
            
            For K = 3 To 5
            
                If Cells(i, K).Value > AhighestValue Or Cells(i, K).Value = AhighestValue Then
                    
                    AhighestValue = Cells(i, K).Value
                
                End If
                
                If Cells(i, K).Value > AsecondHighestValue And Cells(i, K).Value < AhighestValue Then
                    
                    AsecondHighestValue = Cells(i, K).Value
                
                End If
                
            Next K
            
            If DhighestValue > AhighestValue Or DhighestValue = AhighestValue Then
            
                Cells(i, 7) = 1
                
            ElseIf DhighestValue < AhighestValue Then
                
                Cells(i, 9) = 1
                
            End If
            
            If DsecondHighestValue > AsecondHighestValue Or DsecondHighestValue = AsecondHighestValue Then
            
                Cells(i, 8) = 1
            
            ElseIf DsecondHighestValue < AsecondHighestValue Then
            
                Cells(i, 10) = 1
                
            End If
        
    Next i
    

End Sub

Sub clearStuff()

    For i = 14 To 7789
    
        For c = 7 To 10
        
            Cells(i, c) = ""
            
        Next c
        
    Next i
    

End Sub
