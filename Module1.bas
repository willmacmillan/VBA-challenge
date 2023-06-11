Attribute VB_Name = "Module1"
Sub Module2Challenge():

    For Each Worksheet In Worksheets
    
        'Create variables
        
        Dim WorksheetName As String
        
        Dim i As Long
        
        Dim j As Long
        
        Dim LastRowA As Long
        
        Dim LastRowI As Long
        
        Dim Counter As Long
    
        Dim PercentChange As Double
     
        Dim GreatestIncrease As Double
        
        Dim GreatestDecrease As Double

        Dim GreatestVolume As Double

        WorksheetName = Worksheet.Name
        
        'Create the labels for table
        Worksheet.Cells(1, 9).Value = "Ticker"
        Worksheet.Cells(1, 10).Value = "Yearly Change"
        Worksheet.Cells(1, 11).Value = "Percent Change"
        Worksheet.Cells(1, 12).Value = "Total Stock Volume"
        Worksheet.Cells(1, 16).Value = "Ticker"
        Worksheet.Cells(1, 17).Value = "Value"
        Worksheet.Cells(2, 14).Value = "Greatest % Increase"
        Worksheet.Cells(3, 14).Value = "Greatest % Decrease"
        Worksheet.Cells(4, 14).Value = "Greatest Total Volume"
        
        ' Code for finding last row in column A
        LastRowA = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Start the counter in row 2
        Counter = 2
        
        ' Start the start row to 2
        j = 2
        
        'Loop through all rows starting at 2
        For i = 2 To LastRowA
            
            'Check if the value of the ticker is different
            If Worksheet.Cells(i + 1, 1).Value <> Worksheet.Cells(i, 1).Value Then
                
                'Put ticker name in appropriate cell
                Worksheet.Cells(Counter, 9).Value = Worksheet.Cells(i, 1).Value
                
                'Calculation for Yearly Change
                Worksheet.Cells(Counter, 10).Value = (Worksheet.Cells(i, 6).Value - Worksheet.Cells(j, 3).Value)
                
                'If statement for coloring Yearly Change values red if negative
                'And green if positive
                If Worksheet.Cells(Counter, 10).Value < 0 Then
                
                    Worksheet.Cells(Counter, 10).Interior.ColorIndex = 3
                
                Else
                
                    Worksheet.Cells(Counter, 10).Interior.ColorIndex = 4
                
                End If
                    
                'Calculation for Percent Change
                If Worksheet.Cells(j, 3).Value <> 0 Then
                
                    PercentChange = ((Worksheet.Cells(i, 6).Value - Worksheet.Cells(j, 3).Value) / Worksheet.Cells(j, 3).Value)
                   
                    'Format into a percent
                     Worksheet.Cells(Counter, 11).Value = Format(PercentChange, "Percent")
                Else
                        
                    Worksheet.Cells(Counter, 11).Value = Format(0, "Percent")
                        
                    
                End If
                    
                'Calculation for Total Volume
                Worksheet.Cells(Counter, 12).Value = WorksheetFunction.Sum(Range(Worksheet.Cells(j, 7), Worksheet.Cells(i, 7)))
                
                'Increase Counter by 1
                Counter = Counter + 1
                
                'Set new start row
                j = i + 1
                
                End If
            
            Next i
            
             
        'Code for finding the last row in Column i
        LastRowI = Worksheet.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Assign values for Greatest increase, decrease, and volume
        GreatestIncrease = Worksheet.Cells(2, 11).Value
        GreatestDecrease = Worksheet.Cells(2, 11).Value
        GreatestVolume = Worksheet.Cells(2, 12).Value
        
            'Loop for summary
            For i = 2 To LastRowI
            
                'Find the greatest volume by looping through each one and replacing the current value with any larger value
                If Worksheet.Cells(i, 12).Value > GreatestVolume Then
                    GreatestVolume = Worksheet.Cells(i, 12).Value
                    Worksheet.Cells(4, 16).Value = Worksheet.Cells(i, 9).Value
                
                Else
                
                    GreatestVolume = GreatestVolume
                
                End If
                
                'Find the greatest increase by looping through each one and replacing the current value with any larger value
                If Worksheet.Cells(i, 11).Value > GreatestIncrease Then
                    GreatestIncrease = Worksheet.Cells(i, 11).Value
                    Worksheet.Cells(2, 16).Value = Worksheet.Cells(i, 9).Value
                
                Else
                
                    GreatestIncrease = GreatestIncrease
                
                End If
                
                'Find the greatest decrease by looping through each one and replacing the current value with any larger value
                If Worksheet.Cells(i, 11).Value < GreatestDecrease Then
                    GreatestDecrease = Worksheet.Cells(i, 11).Value
                    Worksheet.Cells(3, 16).Value = Worksheet.Cells(i, 9).Value
                
                Else
                
                    GreatestDecrease = GreatestDecrease
                
                End If
                
            'Write Greatest increase, decrease, and volume values into corresponding cells
            Worksheet.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            Worksheet.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            Worksheet.Cells(4, 17).Value = GreatestVolume
            
            Next i
            
            'Programatically change cell width
            Worksheets(WorksheetName).Columns("A:Z").AutoFit
    
    Next Worksheet
                    
            
End Sub
