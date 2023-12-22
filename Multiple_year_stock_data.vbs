' Module 2 Challenge
' Mitchell Fairgrieve
' Due 12/21/2023

Sub MultipleYearStockData():

    ' Begin process to enable VBA script to run on every worksheet (that is, every year) at once
    For Each Sheet In Worksheets
    
        Dim WsName As String
        
        ' Obtain WorksheetName for Worksheets Loop
        WsName = Sheet.Name
        
        
        ' Row Variable
        Dim Row As Long
        
        ' Ticker Starting Row
        Dim TickerRow As Long
        
        ' Ticker Row Counter
        Dim TickerCounter As Long
        
        ' Percent Change Variable
        Dim PercentChange As Double
        
        ' Greatest Increase Variable
        Dim GreatestIncrease As Double
        
        ' Greatest Decrease Variable
        Dim GreatestDecrease As Double
        
        ' Greatest Total Volume Variable
        Dim GreatestVolume As Double
        
        ' Last row in column 1
        Dim FinalColumn1 As Long
        
        ' Last row in column 9
        Dim FinalColumn9 As Long
        
        
        ' Input the Column Headers into Sheet Cells
        Sheet.Cells(1, 9).Value = "Ticker"
        Sheet.Cells(1, 10).Value = "Yearly Change"
        Sheet.Cells(1, 11).Value = "Percent Change"
        Sheet.Cells(1, 12).Value = "Total Stock Volume"
        Sheet.Cells(1, 16).Value = "Ticker"
        Sheet.Cells(1, 17).Value = "Value"
        Sheet.Cells(2, 15).Value = "Greatest % Increase"
        Sheet.Cells(3, 15).Value = "Greatest % Decrease"
        Sheet.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        ' Set the Ticker Counter to start on the First Row
        TickerCounter = 2
        
        ' Set the starting row to 2 (same as Ticker)
        TickerRow = 2
        
        
        ' Find the last POPULATED cell in column #1
        FinalColumn1 = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        
            ' Begin For Loop through all Rows
            For Row = 2 To FinalColumn1
            
                If Sheet.Cells(Row + 1, 1).Value <> Sheet.Cells(Row, 1).Value Then
                
                ' Input Ticker Counter Value in column #9
                Sheet.Cells(TickerCounter, 9).Value = Sheet.Cells(Row, 1).Value
                
                ' Finalize and Input Yearly Change in column #10
                Sheet.Cells(TickerCounter, 10).Value = Sheet.Cells(Row, 6).Value - Sheet.Cells(TickerRow, 3).Value
                
                    ' Begin Coloring
                    If Sheet.Cells(TickerCounter, 10).Value < 0 Then
                
                    ' Background = Red
                    Sheet.Cells(TickerCounter, 10).Interior.ColorIndex = 3
                
                    Else
                
                    ' Background = Green
                    Sheet.Cells(TickerCounter, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    ' Finalize and Input Percent Change into column #11
                    If Sheet.Cells(TickerRow, 3).Value <> 0 Then
                    PercentChange = ((Sheet.Cells(Row, 6).Value - Sheet.Cells(TickerRow, 3).Value) / Sheet.Cells(TickerRow, 3).Value)
                    
                    Sheet.Cells(TickerCounter, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    Sheet.Cells(TickerCounter, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                ' Finalize and Input Total Volume into column #12
                Sheet.Cells(TickerCounter, 12).Value = WorksheetFunction.Sum(Range(Sheet.Cells(TickerRow, 7), Sheet.Cells(Row, 7)))
                
                ' Add 1 to Ticker Counter
                TickerCounter = TickerCounter + 1
                
                ' Initialize new Row for the Ticker Row
                TickerRow = Row + 1
                
                End If
            
            Next Row
            
            
        ' Find last POPULATED cell in column #9
        FinalColumn9 = Sheet.Cells(Rows.Count, 9).End(xlUp).Row


        
        ' Begin Summary
        GreatestVolume = Sheet.Cells(2, 12).Value
        GreatestIncrease = Sheet.Cells(2, 11).Value
        GreatestDecrease = Sheet.Cells(2, 11).Value
        
            For Row = 2 To FinalColumn9
            
                ' Calculate the Greatest Total Volume
                If Sheet.Cells(Row, 12).Value > GreatestVolume Then
                GreatestVolume = Sheet.Cells(Row, 12).Value
                Sheet.Cells(4, 16).Value = Sheet.Cells(Row, 9).Value
                
                Else
                
                GreatestVolume = GreatestVolume
                
                End If
                
                ' Calculate the Greatest Increase
                If Sheet.Cells(Row, 11).Value > GreatestIncrease Then
                GreatestIncrease = Sheet.Cells(Row, 11).Value
                Sheet.Cells(2, 16).Value = Sheet.Cells(Row, 9).Value
                
                Else
                
                GreatestIncrease = GreatestIncrease
                
                End If
                
                ' Calculate the Greatest Decrease
                If Sheet.Cells(Row, 11).Value < GreatestDecrease Then
                GreatestDecrease = Sheet.Cells(Row, 11).Value
                Sheet.Cells(3, 16).Value = Sheet.Cells(Row, 9).Value
                
                Else
                
                GreatestDecrease = GreatestDecrease
                
                End If
                
            ' Finalize and Input the summary results into Sheet cells
            Sheet.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            Sheet.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            Sheet.Cells(4, 17).Value = Format(GreatestVolume, "Scientific")
            
            Next Row
            
    Next Sheet
        
End Sub


' Coding References:
'
'
' 1.) Site used to reference for function that enables a VBA to run on every worksheet (that is, every year) at once
' --> https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
'
'
' 2.) Site used to reference for function that finds last row in any given column
' --> https://www.thespreadsheetguru.com/last-row-column-vba/
'
' 3.) Site used to reference data types and determine difference between long and double
' --> https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary
'
'

' end VBA file


