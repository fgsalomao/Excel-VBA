Attribute VB_Name = "Module1"
Sub Alfabetical_testing()

'Declare variables
Dim ws As Worksheet
Dim sort_range As Range
Dim lastrow As Long 'To find last row and make table iteration dynamic
Dim ticker As String 'For summary table
Dim ychange As Double 'For summary table Year change calculation
Dim Perc_chg As Double 'For summary table percentage change calculation
Dim stk_total As Double 'For summary table stock total calculation
Dim sum_tbl_row As Long 'Help move to next row on summary table
Dim start As Double 'Start ro
Dim i As Integer

'to iterate from each existing worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate

'Assign variables

    start = 2
    stk_total = 0
    sum_tbl_row = 2
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row

' Create Summary Table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"


'Loop through all table rows
    For i = 2 To lastrow
    
        'Assign value to ticker
        ticker = Cells(start, 1).Value
    
        'Conditional - If next row ticker symbol <> previous
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                       
            'Assing value to variables (Yearly Change, % Change,
            'Total Stock Volume)
            
            ychange = Cells(i, 6).Value - Cells(start, 3)
            Perc_chg = ychange / Cells(start, 3)
            stk_total = stk_total + Cells(i, 7).Value
            
            
            'Add variables value to summary table
            Range("I" & sum_tbl_row).Value = ticker
            Range("J" & sum_tbl_row).Value = ychange
            Range("K" & sum_tbl_row).Value = Format(Perc_chg, "0.00%")
            Range("L" & sum_tbl_row).Value = Format(stk_total, "$0,00")
            
            
            If Cells(sum_tbl_row, 10).Value < 0 Then
            
            Cells(sum_tbl_row, 10).Interior.ColorIndex = 3
            
            Else
            
            Cells(sum_tbl_row, 10).Interior.ColorIndex = 4
            
            End If

            'Reset stk_total
            stk_total = 0
                        
               
            'Change position of sum_tbl_row and start to next i
            sum_tbl_row = sum_tbl_row + 1
            start = i + 1
        
            
        Else 'If next ticker <> previous, then continue to add to volume to stk_total

            stk_total = stk_total + Cells(i, 7).Value

        End If

'Next table iteration
    Next i

'Adjust summary columns width
    Columns("I:L").AutoFit

'Variables for max, min and greater total
    Dim max As Double
    Dim min As Double
    Dim max_total As Double

'Assign value to variables
    lastrow = Cells(Rows.Count, "J").End(xlUp).Row 'Resets lastrow for last range
    max = Cells(2, 11).Value
    min = Cells(2, 11).Value
    max_total = Cells(2, 12).Value

    ' Loop through the column values
    For i = 2 To lastrow

        If Cells(i, 11).Value > max Then
            max = Cells(i, 11).Value
            Cells(2, 15).Value = Cells(i, 9).Value
            
        ElseIf Cells(i, 11).Value < min Then
            min = Cells(i, 11).Value
            Cells(3, 15).Value = Cells(i, 9).Value

        End If
        
        If Cells(i, 12).Value > max_total Then
        
            max_total = Cells(i, 12).Value
            Cells(4, 15).Value = Cells(i, 9).Value
            
        End If
        
    Next i

' Create Summary Table
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Gratest % Increase"
    Cells(3, 14).Value = "Gratest % Decrease"
    Cells(4, 14).Value = "Gratest Total Volume"
    Cells(2, 16).Value = Format(max, "0.00%")
    Cells(3, 16).Value = Format(min, "0.00%")
    Cells(4, 16).Value = Format(max_total, "$0.00")

Columns("N:P").AutoFit

'Next worksheet
Next ws
    

End Sub

