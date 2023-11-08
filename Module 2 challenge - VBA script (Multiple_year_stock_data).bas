Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

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

    start = 2 'Serve as pointer (location)
    stk_total = 0
    sum_tbl_row = 2 'Serve as pointer (location)
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row

' Create Summary Table
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"


'Loop through all table rows
    For i = 2 To lastrow
    
        'Assign pointer to ticker
        ticker = Cells(start, "A").Value
    
        'Conditional - If next row ticker symbol <> previous
        If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
                       
            'Assing value to variables
            
            ychange = Cells(i, 6).Value - Cells(start, 3)
            Perc_chg = ychange / Cells(start, 3)
            stk_total = stk_total + Cells(i, "G").Value
            
            
            'Add variables value to summary table
            Range("I" & sum_tbl_row).Value = ticker
            Range("J" & sum_tbl_row).Value = ychange
            Range("K" & sum_tbl_row).Value = Format(Perc_chg, "0.00%")
            Range("L" & sum_tbl_row).Value = Format(stk_total, "$0,00")
            
            
            If Cells(sum_tbl_row, "J").Value < 0 Then
            
                Cells(sum_tbl_row, "J").Interior.ColorIndex = 3
            
            Else
            
                Cells(sum_tbl_row, "J").Interior.ColorIndex = 4
            
            End If

            'Reset stk_total
            stk_total = 0
                        
               
            'Change position of sum_tbl_row and start to next i
            sum_tbl_row = sum_tbl_row + 1
            start = i + 1
        
            
        Else 'If next ticker = previous, then continue to add value from volume to stk_total

            stk_total = stk_total + Cells(i, "G").Value

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
    max = 0
    min = 0
    max_total = 0
    

    ' Loop through the column values
    For i = 2 To lastrow

        'If value ychange is greater than the current value assigned to max, then update value on max
        If Cells(i, "K").Value > max Then

            max = Cells(i, "K").Value
            Cells(2, "O").Value = Cells(i, "I").Value

        'If value ychange is lower than the current value assigned to min
        'then update value on min.
        ElseIf Cells(i, "K").Value < min Then

            min = Cells(i, "K").Value
            Cells(3, "O").Value = Cells(i, "I").Value

        End If
        
        'If value Total Stock Volume is greater than the current value assigned to max_total
        'then update value on max_total.
        If Cells(i, "L").Value > max_total Then
        
            max_total = Cells(i, "L").Value
            Cells(4, "O").Value = Cells(i, "I").Value
            
        End If
        
    Next i

' Create Summary Table
    Cells(1, "O").Value = "Ticker"
    Cells(1, "P").Value = "Value"
    Cells(2, "N").Value = "Gratest % Increase"
    Cells(3, "N").Value = "Gratest % Decrease"
    Cells(4, "N").Value = "Gratest Total Volume"
    Cells(2, "P").Value = Format(max, "0.00%")
    Cells(3, "P").Value = Format(min, "0.00%")
    Cells(4, "P").Value = Format(max_total, "$0.00")

Columns("N:P").AutoFit

'Next worksheet
Next ws

End Sub

