Sub Button2_Click_feature_test()
'Unfinished: Format cells in Yearly Change

'Declare Variables
Dim totalVolume As Double
Dim ticker As String
Dim summary_table_row As Integer
Dim opening As Double
Dim closing As Double
Dim yearlyChange As Double
Dim percentChange As Double

'Begin Worksheet Loop
For Each ws In Worksheets

'Set column headers
ws.Range("I1").Value = "<ticker>"
ws.Range("J1").Value = "Total Vol"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"

'Begin with empty variables and set opening to initial value
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
totalVolume = 0
yearlyChange = 0
percentChange = 0
summary_table_row = 2
opening = ws.Cells(2, 3).Value
    
    'Begin iterating through rows
    For i = 2 To lastRow

        'Begin searching for the new ticker value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set the starting values
            ticker = ws.Cells(i, 1).Value
            opening = ws.Cells(i + 1, 3).Value
            closing = ws.Cells(i, 6).Value
            
            'Make sure nothing is dividing by zero. Also handles Overflow
            If opening = 0 Then

                opening = 0.001

            End If
            
            'Perform the math
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            yearlyChange = opening - closing
            percentChange = (closing - opening) / opening
    
            'Print the yearly change and percent change to columns I,J,K,L respectively
            ws.Range("I" & summary_table_row).Value = ticker
            ws.Range("J" & summary_table_row).Value = totalVolume
            ws.Range("K" & summary_table_row).Value = yearlyChange
            ws.Range("L" & summary_table_row).Value = percentChange

            'Format cells
    
            'increment summary_table_row
            summary_table_row = summary_table_row + 1
    
            'reset the variables so the loop can run again
            totalVolume = 0
            yearlyChange = 0
            percentChange = 0
            opening = 0
    
            Else
                'Else if the current row value is the same as the next
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
        'End Different Ticker Condition
        End If
    
    'Next Row
    Next i

    'Format Cells for loop
    For i = 2 to lastRow

        If ws.Cells(i,11).Value < 0 Then

            ws.Cells(i,11).Interior.ColorIndex = 3

            Else 

            ws.Cells(i,11).Interior.ColorIndex = 4

        End if 

        summary_table_row = summary_table_row + 1

    Next i
    
'Next worksheet
Next ws

'Ding
MsgBox ("I'm Done!")


End Sub


