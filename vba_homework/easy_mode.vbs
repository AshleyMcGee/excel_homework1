Sub Button1_Click_easy_mode()

    'Declare Variables
    
    Dim totalVolume As Double
    Dim ticker As String
    Dim summary_table_row As Integer
    
    
    'For Each loop creates the variable ws inside the new object, Worksheets
    For Each ws In Worksheets
    

    'Add a header to Column I and Column J
    'ws. member access operator makes ranges and cells properties of the object Worksheets
    ws.Range("I1").Value = "<ticker>"
    ws.Range("J1").Value = "Total Vol"

    'Begin empty totalVolume variable and set summary table row
    totalVolume = 0
    summary_table_row = 2
    
    'Begin total volume loop
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Check to see if the cells have the same ticker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set the ticker symbol.
            ticker = ws.Cells(i, 1).Value

            'Add to the total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value

            'Print the ticker and volume to columns I and J respectively
            ws.Range("I" & summary_table_row).Value = ticker
            ws.Range("J" & summary_table_row).Value = totalVolume

            'Increment the summary table row
            summary_table_row = summary_table_row + 1

            'reset the totalVolume
            totalVolume = 0

        Else
            'Else if the current row value is the same as the next
            totalVolume = totalVolume + ws.Cells(i, 7).Value

        End If
        
    'Close total volume loop
    Next i
    
    'Close sheet loop
    Next ws
    
    MsgBox ("All Done!")

End Sub
