Sub Button_Click1_EasyMode():

    'Declare Variables
    
    Dim totalVolume As Double
    Dim ticker As String
    Dim summary_table_row as Integer

    'Add a header to Column I and Column J
    Range("I1").Value = "<ticker>"
    Range("J1").Value = "Total Vol"

    'Begin empty totalVolume variable and set summary table row
    totalVolume = 0
    summary_table_row = 2

    For i = 2 to Cells(Rows.Count,1).End(xlUp).Row

    'Check to see if the cells have the same ticker symbol
        If Cells(i + 1, 1).Value <> Cells(i,1).Value Then

            'Set the ticker symbol.
            ticker = Cells(i,1.).Value

            'Add to the total volume
            totalVolume = totalVolume + Cells(i,7).Value 

            'Print the ticker and volume to columns I and J respectively
            Range("I" & summary_table_row).Value = ticker
            Range("J" & summary_table_row).Value = totalVolume

            'Increment the summary table row
            summary_table_row = summary_table_row + 1

            'reset the totalVolume
            totalVolume = 0

        Else 
            'Else if the current row is the same as the next
            totalVolume = totalVolume + Cells(i,7).Value

        End If

    Next i


End Sub