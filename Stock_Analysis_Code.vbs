' ##Bonus - Part 2
Sub Run_code_all_worksheets()

' Making the code runs on every worksheet
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        ' Setting zoom to 150% on every worksheet (better vizualization on my screen)
        xSh.Activate
        ActiveWindow.Zoom = 150
        ' Running the code on every worksheet
        xSh.Select
        Call Stock_Analysis
    Next
    Application.ScreenUpdating = True
End Sub

' ## Stock Market Analysis Code
Sub Stock_Analysis()

' Adding header to the new columns
Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "yearly_change"
Cells(1, 11).Value = "percent_change"
Cells(1, 12).Value = "total_stock_volume"

' Setting the Cell Format for the new columns
Range("J:J").NumberFormat = "0.00"
Range("K:K").NumberFormat = "0.00%"
      
' Setting the variables
Dim ticker As String
Dim openning As Double
Dim closing As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim summary As Integer

' Setting the integer value
summary = 2

' Last Row for the end of the loop
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all the stocks for one year
For i = 2 To Lastrow
        
        ' Conditional to find the final data of one stock and output the data desired on the new column
        If Cells(i + 1, 1) <> Cells(i, 1) Then
        
        ' Finding the ticker symbols and outputting on the correspondent cell
        ticker = Cells(i, 1).Value
        Cells(summary, 9).Value = ticker
        ticker = 0
        
        ' Finding the last closing cost of the year for each stock, calculating the yearly closing and outputting on the correspondent cell
        closing = Cells(i, 6).Value
        yearly_change = closing - openning
        Cells(summary, 10).Value = yearly_change
        yearly_closing = 0
         ' Conditional formatting to highligh positive change in green and negative change in red
        Set condition1 = Cells(summary, 10).FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
        Set condition2 = Cells(summary, 10).FormatConditions.Add(xlCellValue, xlLess, "=0")
            With condition1
                .Interior.ColorIndex = 10
            With condition2
                .Interior.ColorIndex = 53
            End With
            End With
        
        ' Calculating the percent change and outputting on the correspondent cell
        percent_change = (closing - openning) / openning
        Cells(summary, 11).Value = percent_change
        percent_change = 0
          ' Conditional formatting to highligh positive change in green and negative change in red
        Set condition1 = Cells(summary, 11).FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
        Set condition2 = Cells(summary, 11).FormatConditions.Add(xlCellValue, xlLess, "=0")
            With condition1
                .Interior.ColorIndex = 10
            With condition2
                .Interior.ColorIndex = 53
            End With
            End With

        ' Final calculating the total stock volume and outputting on the correspondent cell
        total_stock_volume = total_stock_volume + Cells(i, 7)
        Cells(summary, 12).Value = total_stock_volume
        total_stock_volume = 0
        
        ' Adding 1 to the integer to start the next outputs on the next row
        summary = summary + 1
        
        Else
        
        ' Calculating the partial total stock volume
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
                  
            ' Finding the first opening cost of the year for each stock
            If Right(Cells(i, 2).Value, 4) = "0102" Then
            openning = Cells(i, 3).Value
                                         
            End If
                
        End If

Next i
  
' Ajusting the size of the new columns
Columns("J:L").AutoFit
  
' ##Bonus - Part 1
' Add functionality to the script to return the stock with te "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
  
' Creating the new cells
Cells(2, 15).Value = "Greatest % increase"
Cells(3, 15).Value = "Greatest % decrease"
Cells(4, 15).Value = "Greatest total volume"
Cells(1, 16).Value = "ticker"
Cells(1, 17).Value = "value"

' Setting the variables
Dim Max_Perc As Double
Dim Min_Perc As Double
Dim Max_Total_Volume As Double

' Finding the "Gratest % increase" and filling the correspondents cells with the informations
Max_Perc = Application.WorksheetFunction.Max(Range("K:K"))
Cells(2, 17).Value = Max_Perc
Range("P2").Value = Application.WorksheetFunction.XLookup(Range("Q2"), Range("K:K"), Range("I:I"), False)

' Finding the "Greatest % decrease" and filling the correspondents cells with the informations
Min_Perc = Application.WorksheetFunction.Min(Range("K:K"))
Cells(3, 17).Value = Min_Perc
Range("P3").Value = Application.WorksheetFunction.XLookup(Range("Q3"), Range("K:K"), Range("I:I"), False)

' Finding the "Greatest total volume" and filling the correspondents cells with the informations
Max_Total_Volume = Application.WorksheetFunction.Max(Range("L:L"))
Cells(4, 17).Value = Max_Total_Volume
Range("P4").Value = Application.WorksheetFunction.XLookup(Range("Q4"), Range("L:L"), Range("I:I"), False)

' Ajusting the size of the new columns
Columns("O:Q").AutoFit

' Setting the Cell Format for the new cells
Range("Q2:Q3").NumberFormat = "0.00%"
Cells(4, 17).NumberFormat = "0"
  
End Sub