# Stock-Analysis
## Analysis on VBA Deliverable 1
  The purpose of this analysis is to compare the effectiveness of using the orginal VBA script in contrast with the refractored VBA script.
  Speed, the code used, and the process to produce the code will be examined. 
### Speed 
--
The refractored script produced the following results, 2017 results were delivered in .140625 seconds.

  --
  <img width="225" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/90878901/136670336-599c876c-157c-4381-890a-2773613f8a70.png">

  --
  The refractored script for 2018 delivered the results in .1445313 seconds. 
  
  --
<img width="169" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/90878901/136670354-e3cd4085-175c-4352-bc79-12ce044ba51a.png">

  --
  This was an increase from the class work that resulted in 1.640625 for 2017 and 1.644531 for 2018.
  
--
### Coding
--
-The orginal script coding was easier on multiple issues and data to reflect how to build the code was methodical. The course work did a good job of explaining the process and how to develop the final results. This sequential step by step process was possible to follow.

-The refractoring challenge reinforced the concepts developed during the course work. This was an extremely difficult challenge for me and took considerable hours including reaching out to TA's, class mates, office hours, online research and reaching out to Ask BCS. The final error was due to the missplaced setting of the nested if in section3(d). Mo did an excellent job of helping me through this last issue. 
The final code can be viewed below:
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate  
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    'Initialize array of all tickers
    Dim tickers(12) As String  
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"  
    'Activate data worksheet
    Worksheets(yearValue).Activate 
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0  
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i      
    ''2b) Loop over all the rows in the spreadsheet.
  
    For i = 2 To RowCount
        '3a) Increase volume for current ticker
       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
      
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If Then Statement
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrice(tickerIndex) = Cells(i, 6).Value             
        'End If
        End If      
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
           
        'End If           
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1     
        End If 
    Next i   
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11   
        Worksheets("All Stocks Analysis").Activate        
    Next i    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then   
            Cells(i, 3).Interior.Color = vbGreen      
        Else
            Cells(i, 3).Interior.Color = vbRed
        End If
    Next i
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
End Sub

 ### Code Development Process
 --
 Of course the work on canvas helped develop the orginal code. The refractoring was a very difficult challenge for me and took considerable hours including reaching out to TA's, class mates, office hours, online research (https://stackoverflow.com/, https://docs.microsoft.com/en-us/office/vba/Language/Concepts/Getting-Started/understanding-visual-basic-syntax and https://excelchamps.com/vba/subscript-out-of-range-error-9/ as online examples. I also reached out out to Ask BCS. The final error was due to the missplaced ending of an nested if. 
### Conclusion
--
While the results of refractoring were considerably quicker approximately 1 second for each years analysis. The amount of effort that was created by redoing the code utimately lead me to consider in this case that it was not necessary to change to refractoring. If the file was considerably larger and the lessons drawn from the troubleshooting were applied it might be advantageous to use this method on future larger projects. 
     
