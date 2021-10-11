# Stock-Analysis
## Analysis on VBA Deliverable 1
  The purpose of this analysis is to compare the effectiveness of using the orginal VBA script in contrast with the refractored VBA script.
  Speed, the code used, and the process to produce the code will be examined. 
### Speed 
--
The refractored script produced the following results, 2017 results were delivered in .863 seconds.

  --
 
<img width="355" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/90878901/136820808-e9dfc954-cb90-422e-bfe0-f82cf880a117.png">

  --
  The refractored script for 2018 delivered the results in .125 seconds. 
  
  --

<img width="335" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/90878901/136820838-c48eee6f-49c7-4729-98c1-5b92a49772f8.png">

  --
  This was an increase from the class work that resulted in 1.640625 for 2017 and 1.644531 for 2018.
  
--
### Coding
--
-The orginal script coding was easier on multiple issues and data to reflect how to build the code was methodical. The course work did a good job of explaining the process and how to develop the final results. This sequential step by step process was possible to follow.

-The refractoring challenge reinforced the concepts developed during the course work. This was an extremely difficult challenge for me and took considerable hours including reaching out to TA's, class mates, office hours, online research and reaching out to Ask BCS. One of the key errors were due to the missplaced setting of the nested if in section3(d). Mo did an excellent job of helping me through this issue. Additionally on the 2nd assignment submission, an entire section of code was missing making the code inoperable. This issue is now resolved. 
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
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrice(11) As Single
    Dim tickerEndingPrice(11) As Single
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
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1 
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

 ### Code Development Process and Research
 --
 Of course the work on canvas helped develop the orginal code. The refractoring was a very difficult challenge for me and took considerable hours including reaching out to TA's(thank you to Nick and Mo), class mates, office hours, online research (https://stackoverflow.com/, https://docs.microsoft.com/en-us/office/vba/Language/Concepts/Getting-Started/understanding-visual-basic-syntax and https://excelchamps.com/vba/subscript-out-of-range-error-9/ as online examples. I also reached out out to Ask BCS. utilizing these tools gained additional understanding of the code and debugging the errors as they arose. 
  I also performed an independent review of refractoring. As personal experiance to this point was challenged. I reviewed the following thesis: https://scholarworks.rit.edu/cgi/viewcontent.cgi?article=11597&context=theses on __Conceptions of Refractoring.__ I also reviewed refractoring on stacked overflow. From https://stackoverflow.com/questions/20624340/what-is-refactoring __Advantages include improved code readability and reduced complexity to improve the maintainability of the source code, as well as a more expressive internal architecture or object model to improve extensibility.__ Also this quote in the same thread. __“Refactoring is the process of changing a software system in such a way that it does not alter the external behavior of the code, yet improves its internal structure. It is a disciplined way to clean up code that minimizes the chances of introducing bugs. In essence when you refactor you are improving the design of the code after it has been written.” - Martin Fowler (Father of Code Smell)__ Martin Fowler is often quoted in many of the sites I reviewed on the concept of refractoring. 
  
### Conclusion
--From https://stackoverflow.com/questions/20624340/what-is-refactoring __ https://stackoverflow.com/users/3703904/masud-shrabon __
__The purposes of refactoring according to M. Fowler are stated in the following:

__Refactoring Improves the Design of Software
Refactoring Makes Software Easier to Understand
Refactoring Helps Finding Bugs
Refactoring Helps Programming Faster__
The refractoring process was challenging for me. Given the additional resources reviewed. The increase in file size, and having debugged the file multiple times. There are advantages of refractoring beyond a speed aspect. My initial review was limited on the speed concept with the amount of work that was put into the file development. 
