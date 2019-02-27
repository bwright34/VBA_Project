Sub Multipleyearstockdata()

        'Declare the variables
        Dim ticker As String
        Dim vol As Double
        Dim i As Double
        Dim LastRow As Double
        Dim sheet1 As Worksheet
        Dim summarytable As Double
        Dim yearlychange As Double
        Dim totalstockvolume As Double
        Dim percentchange As Double
        Dim yearclose As Double
        Dim yearopen As Double
        Dim row As Double
        
        'Assign the variables
        summarytable = 2
        vol = 0
        totalstockvolume = 0
        Application.ScreenUpdating = False
        
        'Loop through all worksheets
        For Each sheet1 In ThisWorkbook.Worksheets
             sheet1.Activate
        
              'Insert headlines on excel on each sheet
              sheet1.Cells(1, 9).Value = "ticker"
              sheet1.Cells(1, 10).Value = "yearly change"
              sheet1.Cells(1, 11).Value = "percent change"
              sheet1.Cells(1, 12).Value = "total stock volume"
        
              'Loop through tickers
              LastRow = Cells(Rows.Count, 1).End(xlUp).row
              ticker = sheet1.Cells(2, 1).Value
              yearopen = sheet1.Cells(2, 3).Value
              For i = 2 To LastRow
        
                  'Check if we are still within the same stock ticker, do an if then conditional statement
                    If ticker <> sheet1.Cells(i, 1).Value Then
        
                      'Print the ticker name into the summary table
                      sheet1.Range("I" & summarytable).Value = ticker
                      ticker = sheet1.Cells(i, 1).Value
        
                     'Print the yearly change into the summary table
                      yearclose = sheet1.Cells(i - 1, 6).Value
                      yearlychange = yearclose - yearopen
                      sheet1.Range("J" & summarytable).Value = yearlychange
        
                     'Print the percent change into the summary table
                     If yearopen = 0 Then
                     
                          For row = i To LastRow
                          
                              If sheet1.Cells(i, 3).Value > 0 Or sheet1.Cells(i, 3).Value < 0 Then
                                  yearopen = sheet1.Cells(i, 3).Value
                                  
                                  Exit For
                                  
                              End If
                              
                          Next row
                          
                     End If
                     
                      percentchange = (yearclose - yearopen) / yearopen
                      sheet1.Range("K" & summarytable).Value = percentchange
        
                     'Set year open
                     yearopen = sheet1.Cells(i, 3).Value
        
                     'Change format to percent
                      sheet1.Columns("K").NumberFormat = "0.00%"
        
                     'Print the total stock volume into the summary table
                      sheet1.Range("L" & summarytable).Value = vol
        
                      'Add one to the summary table
                      summarytable = summarytable + 1
        
                      'Reset the total stock volume amount
                      vol = 0
        
        
                  End If
        
                  vol = vol + sheet1.Cells(i, 7).Value
        
                  'Formatting color
        
                  If sheet1.Cells(i, 10).Value > 0 Then
                       sheet1.Cells(i, 10).Interior.ColorIndex = 4
                  Else
                      sheet1.Cells(i, 10).Interior.ColorIndex = 3
                  End If
        
              Next i
        
        'reset summarytable to 2 before moving onto next sheet
        summarytable = 2
        
        'move to next worksheet
        Next sheet1
        
        
        
        Application.ScreenUpdating = True
        

End Sub

