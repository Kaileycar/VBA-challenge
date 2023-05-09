Sub VBA():

'set variables
    Dim Worksheet As String
    Dim WS_Count As Integer
    Dim ticker As String
    Dim volume As LongLong
    Dim LastRow As Long
    
    Dim prev_ticker As String
    Dim Display_Row As Long
    Dim Open_Price As Double
    Dim Close_Price As Double

    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As LongLong
    
    Dim i As Integer
    Dim j As Long

    
    
'determine # of worksheets
    WS_Count = ThisWorkbook.Worksheets.Count
'set worksheet
    For i = 1 To WS_Count
    
    
'creat variable for ws
        Dim WorksheetName As String
    
    
'Determine last row : Sheets(i) circulates through all worksheets
        LastRow = Sheets(i).Cells(Rows.Count, 1).End(xlUp).Row
        
       ' MsgBox (LastRow)
    
'grab ws name - (i) equals which ws we are one from the "i" variable we set
        WorksheetName = Sheets(i).Name
        
        
        'MsgBox (WorksheetName)
        
                
                
        'Add columns for ticker, yearly change, percentage change, and volume
        'Using column I because everything will be pushed over to the left
        Sheets(i).Range("I1").EntireColumn.Insert
        Sheets(i).Range("I1").EntireColumn.Insert
        Sheets(i).Range("I1").EntireColumn.Insert
        Sheets(i).Range("I1").EntireColumn.Insert

        'Add columns for % increase/decrease/total volume, ticker, and value
        Sheets(i).Range("N1").EntireColumn.Insert
        Sheets(i).Range("N1").EntireColumn.Insert
        Sheets(i).Range("N1").EntireColumn.Insert
        
        
        'Add name to collumns
       Sheets(i).Cells(1, 9).Value = "Ticker"
       Sheets(i).Cells(1, 10).Value = "Yearly Change"
       Sheets(i).Cells(1, 11).Value = "Percent Change"
       Sheets(i).Cells(1, 12).Value = "Total Stock Volume"

       Sheets(i).Cells(1, 15).Value = "Ticker"
       Sheets(i).Cells(1, 16).Value = "Value"
        
        'Add name to rows
       Sheets(i).Cells(2, 14).Value = "Greatest % Increase"
       Sheets(i).Cells(3, 14).Value = "Greatest % Decrease"
       Sheets(i).Cells(4, 14).Value = "Greatest Total Volume"
       
     


       'initialize previous ticker & DR
       prev_ticker = " "
       Display_Row = 1
       
'start second loop
             For j = 2 To LastRow
            
    
'Retrieve and store data in each variable
        ticker = Sheets(i).Cells(j, 1).Value
        
        If ticker = prev_ticker Then
        
            'MsgBox ("same ticker")
            'MsgBox (ticker)
            'MsgBox (volume)
            
            Close_Price = Sheets(i).Cells(j, 6).Value
            volume = volume + Sheets(i).Cells(j, 7).Value
        
        Else
            'MsgBox ("new ticker")
            'MsgBox (prev_ticker)
            'MsgBox (ticker)
            
            
        'if statement for summary row. Add end if after ticker and dr
        'writing out previous ticker only if greater than row2. This is what we are going to do always
             If j > 2 Then
                Yearly_Change = (Close_Price - Open_Price)
                Percent_Change = (Yearly_Change / Open_Price)
                'Percent_Change = Application.WorksheetFunction.RoundUp(Percent_Change, 2)


        'Put output in Display Row
                Sheets(i).Cells(Display_Row, 9).Value = prev_ticker
                Sheets(i).Cells(Display_Row, 10).Value = Yearly_Change
                Sheets(i).Cells(Display_Row, 11).Value = Percent_Change
                Sheets(i).Cells(Display_Row, 12).Value = volume
                
                'format to color change
                 If Yearly_Change < 0 Then
                    Sheets(i).Cells(Display_Row, 10).Interior.ColorIndex = 3

                Else
                    Sheets(i).Cells(Display_Row, 10).Interior.ColorIndex = 4

                End If
                
                'format to percent sign
                Sheets(i).Cells(Display_Row, 11).NumberFormat = "0.00%"
                Percent_Change = Application.WorksheetFunction.RoundUp(Percent_Change, 2)


             End If

            'fill in previous ticker with new ticker
            prev_ticker = ticker
            'Keep going down 1 Display Row to fill in data
            Display_Row = Display_Row + 1
            Open_Price = Sheets(i).Cells(j, 3).Value
            Close_Price = Sheets(i).Cells(j, 6).Value
            volume = Sheets(i).Cells(j, 7).Value
        End If
        
        'MsgBox (ticker)
       
                
            Next j
            
           
            
            'Put final summary out for last ticker
            Yearly_Change = (Close_Price - Open_Price)
            Percent_Change = (Yearly_Change / Open_Price)

            Greatest_Increase = Application.WorksheetFunction.Max(Sheets(i).Range("K:K"))
            Greatest_Decrease = Application.WorksheetFunction.Min(Sheets(i).Range("K:K"))
            Greatest_Volume = Application.WorksheetFunction.Max(Sheets(i).Range("L:L"))
            
              'format to color change : left out last row before so added this one down here
                 If Yearly_Change < 0 Then
                    Sheets(i).Cells(Display_Row, 10).Interior.ColorIndex = 3

                Else
                    Sheets(i).Cells(Display_Row, 10).Interior.ColorIndex = 4

                End If
                
                'format to percent sign : left out last row before so added this one down here
                Sheets(i).Cells(Display_Row, 11).NumberFormat = "0.00%"
                Percent_Change = Application.WorksheetFunction.RoundUp(Percent_Change, 2)


       

        'Put output in Display Row
                Sheets(i).Cells(Display_Row, 9).Value = prev_ticker
                Sheets(i).Cells(Display_Row, 10).Value = Yearly_Change
                Sheets(i).Cells(Display_Row, 11).Value = Percent_Change
                Sheets(i).Cells(Display_Row, 12).Value = volume

                Sheets(i).Cells(2, 16).Value = Greatest_Increase
                Sheets(i).Cells(3, 16).Value = Greatest_Decrease
                Sheets(i).Cells(4, 16).Value = Greatest_Volume

                'fromat to percent sign
                Sheets(i).Cells(2, 16).NumberFormat = "0.00%"
                Sheets(i).Cells(3, 16).NumberFormat = "0.00%"
                
            Dim rng1 As Range
            Dim rng2 As Range
            Dim rng3 As Range

            Set rng1 = Sheets(i).Range("K:K")
            Set rng2 = Sheets(i).Range("L:L")
            Set rng3 = Sheets(i).Range("I:I")
                

               Sheets(i).Cells(2, 15).Value = Application.Index(rng3, Application.Match(Application.Max(rng1), rng1, 0))
               Sheets(i).Cells(3, 15).Value = Application.Index(rng3, Application.Match(Application.Min(rng1), rng1, 0))
               Sheets(i).Cells(4, 15).Value = Application.Index(rng3, Application.Match(Application.Max(rng2), rng2, 0))
              

                'Need to fit text to column
       Sheets(i).Range("J1:L1").EntireColumn.AutoFit
       Sheets(i).Range("N1:P1").EntireColumn.AutoFit



                
        
  Next i
    
End Sub