Attribute VB_Name = "Module1"

         
            
Sub stockVolume()
        ' Declare and set worksheet
    Dim ws As Worksheet
    

       
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
   
     
         ' Set  variable
        Dim totalStockVolume, perchange, yrchange, openprice, closeprice, maxpercent, minpercent, maxvolume As Double
        Dim ticker, maxpercentvalue, minpercentvalue, maxvolumevalue As String
        Dim lastrow As Long
                  
         ' initialize variable values
        openprice = 0
        closeprice = 0
        totalStockVolume = 0
        perchange = 0
        yrchange = 0

         Dim summaryTableRow As Long

        ' variable to hold the summary table starter row
            summaryTableRow = 2
        
        ' find the last row
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' for open price
            openprice = ws.Cells(2, 3).Value

        
        ' loop from row 2 to last row
         For Row = 2 To lastrow
    

        
            ' check to see if ticker changes
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                ' if change is successful, then
                
                ' set the ticker name
               ticker = ws.Cells(Row, 1).Value


                ' for close price
                 closeprice = ws.Cells(Row, 6).Value

               
               ' add the last charge from the row
               totalStockVolume = totalStockVolume + ws.Cells(Row, 7).Value
               
               ' set values for yearly and percent change

           
                ' calculate yearly change
                    yrchange = closeprice - openprice
                
                ' calculate percent change
                    perchange = (closeprice - openprice) / openprice

               
               ' add the ticker name to the I column in the summary table
               ws.Cells(summaryTableRow, 9).Value = ticker
               
               ' add the summary of stock volume to the L column in the summary table
               ws.Cells(summaryTableRow, 12).Value = totalStockVolume
               
               ' add the yearly change to the J column in the summary table
                ws.Cells(summaryTableRow, 10).Value = yrchange
                
                ' add the percent change to the K column in the summary table
                ws.Cells(summaryTableRow, 11).Value = perchange
                ws.Cells(summaryTableRow, 11).Style = "percent"
               
               
                   'conditional format the positive and negative change
                    If yrchange > 0 Then
                        ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
     
                    Else
                        ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
     
                    End If

                               
               ' next summary table row
               summaryTableRow = summaryTableRow + 1
               
               'reset the total
               totalStockVolume = 0
               
               ' reset the open price
               openprice = ws.Cells(Row + 1, 3).Value
               
               ' reset percent change
               perchange = 0
            
            
            Else
                 ' if ticker remains the same, then
                 
                 ' add on to the total charges from volume column
                 totalStockVolume = totalStockVolume + ws.Cells(Row, 7).Value
               
        End If
        
    

        Next Row
    
     Next ws
    
     For Each ws In ActiveWorkbook.Worksheets
        ws.Columns.AutoFit
    Next ws
    
End Sub







