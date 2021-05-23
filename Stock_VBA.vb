Sub Stock()

'Loop through all sheets
Dim ws as Worksheet
Dim TickerName as String
Dim TotalStockVolume as Double
Dim LastRow as Long
Dim TableRow as Integer
Dim OpenPrice as Double
Dim ClosePrice as Double
Dim YearlyChange as Double
Dim PercentageChange as Double  

'Inserting Table Headers

For Each ws in Worksheets

ws.Range ("I1").Value = "Ticker"
ws.Range ("J1").Value = "Yearly_Change"
ws.Range ("K1").Value = "Percetage_Change"
ws.Range ("L1").Value = "Total_Stock_Volume"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

OpenPrice = ws.Cells(2,3).Value

TotalStockVolume = 0

TableRow = 2

    For i = 2 to LastRow

        'Input Ticker Names
        If Cells(i + 1 , 1).Value <> Cells(i,1) Then
            TickerName = Cells(i, 1).Value
            ws.Cells(TableRow,9).Value = TickerName
            
            'Calculating closing price
            ClosePrice = ws.Cells(i,6).Value

            'Calculating yearly change
            YearlyChange = ClosePrice - OpenPrice

            'Input Yearly change
            ws.Cells(TableRow,10).Value = YearlyChange


            'Coloring Cells for positive and negative yearly change
            If YearlyChange > 0 Then
            ws.Cells(TableRow,10).Interior.ColorIndex = 4
            Elseif YearlyChange <= 0 Then
            ws.Cells(TableRow,10).Interior.ColorIndex = 3
            End if

            
                        
            'Calculate Percentage Change
            'If fucntion to remove zero open price

            If OpenPrice <> 0 Then
            PercentageChange = (YearlyChange/OpenPrice) * 100
            Elseif OpenPrice = 0 Then
            PercentageChange = 0
            End if

            'Input Percentage Change
            ws.cells(TableRow,11).Value = PercentageChange   
            
                        
            'Adding % sign
            ws.Range("K" & TableRow).Value = (CStr(PercentageChange & "%"))
            
            
            'Reset Open Price
            OpenPrice = ws.Cells(i + 1, 3).Value


            
            'Adding Total Stock Volume
            TotalStockVolume = TotalStockVolume + ws.cells(i,7).Value

            ws.Range("L" & TableRow).Value = TotalStockVolume
            
            TableRow = TableRow + 1

            TotalStockVolume = 0
            
        Else 
            TotalStockVolume = TotalStockVolume + ws.cells(i,7).Value
                                   
        End If
       

    Next i

Next ws

End Sub

