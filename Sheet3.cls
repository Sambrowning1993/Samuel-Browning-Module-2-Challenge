VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockAnalysis()
    ' Declare all of the variables
    Dim StockVolume As Double
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim lastrow As Long

    ' Loop through all of the worksheets in the workbook
    For Each WS In Worksheets
    
        ' Initialize dataentryrow
        Dim dataentryrow
        dataentryrow = 1
        
        ' Headers for dataentryrow
        WS.Cells(dataentryrow, 9) = "Ticker"
        WS.Cells(dataentryrow, 10) = "Yearly Change"
        WS.Cells(dataentryrow, 11) = "Percent Change"
        WS.Cells(dataentryrow, 12) = "Total Stock Volume"

        ' State what the last row is
        lastrow = WS.Cells(WS.Rows.Count, 1).End(xlUp).Row

        ' Initial Opening Price
        OpeningPrice = WS.Cells(2, 3).Value

        ' Initial Stock Volume
        StockVolume = 0

        ' Loop through the rows
        For i = 2 To lastrow
            If WS.Cells(i, 1).Value <> WS.Cells(i + 1, 1).Value Then
                dataentryrow = dataentryrow + 1
                
                ' Ticker
                WS.Cells(dataentryrow, 9).Value = WS.Cells(i, 1).Value
                
                ' Closing Price
                ClosingPrice = WS.Cells(i, 6).Value
                
                ' Yearly Change
                YearlyChange = ClosingPrice - OpeningPrice
                WS.Cells(dataentryrow, 10).Value = YearlyChange
                
                ' Percent Change
                PercentChange = YearlyChange / OpeningPrice
                WS.Cells(dataentryrow, 11).Value = PercentChange
                
                ' Recalculate OpeningPrice for the new Ticker
                OpeningPrice = WS.Cells(i + 1, 3).Value
                
                ' Percent Change formatting
                WS.Cells(dataentryrow, 11).NumberFormat = "0.00%"
                
                ' Add last row and add data
                StockVolume = StockVolume + WS.Cells(i, 7).Value
                WS.Cells(dataentryrow, 12).Value = StockVolume
                
            Else
            
                ' Add the StockTotal
                StockTotal = StockTotal + WS.Cells(i, 7).Value
                
            End If
        Next i
    Next WS
End Sub
