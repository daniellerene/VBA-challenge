Attribute VB_Name = "Module2"
Sub Main()
    Dim ws As Worksheet
    Dim Ticker_symbol As String
    Dim Total_Vol As Double
    Dim Table_ticker As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim stockPriceAlreadyCaptured As Boolean
    Dim Opening_Price As Double
    Dim closing_Price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Sheets
        ws.Activate ' Activate the current worksheet
        
        ' Reset variables for each sheet
        Total_Vol = 0
        Table_ticker = 2
        stockPriceAlreadyCaptured = False

        lastRow = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker_symbol = Cells(i, 1).Value
                Total_Vol = Total_Vol + Cells(i, 7).Value
                Range("I" & Table_ticker).Value = Ticker_symbol
                Range("L" & Table_ticker).Value = Total_Vol
                Table_ticker = Table_ticker + 1
                stockPriceAlreadyCaptured = False
                Total_Vol = 0
            Else
                Total_Vol = Total_Vol + Cells(i, 7).Value
            End If
            
            If stockPriceAlreadyCaptured = False Then
                Opening_Price = Cells(i, 3).Value
                stockPriceAlreadyCaptured = True
            End If
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                closing_Price = Cells(i, 6).Value
                yearly_change = Opening_Price - closing_Price
                Range("J" & Table_ticker).Value = yearly_change
                percent_change = (yearly_change / Opening_Price)
                Range("K" & Table_ticker).Value = percent_change
                yearly_change = 0
                percent_change = 0
                stockPriceAlreadyCaptured = False
            End If
        Next i
    Next ws
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String

    ' Initialize variables
    maxPercentIncrease = 0
    maxPercentDecrease = 0
    maxTotalVolume = 0

    ' Loop through the results to find the maximum values
    For i = 2 To Table_ticker - 1
        ' Check for greatest % increase
        If Range("K" & i).Value > maxPercentIncrease Then
            maxPercentIncrease = Range("K" & i).Value
            maxPercentIncreaseTicker = Range("I" & i).Value
        End If

        ' Check for greatest % decrease
        If Range("K" & i).Value < maxPercentDecrease Then
            maxPercentDecrease = Range("K" & i).Value
            maxPercentDecreaseTicker = Range("I" & i).Value
        End If

        ' Check for greatest total volume
        If Range("L" & i).Value > maxTotalVolume Then
            maxTotalVolume = Range("L" & i).Value
            maxTotalVolumeTicker = Range("I" & i).Value
        End If
    Next i

    ' Display the results in cells P2:Q5
    Range("P2").Value = "Greatest % Increase"
    Range("P3").Value = "Greatest % Decrease"
    Range("P4").Value = "Greatest Total Volume"
    Range("P5").Value = ""

    Range("Q2").Value = maxPercentIncreaseTicker
    Range("Q3").Value = maxPercentDecreaseTicker
    Range("Q4").Value = maxTotalVolumeTicker
    Range("Q5").Value = ""

    
End Sub
