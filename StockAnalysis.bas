Attribute VB_Name = "Module1"
Sub StockAnalysis()

    ' Declare variables
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentIncrease As Double
    Dim maxPercentDecreaseTicker As String
    Dim maxPercentDecrease As Double
    Dim maxTotalVolumeTicker As String
    Dim maxTotalVolume As Double
    
    ' Set initial values
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ticker = ""
    openingPrice = 0
    closingPrice = 0
    yearlyChange = 0
    percentChange = 0
    totalVolume = 0
    outputRow = 2
    maxPercentIncreaseTicker = ""
    maxPercentIncrease = 0
    maxPercentDecreaseTicker = ""
    maxPercentDecrease = 0
    maxTotalVolumeTicker = ""
    maxTotalVolume = 0
    
    ' Output column headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    ' Loop through all rows
    For i = 2 To lastRow
        
        ' Check if ticker has changed
        If Cells(i, 1).Value <> ticker Then
            
            ' Output previous ticker data
            If ticker <> "" Then
                Range("I" & outputRow).Value = ticker
                Range("J" & outputRow).Value = yearlyChange
                Range("K" & outputRow).Value = percentChange
                Range("L" & outputRow).Value = totalVolume
                outputRow = outputRow + 1
                
                ' Check if new max percent increase or decrease or total volume
                If percentChange > maxPercentIncrease Then
                    maxPercentIncreaseTicker = ticker
                    maxPercentIncrease = percentChange
                ElseIf percentChange < maxPercentDecrease Then
                    maxPercentDecreaseTicker = ticker
                    maxPercentDecrease = percentChange
                End If
                If totalVolume > maxTotalVolume Then
                    maxTotalVolumeTicker = ticker
                    maxTotalVolume = totalVolume
                End If
                
            End If
            
            ' Reset values for new ticker
            ticker = Cells(i, 1).Value
            openingPrice = Cells(i, 3).Value
            totalVolume = 0
            
        End If
        
        ' Calculate closing price and total volume for current ticker
        closingPrice = Cells(i, 6).Value
        totalVolume = totalVolume + Cells(i, 7).Value
        
        ' Calculate yearly change and percent change
        yearlyChange = closingPrice - openingPrice
        If openingPrice <> 0 Then
            percentChange = yearlyChange / openingPrice
        End If
        
        ' Check if last row for current ticker
        If i = lastRow Then
            Range("I" & outputRow).Value = ticker
            Range("J" & outputRow).Value = yearlyChange
            Range("K" & outputRow).Value = percentChange
            Range("L" & outputRow).Value = totalVolume
            
            ' Check if new max percent increase
            If percentChange > maxPercentIncrease Then
                maxPercentIncreaseTicker = ticker
                maxPercentIncrease = percentChange
            ElseIf percentChange < maxPercentDecrease Then
                maxPercentDecreaseTicker = ticker
                maxPercentDecrease = percentChange
            End If
            If totalVolume > maxTotalVolume Then
                maxTotalVolumeTicker = ticker
                maxTotalVolume = totalVolume
            End If
        End If
        
    Next i
    
    ' Output results for greatest % increase, greatest % decrease, and greatest total volume
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("O2").Value = maxPercentIncreaseTicker
    Range("O3").Value = maxPercentDecreaseTicker
    Range("O4").Value = maxTotalVolumeTicker
    Range("P2").Value = Format(maxPercentIncrease, "Percent")
    Range("P3").Value = Format(maxPercentDecrease, "Percent")
    Range("P4").Value = maxTotalVolume
    
End Sub

