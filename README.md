# VBA-challenge

Boot Camp Challenge 2 files.

    Sub Workbook_Analysis()
        Dim stocksheet As Worksheet
        For Each stocksheet In ThisWorkbook.Worksheets
            stocksheet.Select
            Call Sheet_Analysis
            stocksheet.Range("I:Q").Columns.AutoFit
        Next stocksheet

    End Sub
    Sub Sheet_Analysis()
        Dim i, lastrow, itemcount As Long
        Dim summ, annualChange, pctMin, pctMax, volumeMax As Double
        Dim priceFlag As Boolean
        Dim pctMinTicker, pctMaxTicker, volumeMaxTicker As String

        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        itemcount = 2
        summ = 0
        priceFlag = True
        pctMin = 1E+99
        pctMax = -1E+99
        volumeMax = -1E+99

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Annual Change"
        Cells(1, 11).Value = "% Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"

        For i = 2 To lastrow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Cells(itemcount, 9).Value = Cells(i, 1).Value

            closePrice = Cells(i, 6).Value
            annualChange = closePrice - openPrice
            Cells(itemcount, 10).Value = annualChange

                If annualChange = 0 Or openPrice = 0 Then
                    Cells(itemcount, 11).Value = 0

                    Else: Cells(itemcount, 11).Value = Format(annualChange / openPrice, "#.##%")

                End If

                summ = summ + Cells(i, 7).Value
                Cells(itemcount, 12).Value = summ

                If Cells(itemcount, 11).Value > pctMax Then

                    If Cells(itemcount, 11).Value = ".%" Then
                    Else
                        pctMax = Cells(itemcount, 11).Value
                        pctMaxTicker = Cells(itemcount, 9).Value
                    End If

                ElseIf Cells(itemcount, 11).Value < pctMin Then
                    pctMin = Cells(itemcount, 11).Value
                    pctMinTicker = Cells(itemcount, 9).Value

                ElseIf Cells(itemcount, 12).Value > volumeMax Then
                    volumeMax = Cells(itemcount, 12).Value
                    volumeMaxTicker = Cells(itemcount, 9).Value

                End If

                If annualChange < 0 Then
                    Cells(itemcount, 10).Interior.ColorIndex = 3
                    Cells(itemcount, 11).Interior.ColorIndex = 3

                ElseIf annualChange > 0 Then
                    Cells(itemcount, 10).Interior.ColorIndex = 4
                    Cells(itemcount, 11).Interior.ColorIndex = 4

                End If

                itemcount = itemcount + 1
                summ = 0
                priceFlag = True

            Else
                If priceFlag Then
                    openPrice = Cells(i, 3).Value
                    priceFlag = False
                End If

                summ = summ + Cells(i, 7).Value

            End If

        Next i

        Cells(2, 17).Value = Format(pctMax, "#.##%")
        Cells(3, 17).Value = Format(pctMin, "#.##%")
        Cells(4, 17).Value = volumeMax

        Cells(2, 16).Value = pctMaxTicker
        Cells(3, 16).Value = pctMinTicker
        Cells(4, 16).Value = volumeMaxTicker


    End Sub

Sources:
Jaimes, David. (2023) https://github.com/davidjaimes/yearly-stock-market-analysis/tree/master
