Attribute VB_Name = "Module1"
Sub MultiYearStockScript()
    '
    ' MultiYearStockScript Macro
    '
    ' Date: 2023/05/07
    ' Author: Prachi Patel
    'Module 2 - VBA Challenge
    ' Description: Loop through all the sheets to create necessary columns for analysed data
    '

    ' Looping through all the sheet in a worksheet
    For Each wrksheet In Worksheets

        'Clear exisiting data from all the columns that we will be writing data to
        wrksheet.Columns("I:Q").EntireColumn.Delete

        ' First Part header
        'Create all the column headers that are required
        wrksheet.Cells(1, 9).Value = "Ticker"
        wrksheet.Cells(1, 10).Value = "Yearly Change"
        wrksheet.Cells(1, 11).Value = "Percent Change"
        wrksheet.Cells(1, 12).Value = "Total Stock Volume"


        'Start looping through all the rows
        'Find the last index of row in Column A which is empty - where the data ends

        'Defining the variable LastRowInA for column A
        Dim LastRowInA As Long

        'Setting the count to last row which is not empty
        LastRowInA = wrksheet.Cells(Rows.Count, 1).End(xlUp).Row


        'Defining variable TickerCount = variable to store the index of the filled the ticker row
        Dim TickerCount As Long

        'Setting the index to start from row 2, ignoring the headers
        TickerCount = 2

        'Defining variable i to hold the current row looping
        Dim i As Long

        'Defining variable j to hold the start row of ticker block
        Dim j As Long

        j = 2

        'Defining variable PercentageChange for percent change calculation
        Dim PercentageChange   As Double


        'Loop through all rows
        For i = 2 To LastRowInA

            'Verify if the ticker name is different
            If wrksheet.Cells(i + 1, 1).Value <> wrksheet.Cells(i, 1).Value Then

                'Update ticker name in COL I
                wrksheet.Cells(TickerCount, 9).Value = wrksheet.Cells(i, 1).Value

                'Populate the yearly change in COL J
                wrksheet.Cells(TickerCount, 10).Value = wrksheet.Cells(i, 6).Value - wrksheet.Cells(j, 3).Value

                'Apply Conditional formating
                If wrksheet.Cells(TickerCount, 10).Value < 0 Then

                    'Set cell background color to red for the percent decrease
                    wrksheet.Cells(TickerCount, 10).Interior.ColorIndex = 3

                Else

                    'Set cell background color to green for increase in percentage
                    wrksheet.Cells(TickerCount, 10).Interior.ColorIndex = 4

                End If

                'Populate percentage change in COL J
                If wrksheet.Cells(j, 3).Value <> 0 Then
                    PercentageChange = ((wrksheet.Cells(i, 6).Value - wrksheet.Cells(j, 3).Value) / wrksheet.Cells(j, 3).Value)

                    'Apply percentage formatting to that cell
                    wrksheet.Cells(TickerCount, 11).Value = Format(PercentageChange, "Percent")

                Else

                    wrksheet.Cells(TickerCount, 11).Value = Format(0, "Percent")

                End If

                'Poplulate total stock volume in COL L
                wrksheet.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(wrksheet.Cells(j, 7), wrksheet.Cells(i, 7)))

                'Move the tickerCount index to next
                TickerCount = TickerCount + 1

                'Increment start row index of the ticker block
                j = i + 1

            End If

        Next i

        ' Second Part hearders
        ' Fill in the header values
        wrksheet.Cells(1, 16).Value = "Ticker"
        wrksheet.Cells(1, 17).Value = "Value"
        wrksheet.Cells(2, 15).Value = "Greatest % Increase"
        wrksheet.Cells(3, 15).Value = "Greatest % Decrease"
        wrksheet.Cells(4, 15).Value = "Greatest Total Volume"


        'Defining the variable LastRowInI for column I
        Dim LastRowInI    As Long

        'Find last non-blank cell in column I
        LastRowInI = wrksheet.Cells(Rows.Count, 9).End(xlUp).Row

        'Defining the variable to store the greatest increase calculation
        Dim GreatestIncrease   As Double

        'Defining the variable to store the greatest decrease calculation
        Dim GreatestDecrease   As Double

        'Defining the variable to store the greatest total volume
        Dim GreatestTotalVolume    As Double

        'Set the inital values to compare with the next ticker
        GreatestTotalVolume = wrksheet.Cells(2, 12).Value
        GreatestIncrease = wrksheet.Cells(2, 11).Value
        GreatestDecrease = wrksheet.Cells(2, 11).Value

        'Loop through all the sorted ticker data
        For i = 2 To LastRowInI

            'if greatest total volume is larger than the next value then set and apply
            If wrksheet.Cells(i, 12).Value > GreatestTotalVolume Then
                GreatestTotalVolume = wrksheet.Cells(i, 12).Value
                wrksheet.Cells(4, 16).Value = wrksheet.Cells(i, 9).Value

            Else

                GreatestTotalVolume = GreatestTotalVolume

            End If

            'if greatest increase is larger then next value then set and apply
            If wrksheet.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = wrksheet.Cells(i, 11).Value
                wrksheet.Cells(2, 16).Value = wrksheet.Cells(i, 9).Value

            Else

                GreatestIncrease = GreatestIncrease

            End If

            'if greatest decrease is smaller then the next value, set and apply
            If wrksheet.Cells(i, 11).Value < GreatestDecrease Then
                GreatestDecrease = wrksheet.Cells(i, 11).Value
                wrksheet.Cells(3, 16).Value = wrksheet.Cells(i, 9).Value

            Else

                GreatestDecrease = GreatestDecrease

            End If

            'Apply formatting
            wrksheet.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            wrksheet.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            wrksheet.Cells(4, 17).Value = Format(GreatestTotalVolume, "Scientific")

        Next i

        Dim WorksheetName As String
        'Get the WorksheetName to autofit the column width
        WorksheetName = wrksheet.Name
        Worksheets(WorksheetName).Columns("A:Z").AutoFit


    Next wrksheet

End Sub

