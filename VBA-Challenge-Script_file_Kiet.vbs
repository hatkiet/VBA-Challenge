Sub StockScript()
    Dim headers() As Variant
    Dim wb As Workbook
    Dim MainWs As Worksheet
    Dim Summary_Table_Row As Long
    Dim Lastrow As Long
    Dim i As Long

    ' Initialize workbook and headers
    Set wb = ActiveWorkbook
    headers = Array("<ticker>", "<date>", "<open>", "<high>", "<low>", "<close>", "<vol>", "", "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", " ", " ", "Ticker", "Value")

    ' Loop through each worksheet in the workbook
    For Each MainWs In wb.Sheets
        ' Set headers and formatting
        SetHeadersAndFormat MainWs, headers
    Next MainWs

    ' Loop through each worksheet for data processing
    For Each MainWs In wb.Sheets
        ' Initialize variables
        Summary_Table_Row = 2
        Lastrow = MainWs.Cells(Rows.Count, 1).End(xlUp).Row

        ' Call subroutine to process data
        ProcessData MainWs, Summary_Table_Row, Lastrow
    Next MainWs
End Sub

'---------Subroutine to set the Headers and Format the outcome tatble

Sub SetHeadersAndFormat(ws As Worksheet, headers() As Variant)
    ' Set headers in the first row of the worksheet
    With ws
        .Rows(1).Value = " "
        For i = LBound(headers) To UBound(headers)
            .Cells(1, i + 1).Value = headers(i)
        Next i
        .Rows(1).VerticalAlignment = xlCenter
        .Rows(1).HorizontalAlignment = xlHAlignLeft
        .Cells.EntireColumn.ColumnWidth = 16
    End With
End Sub

' ------------ Sub rountine to Process Data--------------------
Sub ProcessData(ws As Worksheet, ByRef Summary_Table_Row As Long, Lastrow As Long)
    ' Define variables' types
    Dim Stock_ID As String
    Dim Total_Stock_Vol As Double
    Dim Initial_Price As Double
    Dim Final_Price As Double
    Dim Yearly_Price_Change As Double
    Dim Yearly_Price_Change_Percent As Double
    Dim Max_Stock_ID As String
    Dim Max_Percent As Double
    Dim Min_Percent As Double
    Dim Max_Vol_Stock_ID As String
    Dim Max_Vol As Double
    Dim Min_Stock_ID As String

    ' Initialize variables
    Stock_ID = " "
    Total_Stock_Vol = 0
    Initial_Price = 0
    Final_Price = 0
    Yearly_Price_Change = 0
    Yearly_Price_Change_Percent = 0
    Max_Stock_ID = " "
    Max_Percent = 0
    Min_Percent = 0
    Max_Vol_Stock_ID = " "
    Max_Vol = 0
    Min_Stock_ID = " "

    ' Identify the beginning price of the year for every stock starting with the first integer in column 3
    Initial_Price = ws.Cells(2, 3).Value

    ' Process data rows
    For i = 2 To Lastrow
        ' Identify a new ticker (stock) symbol every time a new value appears in the first column
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' If tickers are not the same, do calculate and store data
            Stock_ID = ws.Cells(i, 1).Value 
            Final_Price = ws.Cells(i, 6).Value
            Yearly_Price_Change = Final_Price - Initial_Price

            If Initial_Price <> 0 Then
                Yearly_Price_Change_Percent = (Yearly_Price_Change / Initial_Price) * 100
            End If

            ' Update total stock Volume every row if the same ticker
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value

            ' Populate summary table
            ws.Range("I" & Summary_Table_Row).Value = Stock_ID
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Price_Change

            ' Apply cell color based on yearly price change, positive in "green" and negative in "red"
            If Yearly_Price_Change > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf Yearly_Price_Change <= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If

            ' Populate additional summary data to summary table 
            ws.Range("K" & Summary_Table_Row).Value = (CStr(Yearly_Price_Change_Percent) & "%")
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Vol

            ' Done with one stock. Next, move to the next row in the summary table
            Summary_Table_Row = Summary_Table_Row + 1
            Initial_Price = ws.Cells(i + 1, 3).Value ' Reset the beginning price of new stock

            ' Use Subroutine "UpdateMinMaxValues"  to update maximum and minimum values
            UpdateMinMaxValues Yearly_Price_Change_Percent, Stock_ID, Max_Percent, Min_Percent, Max_Stock_ID, Min_Stock_ID

            If Total_Stock_Vol > Max_Vol Then
                Max_Vol = Total_Stock_Vol
                Max_Vol_Stock_ID = Stock_ID
            End If

            ' Reset variables
            Yearly_Price_Change_Percent = 0
            Total_Stock_Vol = 0
        Else
            ' Accumulate total stock ID volume for the same stock
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
        End If
    Next i

    ' Populate summary data in the worksheet
    ws.Cells(2, 16).Value = (CStr(Max_Percent) & "%")
    ws.Cells(3, 16).Value = (CStr(Min_Percent) & "%")
    ws.Cells(2, 15).Value = Max_Stock_ID
    ws.Cells(3, 15).Value = Min_Stock_ID
    ws.Cells(4, 15).Value = Max_Vol_Stock_ID
    ws.Cells(4, 16).Value = Max_Vol
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
End Sub

'----------Subroutine to Update minimum and maximum values-------------------------
Sub UpdateMinMaxValues(Percentage As Double, Ticker As String, ByRef MaxPercent As Double, ByRef MinPercent As Double, ByRef MaxTickerName As String, ByRef MinTickerName As String)
    If Percentage > MaxPercent Then
        MaxPercent = Percentage
        MaxTickerName = Ticker
    ElseIf Percentage < MinPercent Then
        MinPercent = Percentage
        MinTickerName = Ticker
    End If
End Sub