Attribute VB_Name = "mdUpdate"
Option Explicit
Dim log, holdings, amountsheet, pricesheet, perf, port As Worksheet
Sub UpdateAll()

    Unprotect
    Application.ScreenUpdating = False
    
    Call HoldingsUpdate
    Call UpdatePrices
    Call UpdateAmounts
    Call UpdateValues
    Call UpdatePortfolio

    Sheets("View").Select
    MsgBox ("Portfolio sucessfully updated!")
    
    Dim conn
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn
    
    Calculate
    view_Lock
    log_Lock
    detail_Lock
    port_Lock
    corr_Lock
    sec_Lock
    Application.ScreenUpdating = True

End Sub



Sub UpdatePrices()
    
    ' Define arrays from public functions and integers for number of holdings
    ' ----------------------------------------------------------------------------------------------------
    Dim holdings(), TradeDates() As Variant
    Dim HoldingsCount, i As Integer
    
    
    ' Define tickers of holdings and ranges for dates
    ' ----------------------------------------------------------------------------------------------------
    
    holdings = func_Holdings()

    HoldingsCount = UBound(holdings)

    TradeDates = func_TradeDates

    
    ' Change to relevant sheet and import dates and prices from first to last trading date
    ' ----------------------------------------------------------------------------------------------------
    
    Sheets("Prices").Activate
    Cells.ClearContents
    Range("A1").Select
    
    Call func_GetDates(TradeDates(1), TradeDates(2))
    
    For i = 1 To HoldingsCount
        ActiveCell.Offset(0, i).Value = holdings(i)
        Call func_GetPrices(TradeDates(1), TradeDates(2), holdings(i), i)
    Next i

    
End Sub

Sub UpdateAmounts()

    ' Define variables and arguments for sumif operation
    ' -------------------------------------------------------------------------------------
    Dim HoldingsCount, DaysCount, Amount, i, j As Integer
    Dim Ticker As String
    Dim tradeDay As Date
    
    Dim SumRng As Range 'the range i want to sum
    Dim CritRng1 As Range 'criteria range 1
    Dim CritRng2 As Range 'criteria range 2
    
    Set amountsheet = ThisWorkbook.Worksheets("Amounts")
    Set pricesheet = ThisWorkbook.Worksheets("Prices")
    Set log = ThisWorkbook.Worksheets("Log")
    
    
    
    ' Copy and Paste the trading days and holdings from other sheet
    ' -------------------------------------------------------------------------------------
    
    amountsheet.Activate
    Cells.ClearContents

    pricesheet.Activate
    Columns("A:A").Select
    Selection.Copy
    
    amountsheet.Activate
    Columns("A:A").Select
    ActiveSheet.Paste
    
    pricesheet.Activate
    Rows("1:1").Select
    Selection.Copy
    
    amountsheet.Activate
    Rows("1:1").Select
    ActiveSheet.Paste
    
    
    ' Get variables with length and width of table for holding amounts
    ' --------------------------------------------------------------------------------------
    
    DaysCount = Application.WorksheetFunction.CountA(Range("A:A"))
    
    HoldingsCount = UBound(func_Holdings)


    ' Set ranges for sumifs operatierons
    ' --------------------------------------------------------------------------------------

    log.Activate
    Set SumRng = log.Range("G:G")
    Set CritRng1 = log.Range("E:E")
    Set CritRng2 = log.Range("C:C")
    

    ' Nested looping column by column
    ' --------------------------------------------------------------------------------------

    amountsheet.Activate
    For i = 1 To HoldingsCount
        For j = 1 To DaysCount - 1
            
            Ticker = Cells(1, i + 1).Value
            tradeDay = Cells(j + 1, 1).Value
            
            Amount = WorksheetFunction.SumIfs(SumRng, CritRng1, Ticker, CritRng2, "<=" & Format(tradeDay, 0))
            Cells(j + 1, i + 1).Value = Amount
        
        Next j
    Next i
    

End Sub



Sub UpdateValues()

    ' Define variables and arguments for sumif operation
    ' -------------------------------------------------------------------------------------
    Dim HoldingsCount, DaysCount, i, j As Integer
    
    
    ' Copy and Paste the trading days and holdings from other sheet
    ' -------------------------------------------------------------------------------------
    
    Sheets("Values").Activate
    Cells.ClearContents

    Sheets("Prices").Activate
    Columns("A:A").Select
    Selection.Copy
    
    Sheets("Values").Activate
    Columns("A:A").Select
    ActiveSheet.Paste
    
    Sheets("Prices").Activate
    Rows("1:1").Select
    Selection.Copy
    
    Sheets("Values").Activate
    Rows("1:1").Select
    ActiveSheet.Paste
    
    
    ' Get variables with length and width of table for holding amounts
    ' --------------------------------------------------------------------------------------
    
    DaysCount = Application.WorksheetFunction.CountA(Range("A:A"))
    
    HoldingsCount = UBound(func_Holdings)
    

    ' Nested looping column by column
    ' --------------------------------------------------------------------------------------

    Sheets("Values").Activate
    For i = 1 To HoldingsCount
        For j = 1 To DaysCount - 1
            
            Cells(j + 1, i + 1).Value = Sheets("Prices").Cells(j + 1, i + 1) * Sheets("Amounts").Cells(j + 1, i + 1)
        
        Next j
    Next i

End Sub

Sub UpdatePortfolio()
    Dim record As Long
    Dim stock As Long
    Dim i, j As Integer
    Dim valuesheet As Worksheet
    Dim portfolio As Worksheet
    Dim valueRange As String
    Dim sumRange, dateRange As String
    Dim trans As Long
    Dim LastRow As Long
    
    Set valuesheet = ThisWorkbook.Worksheets("Values")
    Set portfolio = ThisWorkbook.Worksheets("PortfolioOverall")
    Set log = ThisWorkbook.Worksheets("Log")
    Set perf = ThisWorkbook.Worksheets("Performance")
    
    portfolio.Unprotect
    
    portfolio.Range("B5:G1000").ClearContents
    
    stock = WorksheetFunction.CountA(valuesheet.Rows(1)) 'last coloumn of valuesheet
    record = WorksheetFunction.CountA(valuesheet.Range("A:A")) 'last row of valuesheet
    trans = WorksheetFunction.CountA(log.Range("C:C")) + 3 ' last row of log sheet
    sumRange = "I5:I" & trans 'sum range for value
    dateRange = "C5:C" & trans 'criteria range for value

    For i = 2 To record
        valueRange = "B" & i & ":" & Cells(i, stock).Address 'range to sum values for a specific date
        portfolio.Cells(i + 3, 2) = valuesheet.Cells(i, 1) 'get date from value sheet
        portfolio.Cells(i + 3, 3) = WorksheetFunction.Sum(valuesheet.Range(valueRange)) 'sum values for a specific date
        portfolio.Cells(i + 3, 4) = ThisWorkbook.Worksheets("View").Range("wsCash") - WorksheetFunction.SumIf(log.Range(dateRange), _
        "<=" & Format(portfolio.Cells(i + 3, 2), 0), log.Range(sumRange)) 'get the cash value for a specific date
        portfolio.Cells(i + 3, 5) = portfolio.Cells(i + 3, 3) + portfolio.Cells(i + 3, 4) 'total value for a specific date
        portfolio.Cells(i + 3, 6) = portfolio.Cells(i + 3, 4) / portfolio.Cells(i + 3, 5) 'cash ratio for a specific date
        portfolio.Cells(i + 3, 7) = 1 - portfolio.Cells(i + 3, 6) 'stock ratio for a specific date
    Next i
    'updating chart
    '_________________________________________________________________________________________________________________________
    Dim tradeRange As String
    Dim totalRange, cashRange, stockRange As String
    Dim endRange As String
    LastRow = record + 3
    tradeRange = "B5:B" & LastRow
    totalRange = "E5:E" & LastRow
    cashRange = "F5:F" & LastRow
    stockRange = "G5:G" & LastRow
    endRange = "B" & LastRow
    
    portfolio.Range("wealth_start") = portfolio.Range("B5").Value
    portfolio.Range("wealth_end") = portfolio.Range(endRange).Value
    
    perf.Range("perf_start") = portfolio.Range("B5").Value
    perf.Range("perf_end") = portfolio.Range(endRange).Value
    
    
    With portfolio.ChartObjects("PortfolioChart1").Chart
        .SeriesCollection(1).Values = portfolio.Range(totalRange)
        .SeriesCollection(1).XValues = portfolio.Range(tradeRange)
    End With
    
    With portfolio.ChartObjects("PortfolioChart2").Chart
        .SeriesCollection("Cash").Values = portfolio.Range(cashRange)
        .SeriesCollection("Cash").XValues = portfolio.Range(tradeRange)
        .SeriesCollection("Stocks").Values = portfolio.Range(stockRange)
        .SeriesCollection("Stocks").XValues = portfolio.Range(tradeRange)
    End With
    
    
End Sub

Sub logChartUpdate()
    Dim log As Worksheet
    Set log = ThisWorkbook.Worksheets("Log")
    
    Application.ScreenUpdating = False

    Dim LastRow As Integer
    Dim myRange, myLabel As String

    With log
        LastRow = WorksheetFunction.CountA(log.Range("$C:$C")) + 4
        'Get the LastRow and choose the data series for chart
        myRange = "I5:I" & LastRow
        myLabel = "A5:A" & LastRow
        'Update Chart
        With .ChartObjects("Log Chart").Chart
        
            .SeriesCollection(1).Values = log.Range(myRange)
            .SeriesCollection(1).XValues = log.Range(myLabel)
        
        End With
    End With

End Sub

Sub EDT_Calculate()
    ThisWorkbook.Sheets("View").Range("EDT_NOW").Calculate
    EDT_Timer
End Sub

Sub StopUpdatingPrice()
    On Error Resume Next
    Application.OnTime earliesttime:=RunWhen, _
    procedure:=cRunWhat, schedule:=False
End Sub

