VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim perf_startDay, perf_endDay As Date
Dim port, perf, pricesheet, amountsheet, holdings As Worksheet
Dim perf_ticker As String
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'click target cells to choose date from calendar form and update the chart
    '--------------------------------------------------------------------------
    Dim LastRow As Long
    Dim firstDay, lastDay, startDay, endDay As Date
    Dim perf_ticker As String
    
    Set port = ThisWorkbook.Worksheets("PortfolioOverall")
    Set perf = ThisWorkbook.Worksheets("Performance")
    
    perf.Unprotect
    
    On Error Resume Next
    
    If Target.Address Like "*:*" Then Exit Sub 'exist multi-cells

    If Not Intersect(Target, [F5,I5]) Is Nothing Then 'assign the event range
        frmCalendar.ShowDate Target
    End If
    
    Dim tradeRange As String
    
    LastRow = WorksheetFunction.CountA(port.Range("B:B")) + 2
    
    startDay = perf.Range("perf_start").Value
    endDay = perf.Range("perf_end").Value
    
    firstDay = port.Range("B5").Value
    tradeRange = "B" & LastRow
    lastDay = port.Range(tradeRange).Value
    
    If DateDiff("d", startDay, firstDay) > 0 Then
        MsgBox ("Start Date cannot be earlier than " & firstDay)
        perf.Range("perf_start").Value = firstDay
        Exit Sub
    End If
    
    If DateDiff("d", lastDay, endDay) > 0 Then
        MsgBox ("End Date cannot be later than " & lastDay)
        perf.Range("perf_end").Value = lastDay
        
        Exit Sub
    End If
    
    If DateDiff("d", endDay, startDay) > 0 Then
        MsgBox ("End Date cannot be later the Start Date")
        perf.Range("perf_start").Value = firstDay
        perf.Range("perf_end").Value = lastDay
        Exit Sub
    End If
    
    perf_ticker = perf.Range("select_ticker").Value
    
    Call Perf_ChartUpdate(perf_startDay, perf_endDay, perf_ticker)
    
    perf.Protect
    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    

    Dim tradeRange, tickerRange As String
    Dim tradeDay(), tradeTicker()
    Dim lastDay As Long
    Dim lastStock As Integer
    
    
    Set perf = ThisWorkbook.Worksheets("Performance")
    Set pricesheet = ThisWorkbook.Worksheets("Prices")
    Set holdings = ThisWorkbook.Worksheets("Holdings")
    
    perf.Unprotect
    
    On Error Resume Next
    'Activate the event right after select the date
    If Target.Name.Name = "perf_start" Or Target.Name.Name = "perf_end" Then
        Range("G5").Select
    End If
    
    'select from ticker list to update chart
    '-------------------------------------------------------------------------------------
    If Target.Name.Name = "select_ticker" Then
        lastDay = WorksheetFunction.CountA(pricesheet.Range("A:A"))
        lastStock = WorksheetFunction.CountA(holdings.Range("A:A"))
        
        tickerRange = "A2:A" & lastStock
        tradeTicker = holdings.Range(tickerRange).Value
        tradeTicker = Application.Transpose(tradeTicker)
        
        With perf.Range("select_ticker").Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=Join(tradeTicker, ",")
            .InCellDropdown = True
        End With
    

    perf_ticker = perf.Range("select_ticker").Value
    
    Call Perf_ChartUpdate(perf_startDay, perf_endDay, perf_ticker)
    
    End If
    
    perf.Protect

End Sub

Sub Perf_ChartUpdate(perf_startDay, perf_endDay, perf_ticker)
    Dim i As Long
    Dim LastRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim j, tickerCol, LastTicker As Integer
    Dim selectDateRange, selectPriceRange, selectAmountRange As String

    On Error Resume Next
    Set perf = ThisWorkbook.Worksheets("Performance")
    Set pricesheet = ThisWorkbook.Worksheets("Prices")
    Set amountsheet = ThisWorkbook.Worksheets("Amounts")
    Set port = ThisWorkbook.Worksheets("PortfolioOverall")
    Set port = ThisWorkbook.Worksheets("PortfolioOverall")
    
    LastRow = WorksheetFunction.CountA(pricesheet.Range("A:A"))
    LastTicker = WorksheetFunction.CountA(pricesheet.Rows(1))
    
    perf_startDay = perf.Range("perf_start").Value
    perf_endDay = perf.Range("perf_end").Value
    
    
    For j = 2 To LastTicker
        If pricesheet.Cells(1, j) = perf_ticker Then tickerCol = j
    Next j
    
    For i = 1 To LastRow
        If pricesheet.Cells(i, 1) = perf_startDay Then
            startRow = i
        ElseIf pricesheet.Cells(i, 1) = perf_endDay Then
            endRow = i
        End If
    Next i
    
    
    selectDateRange = pricesheet.Cells(startRow, 1).Address _
    & ":" & pricesheet.Cells(endRow, 1).Address
    selectPriceRange = pricesheet.Cells(startRow, tickerCol).Address _
    & ":" & pricesheet.Cells(endRow, tickerCol).Address
    selectAmountRange = amountsheet.Cells(startRow, tickerCol).Address _
    & ":" & amountsheet.Cells(endRow, tickerCol).Address
    
    
    With perf.ChartObjects("Perf_Chart").Chart
        .SeriesCollection("Prices").Values = pricesheet.Range(selectPriceRange)
        .SeriesCollection("Prices").XValues = pricesheet.Range(selectDateRange)
        .SeriesCollection("Amounts").Values = amountsheet.Range(selectAmountRange)
        .SeriesCollection("Amounts").XValues = amountsheet.Range(selectDateRange)
    End With
        
End Sub

