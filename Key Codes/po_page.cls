VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim startDay, endDay As Date
  

Sub PortChartUpdate(startDay, endDay)
    Dim i As Long
    Dim LastRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim selectDateRange, selectTotalRange, selectCashRange, selectStockRange As String
    
    Set port = ThisWorkbook.Worksheets("PortfolioOverall")
    port.Unprotect
    
    LastRow = WorksheetFunction.CountA(port.Range("B:B")) + 2
    
    For i = 5 To LastRow
        If port.Cells(i, 2) = startDay Then
            startRow = i
        ElseIf port.Cells(i, 2) = endDay Then
            endRow = i
        End If
    Next i
    
    If startRow = 0 Then
        startDay = DateAdd("d", 2, startDay)
        For i = 5 To LastRow
            If port.Cells(i, 2) = startDay Then
                startRow = i
            End If
        Next
    End If
    
    If endRow = 0 Then
        endDay = DateAdd("d", 2, endDay)
        For i = 5 To LastRow
            If port.Cells(i, 2) = endDay Then
                endDay = i
            End If
        Next
        
        If endRow = 0 Then
            endRow = LastRow
        End If
    End If
    
    selectDateRange = "B" & startRow & ":B" & endRow
    selectTotalRange = "E" & startRow & ":E" & endRow
    selectCashRange = "F" & startRow & ":F" & endRow
    selectStockRange = "G" & startRow & ":G" & endRow
    
    With port.ChartObjects("PortfolioChart1").Chart
        .SeriesCollection(1).Values = port.Range(selectTotalRange)
        .SeriesCollection(1).XValues = port.Range(selectDateRange)
    End With
    
    With port.ChartObjects("PortfolioChart2").Chart
        .SeriesCollection("Cash").Values = port.Range(selectCashRange)
        .SeriesCollection("Cash").XValues = port.Range(selectDateRange)
        .SeriesCollection("Stocks").Values = port.Range(selectStockRange)
        .SeriesCollection("Stocks").XValues = port.Range(selectDateRange)
    End With
    port_Lock
        
End Sub



Private Sub Worksheet_SelectionChange(ByVal Target As Range)

        Dim tradeRange As String
        Dim LastRow As Long
        Dim firstDay, lastDay As Date

        Dim port As Worksheet
    
        Set port = ThisWorkbook.Worksheets("PortfolioOverall")
        port.Unprotect
        

    'click target cells to choose date from calendar form and update the chart
        
        If Target.Address Like "*:*" Then Exit Sub 'exist multi-cells

        If Not Intersect(Target, [J4,M4]) Is Nothing Then 'assign the event range
            frmCalendar.ShowDate Target
        End If
    

            
        LastRow = WorksheetFunction.CountA(port.Range("B:B")) + 2
    
        startDay = port.Range("wealth_start").Value
        endDay = port.Range("wealth_end").Value
    
        firstDay = port.Range("B5").Value
        tradeRange = "B" & LastRow
        lastDay = port.Range(tradeRange).Value
    
        If DateDiff("d", startDay, firstDay) > 0 Then
            MsgBox ("Start Date cannot be earlier than " & firstDay)
            port.Range("wealth_start") = firstDay
            Exit Sub
        End If
    
        If DateDiff("d", lastDay, endDay) > 0 Then
            MsgBox ("End Date cannot be later than " & lastDay)
            port.Range("wealth_end") = lastDay
            Exit Sub
        End If
    
        If DateDiff("d", endDay, startDay) > 0 Then
            MsgBox ("End Date cannot be later the Start Date")
            port.Range("wealth_start") = firstDay
            port.Range("wealth_start") = lastDay
            Exit Sub
        End If
    
        Call PortChartUpdate(startDay, endDay)
        
        port.Protect

End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    'Activate the event right after select the date
    On Error Resume Next
    If Target.Address = "$J$4" Or Target.Address = "$M$4" Then
    Range("K4").Select
    End If
    
End Sub
