Attribute VB_Name = "mdFunctions"
Option Explicit


Public Function func_Holdings() As Variant()
    
    ' Define a dynamic array and integer variables that count number of holdings
    ' -----------------------------------------------------------------------------------------------------------------
    
    Dim holdings() As Variant
    Dim HoldingsCount, i As Integer


    ' Count the number of holdings and re-dimension the dynamic array to fixed size
    ' -----------------------------------------------------------------------------------------------------------------
    
    ThisWorkbook.Sheets("Holdings").Activate
    Range("A1").Select
    HoldingsCount = WorksheetFunction.CountA(Range("A:A")) - 1
    
    On Error Resume Next
    
    ReDim holdings(1 To HoldingsCount)
    
    
    ' Iterate through each row and store the ticker into an array
    ' -----------------------------------------------------------------------------------------------------------------
    For i = 1 To HoldingsCount
        holdings(i) = ActiveCell.Offset(i, 0).Value
    Next i
    
    func_Holdings = holdings
    

End Function


Public Function func_TradeDates() As Variant()

    Dim dateRange As Range
    Dim TradeDates(1 To 2) As Variant
    
    Sheets("Log").Select
    Set dateRange = Range("C:C")
    
    TradeDates(1) = Application.WorksheetFunction.Min(dateRange)
    TradeDates(2) = Application.WorksheetFunction.Max(dateRange)
    
    func_TradeDates = TradeDates

End Function


Public Function func_GetPrices(startDate, endDate, Ticker, Col)

    ' Dimension and transform data variables into required format
    ' --------------------------------------------------------------
    
    Dim StartDateSec, EndDateSec As Double

    StartDateSec = (startDate - DateSerial(1970, 1, 1)) * 86400
    EndDateSec = (endDate - DateSerial(1969, 12, 31)) * 86400
    
    ' Import data
    ' -------------------------------------------------------------

    With ActiveSheet.QueryTables.Add(Connection:= _
            "TEXT;https://query1.finance.yahoo.com/v7/finance/download/" & Ticker & "?period1=" & StartDateSec & "&period2=" & EndDateSec & "&interval=1d&events=history" _
            , Destination:=Cells(1, Col + 1))
            .Name = "table.csv?s=AAPL&d=4&e=10&f=2016&g=d&a=11&b=12&c=1980&ignore="
            .FieldNames = False
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlOverwriteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = False
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 850
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = True
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(9, 9, 9, 9, 9, 1, 9)
            .TextFileTrailingMinusNumbers = True
            On Error Resume Next
            .Refresh BackgroundQuery:=False
        End With
        
    Sheets("Prices").Activate
    Cells(1, Col + 1).Value = Ticker

End Function


Public Function func_GetDates(startDate, endDate)

    ' Dimension and transform data variables into required format
    ' --------------------------------------------------------------
    
    Dim StartDateSec, EndDateSec As Double

    StartDateSec = (startDate - DateSerial(1970, 1, 1)) * 86400
    EndDateSec = (endDate - DateSerial(1969, 12, 31)) * 86400
    
    ' Import data
    ' -------------------------------------------------------------

    With ActiveSheet.QueryTables.Add(Connection:= _
            "TEXT;https://query1.finance.yahoo.com/v7/finance/download/AAPL" & "?period1=" & StartDateSec & "&period2=" & EndDateSec & "&interval=1d&events=history" _
            , Destination:=Cells(1, 1))
            .Name = "table.csv?s=AAPL&d=4&e=10&f=2016&g=d&a=11&b=12&c=1980&ignore="
            .FieldNames = False
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlOverwriteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = False
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 850
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = True
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(4, 9, 9, 9, 9, 9, 9)
            .TextFileTrailingMinusNumbers = True
            'On Error Resume Next
            .Refresh BackgroundQuery:=False
        End With
        
    Sheets("Prices").Activate

End Function
Public Function Unprotect()
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    ws.Unprotect
Next

End Function

Public Function port_Lock()
    ThisWorkbook.Sheets("PortfolioOverall").Range("A:G").Locked = True
    Worksheets("PortfolioOverall").Protect
End Function

Public Function port_Unlock()
    Worksheets("PortfolioOverall").Unprotect
End Function

Public Function log_Lock()
    Worksheets("Log").Protect
End Function
Public Function log_UnLock()
    Worksheets("Log").Unprotect
End Function
Public Function view_UnLock()
    Worksheets("View").Unprotect
End Function

Public Function view_Lock()
    Worksheets("View").Protect
End Function

Public Function detail_Lock()
    Worksheets("Details").Protect
End Function
Public Function corr_Lock()
    Worksheets("Correlation").Protect
End Function
Public Function sec_Lock()
    Worksheets("Sectors").Protect
End Function

Public Function perf_Lock()
    Worksheets("Performance").Protect
End Function
