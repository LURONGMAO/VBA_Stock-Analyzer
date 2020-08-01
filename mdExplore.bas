Attribute VB_Name = "mdExplore"
Option Explicit

    Public CurrentPrice, ChangeRatio, ChangePrice, OpenPrice, PreviousClose As String
    Public Request As MSXML2.ServerXMLHTTP60
    Public StockTime

    Dim Ticker, Min_Day, Max_Day, Min_52w, Max_52w As String
    Dim Volume, MarketCap, EPS, PE_Ratio, ShareCapital As String
    
    Dim startDate, endDate As Date
    Dim StartDateSec, EndDateSec As Double
    
    Dim conn As Variant
    Dim LastRow As Integer
    Dim history, view As Worksheet


Sub ExploreStock()
    Unprotect
    
    Application.ScreenUpdating = False
    
    ChooseTicker
    GetHistory
    StockInfo (Ticker)
    StoreTicker
    UpdateView
    Chart
    MsgBox ("Data sucessfully retrieved!")
    
    If ThisWorkbook.Worksheets("view").Range("close_open") = "At Close" Then
        log_Lock
        perf_Lock
        detail_Lock
        port_Lock
        corr_Lock
        sec_Lock
        Application.ScreenUpdating = True
        Exit Sub
    Else
        log_Lock
        perf_Lock
        detail_Lock
        port_Lock
        corr_Lock
        sec_Lock
        Timer
    End If
    
    Application.ScreenUpdating = True
    
End Sub


Sub frmInitialize()

    frmGetData.Show
    
End Sub


Sub ChooseTicker()
    ' Define start and end day of price history
    
    startDate = DateSerial(Year(frmGetData.startDate.Text), Month(frmGetData.startDate.Text), Day(frmGetData.startDate.Text))
    endDate = DateSerial(Year(frmGetData.endDate.Text), Month(frmGetData.endDate.Text), Day(frmGetData.endDate.Text))
    
    ' Get the right Ticker
    Ticker = Split(frmGetData.tbTicker.Text)(0)
    
    
    ' Convert date into seconds for API query
    StartDateSec = (startDate - DateSerial(1970, 1, 1)) * 86400
    EndDateSec = (endDate - DateSerial(1969, 12, 31)) * 86400
    
End Sub


Sub GetHistory()

    ' API QUERY
    ' --------------------------------------------------------------------------
    
    Dim RequestString As String
    Dim HistoryRequest As MSXML2.ServerXMLHTTP60
    Dim HistoryResponse As Variant
    

    Set history = ThisWorkbook.Worksheets("myHistory")
    
    ' Clear previous record
    history.UsedRange.Clear
    
    ' Query url
    RequestString = "https://query1.finance.yahoo.com/v7/finance/download/" & Ticker & "?period1=" & StartDateSec & "&period2=" & EndDateSec & "&interval=1d&events=history"
    
    ' Create ServerXMLHTTP60 object to connect to the Internet
    Set HistoryRequest = New ServerXMLHTTP60
    
    With HistoryRequest
        .Open "GET", RequestString, False, ","
        .send
        'Store the queried data in string format
        HistoryResponse = .responseText
    End With
    
    Set HistoryRequest = Nothing


    ' CSV WRITER
    ' --------------------------------------------------------------------------
    
    Dim nColumns As Integer
    Dim csv_row() As String
    Dim csv_file() As String
    Dim csv_range, closeRange As String
    Dim iRows As Integer
    Dim cell
    
    nColumns = 6
    
    'Split the data into array
    csv_row() = Split(HistoryResponse, Chr(10))
    
    'Assign array value to cells
    For iRows = 0 To UBound(csv_row)
        csv_range = "A" & iRows + 1 & ":F" & iRows + 1
        csv_file = Split(csv_row(iRows), ",")
        history.Range(csv_range).Value = csv_file
    Next iRows
    
    'Get the last row of the data
    With history
       .UsedRange.Calculate
        LastRow = .Cells.SpecialCells(xlCellTypeLastCell).row
    End With
    
    'Covert string to number
    closeRange = "F2:F" & LastRow
    For Each cell In history.Range(closeRange).Cells
        cell.Value = Val(cell.Value)
    Next

End Sub


Sub StockInfo(StockCode As String)

    ' API QUERY
    ' --------------------------------------------------------------------------
    
    Dim Url, RequestString, ReqType As String
    Dim Response As String
    Dim StockData() As String
    
    'Query url
    Url = "http://hq.sinajs.cn/list=gb_"
    
    'Ensure ticker format align with api requirement
    StockCode = LCase(StockCode)
    If InStr(StockCode, "-") > 0 Then StockCode = Replace(StockCode, "-", "$")
    
    RequestString = Url & StockCode
    
    'Create ServerXMLHTTP60 object to connect to the Internet
    Set Request = New ServerXMLHTTP60
    
    With Request
        .Open "GET", RequestString, False, ","
        .send
        'Store data in string format
        Response = .responseText
    End With
    
    Set Request = Nothing
    
    ' STORING VALUES IN ARRAY
    ' ---------------------------------------------------------------------------
    
    StockData = Split(Response, ",")
    
    ' Assign data to each variable
    CurrentPrice = StockData(1)
    ChangeRatio = StockData(2)
    ChangePrice = StockData(4)
    OpenPrice = StockData(5)
    PreviousClose = StockData(26)
    Min_Day = StockData(6)
    Max_Day = StockData(7)
    Min_52w = StockData(9)
    Max_52w = StockData(8)
    Volume = StockData(10)
    MarketCap = StockData(12)
    EPS = StockData(13)
    PE_Ratio = StockData(14)
    ShareCapital = StockData(19)
    StockTime = StockData(25)
    
    ' Get the time for live price
    StockTime = Split(StockTime, " EDT")(0)
    
End Sub

Sub StoreTicker()
    
    Set view = ThisWorkbook.Worksheets("View")
    Dim intRow As Long
    intRow = 1
    
    ' Store Ticker to cells so that the procedure can get Ticker from cell when refreshing
    ' Update Ticker Name from TickerList
    With view
        .Range("wsTickerName").Value = Application.VLookup(Ticker, ThisWorkbook.Worksheets("TickerList").Range("A2:B483"), 2, False)
        .Range("wsTicker").Value = Ticker
    End With

End Sub


Sub UpdateView()
    view_UnLock
    ' Assign variable values to cells
    With ThisWorkbook.Worksheets("View")
        .Range("wsLivePrice").Value = CurrentPrice
        .Range("wsLivePrice").NumberFormat = "[$$-en-US]#,##0.00" 'Fixed the displaying format
        .Range("wsPriceChange").Value = ChangePrice
        .Range("wsChangeRatio").Value = ChangeRatio
        .Range("wsVolume").Value = Volume
        .Range("wsOpenPrice").Value = OpenPrice
        .Range("wsClosePrice").Value = PreviousClose
        .Range("wsMinPrice").Value = Min_Day
        .Range("wsMaxPrice").Value = Max_Day
        .Range("wsMinYear").Value = Min_52w
        .Range("wsMaxYear").Value = Max_52w
        .Range("wsEPS").Value = EPS
        .Range("wsPE").Value = PE_Ratio
        .Range("wsMktCap").Value = MarketCap
        .Range("wsShareCap").Value = ShareCapital
        
        ' Check if the market close or open
        If DateDiff("n", StockTime, .Range("EDT_NOW")) > 5 Then
            .Range("close_open") = "At Close"
            .Range("close_open").Font.Color = RGB(165, 165, 165)
        Else
            .Range("close_open") = "At Open"
            .Range("close_open").Font.Color = RGB(45, 199, 126)
        End If
        
        
    End With
    
    view_Lock
End Sub

Sub Chart()
    view_UnLock
    Dim LastRow As Integer
    Dim myRange, myLabel As String      ' set the data series range for charts
    
    ' Get the last row of history records
    With ThisWorkbook.Worksheets("myHistory")
        LastRow = .Cells.SpecialCells(xlCellTypeLastCell).row
    End With
    
    myRange = "F2:F" & LastRow
    myLabel = "A2:A" & LastRow
        
    ' Assign new data series to chart and update the chart in View Page
    With ThisWorkbook.Sheets("View").ChartObjects("Stock Price").Chart
        .SeriesCollection(1).Values = ThisWorkbook.Worksheets("myHistory").Range(myRange)
        .SeriesCollection(1).XValues = ThisWorkbook.Worksheets("myHistory").Range(myLabel)
    End With
    
    view_Lock

End Sub

Sub Refresh()
    
    ' Refresh LivePrice
    Dim refreshCode As String
    view_UnLock
    refreshCode = ThisWorkbook.Worksheets("View").Range("wsTicker").Value
    StockInfo (refreshCode)
    UpdateView
    view_Lock
    Timer
    
End Sub
