Attribute VB_Name = "mdOrder"
Option Explicit

Dim TradeDate As Date
Dim TradeTime As Date
Dim Ticer As Integer
Dim LivePrice As Double
Dim Amount As Integer
Dim history, log, view, holdings As Worksheet


Sub buyInitialize()
    
    buyTickerForm.Show
    
End Sub


Sub sellInitialize()

    sellTickerForm.Show
    
End Sub


Sub BuyLog(Ticker, LivePrice)
Attribute BuyLog.VB_ProcData.VB_Invoke_Func = " \n14"
    log_UnLock

    Set view = ThisWorkbook.Worksheets("View")
    Set log = ThisWorkbook.Sheets("Log")
    
    Application.ScreenUpdating = False


    Dim Trades, HoldingsCount As Integer
    Dim LastRow As Integer
    Dim myRange, myLabel As String

    view.Activate
    
    TradeDate = Date
    TradeTime = time()
    
    Ticker = UCase(Ticker)
    
    Application.DisplayAlerts = False 'avoid system alert when input is empty
    Amount = Int(Application.InputBox(Prompt:="The current price for " & Ticker & " is USD " _
    & LivePrice & ", how many shares do you want to BUY? (Please input an integer)", Default:=0, Type:=1))
    
    HoldingsCount = WorksheetFunction.CountA(ThisWorkbook.Sheets("Holdings").Range("A:A")) - 1
    
    ' Trade restrictions
    ' -------------------------------------------------------------------------------------------------
    ' User can only buy, if he has enough cash
    If HoldingsCount > 15 Then
        MsgBox ("You cannot hold more than 15 different positions")
        Exit Sub
    
    ElseIf Amount = 0 Then
        MsgBox ("Please input a valid number larger than 0")
        Amount = Int(Application.InputBox(Prompt:="The current price for " & Ticker & " is USD " _
        & LivePrice & ", how many shares do you want to BUY? (Please input an integer)", Default:=0, Type:=1))
    
    ElseIf Amount * LivePrice > Range("wsCashNow") Then
        MsgBox ("You do not have enough funds.")
        Amount = Int(Application.InputBox(Prompt:="The current price for " & Ticker & " is USD " _
        & LivePrice & ", how many shares do you want to BUY? (Please input an integer)", Default:=0, Type:=1))
    
    Else
    Application.DisplayAlerts = True
    ' Paste stored values into log sheet
    ' -------------------------------------------------------------------------------------------------
    
        With log
            LastRow = WorksheetFunction.CountA(log.Range("$C:$C")) + 4
            .Cells(LastRow, 3).Value = TradeDate
            .Cells(LastRow, 3).NumberFormat = "dd/mm/yyyy"
            .Cells(LastRow, 4).Value = TradeTime
            .Cells(LastRow, 5).Value = Ticker
        
            If Amount > 0 Then
                .Cells(LastRow, 6).Value = "BUY"
            Else
                .Cells(LastRow, 6).Value = "SELL"
            End If
        
            .Cells(LastRow, 7).Value = Amount
            .Cells(LastRow, 8).Value = LivePrice
            .Cells(LastRow, 9).Value = Amount * LivePrice
            .Cells(LastRow, 2).Value = WorksheetFunction.Text(TradeDate, "DDMMYYYY") _
            & "_" & Ticker & "_" & .Cells(LastRow, 6).Value
            
            'Get the LastRow and choose the data series for chart
            myRange = "I5:I" & LastRow
            myLabel = "A5:A" & LastRow
            
            'Update Chart
            With .ChartObjects("Log Chart").Chart
        
                .SeriesCollection(1).Values = log.Range(myRange)
                .SeriesCollection(1).XValues = log.Range(myLabel)
        
            End With
        End With
    
    MsgBox ("You have bought " & Amount & " shares of " & Ticker & " succesfully!")
    
    End If
    
    Call HoldingsUpdate
    
    view.Select
    Calculate
    log_Lock
    Application.ScreenUpdating = True
    
End Sub

Sub SellLog(Ticker, LivePrice)

    Application.ScreenUpdating = False
    log_UnLock

    
    ' Define variables
    ' -------------------------------------------------------------------------------------------------
    
    Dim Trades, AmountHeld As Integer
    Dim LastRow As Long
    Dim myRange, myLabel As String
    Dim holdings As Variant
    Dim HoldingsCount, i, row As Integer
    Set log = ThisWorkbook.Worksheets("Log")
    Set view = ThisWorkbook.Worksheets("View")
     

    TradeDate = Date
    TradeTime = time()
    
    Ticker = UCase(Ticker)
    
    ' Show inputbox and save data to variables
    ' -------------------------------------------------------------------------------------------------

    Application.DisplayAlerts = False 'avoid system alert when input is empty
    Amount = Int(Application.InputBox(Prompt:="The current price for " & Ticker & " is USD " & LivePrice & ". You have " _
    & AmountHeld & " shares in total, how many shares do you want to SELL? (Please input an integer)", Default:=0, Type:=1))
    
    holdings = func_Holdings()
    HoldingsCount = UBound(holdings)
    
    For i = 1 To HoldingsCount
        If holdings(i) = Ticker Then
        row = i
        End If
    Next i
    
    ' Trade restrictions
    ' -----------------------------------------------------------------------------------------------------
    ' User can only sell more of a stock than he owns beforehand
    AmountHeld = ThisWorkbook.Worksheets("Holdings").Cells(row + 1, 2).Value
    If Amount = 0 Then
        MsgBox ("Please input a valid number larger than 0!")
        Amount = Int(Application.InputBox(Prompt:="The current price for " & Ticker & " is USD " & LivePrice & ". You have " _
        & AmountHeld & " shares in total, how many shares do you want to SELL? (Please input an integer)", Default:=0, Type:=1))
    End If
    
    If row = 0 Then
        MsgBox ("You cannot sell a stock, that you do not have.")
        Exit Sub
    ElseIf Amount > AmountHeld Then
        MsgBox ("You only hold " & AmountHeld & " stocks, which you can SELL.")
        Amount = Int(Application.InputBox(Prompt:="The current price for " & Ticker & " is USD " & LivePrice & ". You have " _
        & AmountHeld & " shares in total, how many shares do you want to SELL? (Please input an integer)", Default:=0, Type:=1))
    Else
    Application.DisplayAlerts = True
    
    
        ' Paste stored values into transactions sheet
        ' -------------------------------------------------------------------------------------------------
        
        Amount = Amount * -1
        With log
            LastRow = WorksheetFunction.CountA(log.Range("$C:$C")) + 4
            .Cells(LastRow, 3).Value = TradeDate
            .Cells(LastRow, 3).NumberFormat = "dd/mm/yyyy"
            .Cells(LastRow, 4).Value = TradeTime
            .Cells(LastRow, 5).Value = Ticker
        
            If Amount > 0 Then
                .Cells(LastRow, 6).Value = "BUY"
            Else
                .Cells(LastRow, 6).Value = "SELL"
            End If
        
            .Cells(LastRow, 7).Value = Amount
            .Cells(LastRow, 8).Value = LivePrice
            .Cells(LastRow, 9).Value = Amount * LivePrice
            .Cells(LastRow, 2).Value = WorksheetFunction.Text(TradeDate, "DDMMYYYY") _
            & "_" & Ticker & "_" & .Cells(LastRow, 6).Value
            
            'Get the LastRow and choose the data series for chart
            myRange = "I5:I" & LastRow
            myLabel = "A5:A" & LastRow
            
            'Update Chart
            With .ChartObjects("Log Chart").Chart
        
                .SeriesCollection(1).Values = log.Range(myRange)
                .SeriesCollection(1).XValues = log.Range(myLabel)
        
            End With
        End With
        
        MsgBox ("You have sold " & Amount * -1 & " shares of " & Ticker & " succesfully!")
        
    End If
    
    Call HoldingsUpdate
    
    view.Select
    Calculate
    log_Lock
    Application.ScreenUpdating = True
    
End Sub


Sub HoldingsUpdate()
Attribute HoldingsUpdate.VB_ProcData.VB_Invoke_Func = " \n14"
    
    ' Define arguments for sumifs function
    ' --------------------------------------------------------------------------------------------------
    
    Dim Arg1 As Range 'the range i want to sum
    Dim Arg2 As Range 'criteria range
    Dim HoldingsCount, i As Integer
    Dim tradeRange As String
    
    Set holdings = ThisWorkbook.Worksheets("Holdings")
    Set log = ThisWorkbook.Worksheets("Log")
    
    Dim LastRow As Long
    
    ' Copy tickers from transactions sheet and remove duplicates
    ' --------------------------------------------------------------------------------------------------
    
    With log
        LastRow = WorksheetFunction.CountA(ThisWorkbook.Worksheets("Log").Range("$C:$C")) + 3
    End With
    
    tradeRange = "E4:E" & LastRow
    
    holdings.Activate
    Cells.ClearContents
    
    log.Activate
    log.Range(tradeRange).Select
    
    Selection.Copy
    
    holdings.Activate
    holdings.Range("A1").Select
    
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
    
    
    ' Count rows and add number of stocks being held
    ' ---------------------------------------------------------------------------------------------------
    
    HoldingsCount = WorksheetFunction.CountA(Range("A:A")) - 1

    Set Arg1 = log.Range("G:G")
    Set Arg2 = log.Range("E:E")
    
    For i = 2 To HoldingsCount + 1
        Cells(i, 2) = Application.WorksheetFunction.SumIfs(Arg1, Arg2, holdings.Cells(i, 1))
    Next i
    
    
    ' Filter for and remove holdings that have 0 number of stocks
    ' ---------------------------------------------------------------------------------------------------
    
    Columns("A:B").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:B").AutoFilter Field:=2, Criteria1:="=0", Operator:=xlOr, Criteria2:="="
    
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    
    ActiveSheet.Range("$A$1:$B$5").AutoFilter Field:=2
    Range("A1").Select
    
    
    ' Sort holdings alphabetically
    ' --------------------------------------------------------------------------------------------------
    
    holdings.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
        
    With holdings.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Selection.AutoFilter
    
End Sub
