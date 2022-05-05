Attribute VB_Name = "Module1"
Option Explicit
'this vba is to to help import stock data straight from yahoo finance

Sub ImportData()
    Dim enddate As Date
    Dim startdate As Date
    Dim stocksymbol As String
    Dim url As String
    Dim nQuery As Name
    Dim lastrow As Integer
    
    Dim Data As Worksheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    Sheets("XData").Cells.Clear
    
    Set Data = Sheets("Candlestick")
    
 'arrange positions for start and end dates along with stock symbol
 'clear sheet before input
 
        startdate = DataSheet.Range("startdate").Value
        enddate = DataSheet.Range("enddate").Value
        stocksymbol = DataSheet.Range("ticker").Value
        Sheets("Data").Range("A1").CurrentRegion.ClearContents
        
        url = "http://ichart.finance.yahoo.com/table.csv?s=" & stocksymbol & _
        "&a=" & Month(startdate) - 1 & "&b=" & Day(enddate) - 1 & "&e" & _
        Day(enddate) & "&f=" & Year(enddate) & "&g=" & Sheets("XData").Range("A1") & _
        "&q=q&y=0&z=" & stocksymbol & "&x=.csv"
 
QueryQuote:
     With Sheets("Data").QueryTables.Add(Connection:="URL;" & qurl, _
                 Destination:=Sheets("XData").Range("A1"))
                .TablesOnlyFromHTML = False
                .Refresh BackgroundQuery:=False
                .SaveData = True
                .BackgroundQuery = True
            
     End With
     
    Sheets("XData").Range("a1").CurrentRegion.TextToColumns Destination:=Sheets("xData").Range("a1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=True, Space:=False, other:=False
            
         Sheets("xData").Columns("A:G").ColumnWidth = 12

    lastrow = Sheets("xData").UsedRange.Row - 2 + Sheets("Data").UsedRange.Rows.Count
'SORT VALUES

    Sheets("XData").Sort.SortFields.Add Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Sheets("Data").Sort
        .SortMethod = xlPinYin
        .Apply
        .SortFields.Clear.SetRange Range("A1:G" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
    End With

End Sub


Sub Candlestick()
    Dim OHLCChart As ChartObject
    Dim nRows As Integer
    Dim ch As ChartObject

    nRows = Sheets("XData").UsedRange.Rows.Count

'Use for each loop to delete any charts there is


    For Each ch In Sheets("CandleChart").ChartObjects
        ch.Delete
    Next

    nRows = Sheets("Data").UsedRange.Rows.Count

'Create candlestick

    Set OHLCChart = Sheets("CandleChart").ChartObjects.Add(Left:=Range("b8").Left, Width:=400, Top:=Range("b8").Top, Height:=250)

    With OHLCChart.Chart
        .SetSourceData Source:=Sheets("Data").Range("a1:e" & nRows)
        .ChartType = xlStockOHLC
        .HasTitle = True
        .ChartTitle.Text = "Candlestick Chart for " & Sheets("CandleChart").Range("ticker")
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Price"
        .HasLegend = False
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(220, 230, 241)
        .ChartArea.Format.Line.Visible = msoFalse
        .Parent.Name = "OHLC Chart"
    End With

End Sub
    
