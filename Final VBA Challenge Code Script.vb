Sub StockData():
'Set Current worksheet as Variable
    Dim CWS As Worksheet
    For Each CWS In Worksheets

'set Headers
    Cells(1, 9).Value = "Ticker Symbol"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
'set Headers for Challenge
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    

'Define Variables for Calculations
    Dim i As Long
    Dim TickerName As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double

'Set Variables equal to 0
    YearOpen = 0
    YearClose = 0
    YearlyChange = 0
    PercentChange = 0
    TotalStockVolume = 0

'Define Variables for Challenge
    Dim MaxTicker As String
    Dim MinTicker As String
    Dim MaxVolumeTicker As String
    Dim MaxPercent As Double
    Dim MinPercent As Double
    Dim MaxVolume As Double
    
'Set Challenge variables equal to 0
    MaxPercent = 0
    MinPercent = 0
    MaxVolume = 0

'Name starting row bc data is on row 2
    Dim Table As Long
    Table = 2

'Define variable to the last row
    Dim LastRow As Long
    LastRow = CWS.Cells(Rows.Count, 1).End(xlUp).Row

'Define Open Price & set variable i range
    YearOpen = CWS.Cells(2, 3).Value
    For i = 2 To LastRow
    
'Finding Ticker Symbols, start for loop
        If CWS.Cells(i + 1, 1).Value <> CWS.Cells(i, 1).Value Then
        TickerName = CWS.Cells(i, 1).Value
        
'Yearly Change math
    YearClose = CWS.Cells(i, 6).Value
    YearlyChange = YearClose - YearOpen
        
'Percent Change math
        If YearOpen <> 0 Then
        PercentChange = (YearlyChange / YearOpen) * 100
        
End If

'Stock Volume math
        TotalStockVolume = TotalStockVolume + CWS.Cells(i, 7).Value
            
'log math into Table cells
        CWS.Range("I" & Table).Value = TickerName
        CWS.Range("J" & Table).Value = YearlyChange
        CWS.Range("K" & Table).Value = (CStr(PercentChange) & "%")
        CWS.Range("L" & Table).Value = TotalStockVolume
       

'formatting Font and Color Format for YearlyChange with conditional formatting
        If (YearlyChange > 0) Then
        'Fill column as Green for + change
        CWS.Range("J" & Table).Interior.ColorIndex = 4
        'Check if Yearly Change is less than 0
        ElseIf (YearlyChange <= 0) Then
        'Fill column with Red for - change
        CWS.Range("J" & Table).Interior.ColorIndex = 3
End If
           
'Log next iterator of ticker
    Table = Table + 1
        
'Reset the values for the next ticker
    YearlyChange = 0
    YearOpen = CWS.Cells(i + 1, 3).Value
    YearClose = 0
            
'Challenge calculations
    If (PercentChange > MaxPercent) Then
    MaxPercent = PercentChange
    MaxTicker = TickerName
    CWS.Range("O2").Value = MaxTicker
    CWS.Range("P2").Value = (CStr(MaxPercent) & "%")
        
    ElseIf (PercentChange < MinPercent) Then
    MinPercent = PercentChange
    MinTicker = TickerName
    CWS.Range("O3").Value = MinTicker
    CWS.Range("P3").Value = (CStr(MinPercent) & "%")
        
End If
        
        If (TotalStockVolume > MaxVolume) Then
        MaxVolume = TotalStockVolume
        MaxVolumeTicker = TickerName
        CWS.Range("O4").Value = MaxVolumeTicker
        CWS.Range("P4").Value = MaxVolume
        
End If
            
'Reset the values for next ticker iteration
            PercentChange = 0
            TotalStockVolume = 0
            
        Else
            TotalStockVolume = TotalStockVolume + CWS.Cells(i, 7).Value
        End If
    Next i
    
'New integar j for used columns & auto fit formula
    Dim j As Integer
    For j = 1 To CWS.UsedRange.Columns.Count
    CWS.Columns(j).EntireColumn.AutoFit
    Next j
    
Next CWS
End Sub

