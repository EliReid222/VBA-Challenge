Attribute VB_Name = "Module1"
Sub Stocks()
    Dim ws As Worksheet
    For Each ws In Worksheets
    ws.Activate

  Dim ticker As String
  Dim Ticker_info As Double
  Dim NextTicker As String
  Dim PastTicker As String
  Dim Ticker_start As Double
  Dim Ticker_close As Double
  Dim Change As Double
  Dim Ticker_volume As Double
  Ticker_info = 2
  Ticker_volume = 0
  
  Cells(1, 10).Value = "Ticker"
  Cells(1, 11).Value = "Yearly Change"
  Cells(1, 12).Value = "Percent Change"
  Cells(1, 13).Value = "Total Stock Volume"
  
  Cells(2, 16).Value = "Greatest % Increase"
  Cells(3, 16).Value = "Greatest % Decrease"
  Cells(4, 16).Value = "Greatest Total Volume"
  Cells(1, 17).Value = "Ticker"
  Cells(1, 18).Value = "Value"
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
    For i = 2 To lastrow
  
    ticker = Cells(i, 1).Value
    NextTicker = Cells(i + 1, 1).Value
    PastTicker = Cells(i - 1, 1).Value
    Ticker_volume = Ticker_volume + Cells(i, 7).Value
    If NextTicker <> ticker Then
    
        Ticker_close = Cells(i, 6).Value
    
        Cells(Ticker_info, 10).Value = ticker
        
        Change = Ticker_close - Ticker_start
       
       Cells(Ticker_info, 11).Value = Change
       
       Cells(Ticker_info, 12).Value = (Change / Ticker_close) * 1
       
       Cells(Ticker_info, 12).Value = FormatPercent(Cells(Ticker_info, 12))
       
       Cells(Ticker_info, 13).Value = Ticker_volume
       
           If Cells(Ticker_info, 11).Value < 0 Then
            Cells(Ticker_info, 11).Interior.ColorIndex = 3
            Else
            Cells(Ticker_info, 11).Interior.ColorIndex = 4
    
        End If
        

       Ticker_info = Ticker_info + 1
       
       Ticker_volume = 0
       
       

    ElseIf PastTicker <> ticker Then
    
        Ticker_start = Cells(i, 3).Value
        

    End If
    
    Next i
    
        Greatest_Percentage_increase = WorksheetFunction.Max(Range("L2:L" & Ticker_info))
        Greatest_Percentage_decrease = WorksheetFunction.Min(Range("L2:L" & Ticker_info))
        Greatest_volume = WorksheetFunction.Max(Range("M2:M" & Ticker_info))
        Cells(2, 18).Value = Greatest_Percentage_increase
        Cells(3, 18).Value = Greatest_Percentage_decrease
        Cells(4, 18).Value = Greatest_volume
             
        Greatest_increase_row = WorksheetFunction.Match(Greatest_Percentage_increase, Range("L2: L" & Ticker_info), 0)
        Greatest_decrease_row = WorksheetFunction.Match(Greatest_Percentage_decrease, Range("L2:L" & Ticker_info), 0)
        Greatest_volume_row = WorksheetFunction.Match(Greatest_volume, Range("M2:M" & Ticker_info), 0)
        Cells(2, 17).Value = Cells(Greatest_increase_row + 1, 10)
        Cells(3, 17).Value = Cells(Greatest_decrease_row + 1, 10)
        Cells(4, 17).Value = Cells(Greatest_volume_row + 1, 10)

 Next ws
End Sub

