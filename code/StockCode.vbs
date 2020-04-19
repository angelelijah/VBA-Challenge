Attribute VB_Name = "Module1"
Sub StockAssignment()

'loop through all sheets
For Each WS In Worksheets


    'variables extracted data
    Dim Ticker As String
    Dim OpenP As Single
    Dim CloseP As Single
    Dim TotalSV As LongLong

    'variables calculated from data
    Dim YearlyChange As Single
    Dim PercentChange As Double

    'titles for output tracking
    WS.Cells(1, 9) = "Ticker"
    WS.Cells(1, 10) = "Yearly Change"
    WS.Cells(1, 11) = "Percent Change"
    WS.Cells(1, 12) = "Total Stock Volume"

    'variable to keep track of row for output
    Dim Count As Long
    Count = 2

    ' set value for starting variables
    OpenP = 0
    TotalSV = 0
    
    'last row tracker variables
    Dim LastRow As Long
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

    'loop through all rows
    Dim i As Long
    
    For i = 2 To LastRow
    
        'Set name for ticker
        Ticker = WS.Cells(i, 1)
    
        'statement to see if next row ticker is NOT the same as current
        If WS.Cells(i, 1).Value <> WS.Cells(i + 1, 1).Value Then
        
            'set value for ending Variables
            CloseP = WS.Cells(i, 6).Value
        
            'calculate values for ending variables
            YearlyChange = (CloseP - OpenP)
        
                'divide by zero fix & calculate percent change
                If OpenP = 0 Then
                PercentChange = 0
                Else
                PercentChange = ((CloseP - OpenP) / OpenP)
                End If
        
            'input desired data into sumary boxes
            WS.Cells(Count, 9).Value = Ticker
            WS.Cells(Count, 10).Value = YearlyChange
            WS.Cells(Count, 11).Value = PercentChange
                'Change PercentChange Cell into percent formatting
                WS.Cells(Count, 11) = Format(WS.Cells(Count, 11).Value, "0.00%")
                'Color cells
                    If PercentChange > 0 Then
                    WS.Cells(Count, 11).Interior.ColorIndex = 3
                    ElseIf PercentChange < 0 Then
                    WS.Cells(Count, 11).Interior.ColorIndex = 4
                    End If
            WS.Cells(Count, 12).Value = TotalSV
            'move down next row
            Count = (Count + 1)
        
            'Reset Starting Variables to 0
            OpenP = 0
            TotalSV = 0
        
        'statement to see if next row ticker is the same as current
        ElseIf WS.Cells(i, 1).Value = WS.Cells(i + 1, 1).Value Then
    
            'reset value for next variable on next stock
            If OpenP = 0 Then
            OpenP = WS.Cells(i, 3)
            End If
            
            If TotalSV = 0 Then
            TotalSV = WS.Cells(i, 7)
            End If
        
        'adding up totalSV over time
        TotalSV = TotalSV + WS.Cells(i + 1, 7).Value
     
        End If
    Next i
        
    '----------------------------------------------------------
    'Challenge portion
    '----------------------------------------------------------
    
    'name desired cells accordingly for summary table
    WS.Cells(2, 15) = "Greatest % increase"
    WS.Cells(3, 15) = "Greatest % decrease"
    WS.Cells(4, 15) = "Greatest Total Volume"
    WS.Cells(1, 16) = "Ticker"
    WS.Cells(1, 17) = "Value"
    
    'create variables
    Dim MaxIncrease As Single
    Dim IncTicker As String
    
    Dim MaxDecrease As Single
    Dim DecTicker As String
    
    Dim MaxTV As LongLong
    Dim TVTicker As String
    
    'Set starting values to variables
    MaxIncrease = WS.Cells(2, 10)
    MaxDecrease = WS.Cells(2, 10)
    MaxTV = WS.Cells(2, 12)
    
    'Loop through yearly change & Total Stock Volume
    Dim j As Long
    
    For j = 2 To LastRow
    
        'Find Max Increase
        If MaxIncrease < WS.Cells(j, 10) Then
            MaxIncrease = WS.Cells(j, 10)
            IncTicker = WS.Cells(j, 9)
        End If
        
        'Find Max Decrease
        If MaxDecrease > WS.Cells(j, 10) Then
            MaxDecrease = WS.Cells(j, 10)
            DecTicker = WS.Cells(j, 9)
        End If
        
        'Find Max Total Volume
        If MaxTV < WS.Cells(j, 12) Then
            MaxTV = WS.Cells(j, 12)
            TVTicker = WS.Cells(j, 9)
        End If
        
    Next j
    ' input correct ticker into summary
    WS.Cells(2, 16).Value = IncTicker
    WS.Cells(3, 16).Value = DecTicker
    WS.Cells(4, 16).Value = TVTicker
    
    
    'input correct value into summary
    WS.Cells(2, 17).Value = MaxIncrease
    WS.Cells(3, 17).Value = MaxDecrease
    WS.Cells(4, 17).Value = MaxTV
    
        
Next WS
        
End Sub

