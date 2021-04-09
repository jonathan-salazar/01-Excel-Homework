Sub wallSt()

' Set CurrentWs
    Dim CurrentWs As Worksheet
    
    ' Loop through all worksheets in workbook
    For Each CurrentWs In Worksheets
    
        ' Set ticker name
        Dim Ticker As String
        Ticker = " "
        
        ' Set total stock volume
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        
        ' Set variables for calculations
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Delta_Price As Double
        Delta_Price = 0
        Dim Delta_Percent As Double
        Delta_Percent = 0

        ' Set variables for Bonus calculations
        Dim MAX_TICKER_NAME As String
        MAX_TICKER_NAME = " "
        Dim MIN_TICKER_NAME As String
        MIN_TICKER_NAME = " "
        Dim MAX_PERCENT As Double
        MAX_PERCENT = 0
        Dim MIN_PERCENT As Double
        MIN_PERCENT = 0
        Dim MAX_VOLUME_TICKER As String
        MAX_VOLUME_TICKER = " "
        Dim MAX_VOLUME As Double
        MAX_VOLUME = 0
        

        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        ' Set row count for CurrentWs
        Dim Lastrow As Long
        Dim i As Long
        
        'Get last row
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        
            ' Set column names
            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"
            
            ' Set additional titles
            CurrentWs.Range("O2").Value = "Greatest % Increase"
            CurrentWs.Range("O3").Value = "Greatest % Decrease"
            CurrentWs.Range("O4").Value = "Greatest Total Volume"
            CurrentWs.Range("P1").Value = "Ticker"
            CurrentWs.Range("Q1").Value = "Value"
        
        'Set open price
        Open_Price = CurrentWs.Cells(2, 3).Value
        
        ' Loop from the beginning of the current worksheet(Row2) till its last row
        For i = 2 To Lastrow
      
            ' Check if we are still within the same ticker name
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
                ' Set the ticker name
                Ticker = CurrentWs.Cells(i, 1).Value
                
                ' Calculate Delta_Price and Delta_Percent
                Close_Price = CurrentWs.Cells(i, 6).Value
                Delta_Price = Close_Price - Open_Price
                ' Check Division by 0 condition
                If Open_Price <> 0 Then
                    Delta_Percent = (Delta_Price / Open_Price) * 100
                Else
                    MsgBox ("Some open price values are zero.")
                End If
                
                ' Add to the Ticker name total volume
                Total_Stock_Volume = Total_Stock_Volume + CurrentWs.Cells(i, 7).Value
              
                'Set values
                CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker
                CurrentWs.Range("J" & Summary_Table_Row).Value = Delta_Price

                'Set colors
                If (Delta_Price > 0) Then
                    'Fill column with GREEN
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Delta_Price <= 0) Then
                    'Fill column with RED
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")
                CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                ' Add 1 to row count
                Summary_Table_Row = Summary_Table_Row + 1

                ' Reset
                Delta_Price = 0
                Close_Price = 0
                Open_Price = CurrentWs.Cells(i + 1, 3).Value
              
                'BONUS
                If (Delta_Percent > MAX_PERCENT) Then
                    MAX_PERCENT = Delta_Percent
                    MAX_TICKER_NAME = Ticker
                ElseIf (Delta_Percent < MIN_PERCENT) Then
                    MIN_PERCENT = Delta_Percent
                    MIN_TICKER_NAME = Ticker
                End If
                       
                If (Total_Stock_Volume > MAX_VOLUME) Then
                    MAX_VOLUME = Total_Stock_Volume
                    MAX_VOLUME_TICKER = Ticker
                End If
                
                'Reset
                Delta_Percent = 0
                Total_Stock_Volume = 0
                
            
            'Else add to Total Stock Volume
            Else
                Total_Stock_Volume = Total_Stock_Volume + CurrentWs.Cells(i, 7).Value
            End If

        Next i
            
                CurrentWs.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
                CurrentWs.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
                CurrentWs.Range("P2").Value = MAX_TICKER_NAME
                CurrentWs.Range("P3").Value = MIN_TICKER_NAME
                CurrentWs.Range("Q4").Value = MAX_VOLUME
                CurrentWs.Range("P4").Value = MAX_VOLUME_TICKER
        
     Next CurrentWs
End Sub
