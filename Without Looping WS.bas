Attribute VB_Name = "Module1"
Sub VBAHomework()
    
    'Define Variables
    
    Dim i As Long
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim TotalStockVolume2 As Double
    Dim FirstVol As Double
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim LastRows As Long
    Dim Summary_Table_Row As Double
    Dim EmptyRow As Double
    
    
    LastRows = Cells(Rows.Count, 1).End(xlUp).Row
    Summary_Row_Table = 2
    EmptyRow = 2
    FirstVol = Cells(2, 7).Value
    TotalStockVolume = 0
    YearlyOpen = Cells(2, 3).Value
    YearlyClose = Cells(2, 6).Value
    YearlyChange = YearlyClose - YearlyOpen
    PercentChange = YearlyChange / YearlyOpen
    
 
    
    'Headers
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O1").Value = "Greatest % Increase"
    Range("P1").Value = "Greatest % Decrease"
    Range("Q1").Value = "Greatest Total Volume"
    Cells.EntireColumn.AutoFit
    
    
        For i = 2 To LastRows
            
            Ticker = Cells(i, 1).Value
            
            'Conditional for columns that do match
    
            If Cells(i + 1, 1) = Cells(i, 1) Then
               
               Cells(i, 9).Value = Ticker
               
            'Calculates Yearly Change
    
               YearlyClose = Cells(i, 6).Value
               YearlyChange = YearlyClose - YearlyOpen
               Cells(i, 10).Value = YearlyChange
            
            End If
            
            
            'Calculates Percent Change
            
            If (YearlyOpen = 0 And YearlyClose = 0) Then
                PercentChange = 0
                Cells(i, 11).Value = PercentChange
            
            Else
                PercentChange = (YearlyChange / YearlyOpen) * 100
                Cells(i, 11).Value = PercentChange
                
            End If
               PercentChange = (YearlyChange / YearlyOpen) * 100
               Cells(i, 11).Value = PercentChange
               
               
            'Calculates Total Stock Volume
            If Cells(i + 1, 1) = Cells(i, 1) Then
               TotalStockVolume = Cells(i, 7).Value + TotalStockVolume
               Cells(i, 12).Value = TotalStockVolume
            
            Else
                Cells(EmptyRow, 9).Value = Ticker
                
                YearlyOpen = Cells(i + 1, 3).Value
                YearlyClose = Cells(i, 6).Value
                YearlyChange = YearlyClose - YearlyOpen
                Cells(EmptyRow, 10).Value = YearlyChange
                Cells(EmptyRow, 11).Value = PercentChange
                
                
                'Adding values in the next empty row
        
                TotalStockVolume2 = Cells(i, 7).Value
                TotalStockVolume = TotalStockVolume + TotalStockVolume2
                Cells(EmptyRow, 12).Value = TotalStockVolume
                TotalStockVolume = 0
                                           
                
                YearlyOpen = Cells(i + 1, 3).Value
                YearlyClose = Cells(i + 1, 6).Value
                
                Summary_Row_Table = Summary_Row_Table + 1
                EmptyRow = EmptyRow + 1
            
            End If
            
                'Conditional Cell Color
                
                If Cells(i, 10).Value > 0 Then
                    Cells(i, 10).Interior.ColorIndex = 4
                ElseIf Cells(i, 10).Value = 0 Then
                    Cells(i, 10).Interior.ColorIndex = 4
                Else
                    Cells(i, 10).Interior.ColorIndex = 3

         
            End If
            
        Next i

End Sub






