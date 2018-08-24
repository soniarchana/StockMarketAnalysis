Sub Compute_YrStckVol()

'Declaring Variables
Dim WorksheetName As String
Dim vtotvol As Variant
Dim yr_st_dt As String, yr_stnxt_dt As String, yr_en_dt As String
Dim i As Long, j As Long, LastRow As Long, LastORow As Long, k As Long
Dim openVal As Double, closeVal As Double
Dim GPer_Inc As Double, GPer_Dec As Double, GTotVol As Variant
Dim strRange As String, rng As Range


For Each ws In Worksheets
    
       'Set column headings in first row
       ws.Range("I1").Value = "Ticker"
       ws.Range("J1").Value = "Yearly Change"
       ws.Range("K1").Value = "Percent Change"
       ws.Range("L1").Value = "Total Stock Volume"
       
       ' Identify the Last Row of Sheet, Initialize total volume and j to 0
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        vtotvol = 0
        j = 2
        
       'Setting first time value for start date
       yr_st_dt = ws.Cells(2, "B").Value
      
      For i = 2 To LastRow
            
            If ws.Cells(i, "A").Value <> ws.Cells(i + 1, "A").Value Then
                
                'Setting Ticker and Computed Total Stock Volume
                vtotvol = vtotvol + ws.Cells(i, "G").Value
                ws.Cells(j, "I").Value = ws.Cells(i, "A").Value
                
                'Capture next start and current end date
                yr_stnxt_dt = ws.Cells(i + 1, "B").Value
                yr_en_dt = ws.Cells(i, "B").Value
                
                'Computing Yearly Change and Percent change
                If ws.Cells(i, "B").Value = yr_en_dt Then
                    closeVal = ws.Cells(i, "F").Value
                End If
                
                ws.Cells(j, "J").Value = closeVal - openVal
                ws.Cells(j, "L").Value = vtotvol
                
                'Setting Conditional Formatting with Red and Green colour
                If ws.Cells(j, "J").Value < 0 Then
                    ws.Cells(j, "J").Interior.ColorIndex = 3
                Else
                    ws.Cells(j, "J").Interior.ColorIndex = 4
                End If
                
                'Avoid divide zero error if any, though not required per current logic
                If closeVal <> 0 And openVal <> 0 Then
                    ws.Cells(j, "K").Value = (closeVal / openVal * 100) - 100
                Else
                    ws.Cells(j, "K").Value = 0
                End If
                
                'Resetting value of variables to intial values for new ticker
                j = j + 1
                vtotvol = 0
                closeVal = 0
                openVal = 0
                yr_st_dt = yr_stnxt_dt
                
            Else
                
                vtotvol = vtotvol + ws.Cells(i, "G").Value
                'Fetch Open Value of Stock Ticker for computing Yearly Change and Percent change
                If ws.Cells(i, "B").Value = yr_st_dt Then
                    openVal = ws.Cells(i, "C").Value
                End If
            
            End If
            
    Next i
    
    'Computer the Last Row of New Output Area
    LastORow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'Setting the header for printing summary
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    strRange = "K2:K" & LastORow
    Set rng = ws.Range(strRange)
    GPer_Inc = WorksheetFunction.Max(rng)
    ws.Range("Q2").Value = GPer_Inc
    GPer_Dec = WorksheetFunction.Min(rng)
    ws.Range("Q3").Value = GPer_Dec
      
    strRange = "L2:L" & LastORow
    Set rng = ws.Range(strRange)
    GTotVol = WorksheetFunction.Max(rng)
    ws.Range("Q4").Value = GTotVol
    
    'Pulling appropriate Ticker for the summary data
    For k = 2 To LastORow
        
        If ws.Cells(k, "K").Value = GPer_Inc Then
            ws.Range("P2").Value = ws.Cells(k, "I").Value
        End If
        If ws.Cells(k, "K").Value = GPer_Dec Then
            ws.Range("P3").Value = ws.Cells(k, "I").Value
        End If
        If ws.Cells(k, "L").Value = GTotVol Then
            ws.Range("P4").Value = ws.Cells(k, "I").Value
        End If
        
        
    Next k
    
Next ws

End Sub


