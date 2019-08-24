Attribute VB_Name = "Module11"
Sub stockanalyst()

' Doing the easy and moderate part

    For Each ws In ThisWorkbook.Worksheets
    
        LastRow_report = Cells(Rows.Count, "A").End(xlUp).Row
        
        volumn = Cells(2, 7).Value
        
        Count = 1
        
        Startp = Cells(2, 3).Value
        
        closep = 0
    
        For i = 2 To LastRow_report
        
            If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
        
                volumn = volumn + Cells(i + 1, 7).Value
            
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                Count = Count + 1
        
    ' Get the yearly change and percent change, also fix empty cells issue
        
                closep = Cells(i, 6).Value
        
                If (Startp = 0) And (closep = 0) And (volumn = 0) Then
        
                    Cells(Count, 9).Value = Cells(i, 1).Value
                    Cells(Count, 12).Value = "N/A"
                    Cells(Count, 10).Value = "N/A"
                    Cells(Count, 11).Value = "N/A"
            
                Else
            
                    yearlychange = closep - Startp
                    changepct = yearlychange / Startp
            
    ' Plug those and the volume info into the sheet and adjust the variables
    
                    Cells(Count, 9).Value = Cells(i, 1).Value
                    Cells(Count, 12).Value = volumn
                    Cells(Count, 10).Value = yearlychange
                    Cells(Count, 11).Value = changepct
            
                End If
        
                volumn = Cells(i + 1, 7).Value
                Startp = Cells(i + 1, 3).Value
        
            End If
        
        Next i

        LastRow_results = Cells(Rows.Count, "J").End(xlUp).Row
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
    
        Range("K2:K" & LastRow_results).NumberFormat = "0.00%"
    
' Doing the hard part

        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
    
        Max = WorksheetFunction.Max(Range("K2:K" & LastRow_results))
        Cells(2, 17).Value = Max
    
        Min = WorksheetFunction.Min(Range("K2:K" & LastRow_results))
        Cells(3, 17).Value = Min
    
        Range("Q2:Q3").NumberFormat = "0.00%"
    
        Maxv = WorksheetFunction.Max(Range("L2:L" & LastRow_results))
        Cells(4, 17).Value = Maxv
    
        For i = 2 To LastRow_results
    
            If Cells(i, 11).Value = Max Then
            Cells(2, 16).Value = Cells(i, 9).Value
        
            ElseIf Cells(i, 11).Value = Min Then
            Cells(3, 16).Value = Cells(i, 9).Value
        
            ElseIf Cells(i, 12).Value = Maxv Then
            Cells(4, 16).Value = Cells(i, 9).Value
        
            End If
        
        Next i
        
    Next ws
    
End Sub

