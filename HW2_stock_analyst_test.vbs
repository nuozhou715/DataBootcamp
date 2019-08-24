Attribute VB_Name = "Module1"
Sub stockanalyst()

    Sheets.Add.Name = "Summary_Report"
    Sheets("Summary_Report").Move Before:=Sheets(1)

    Set combined = Worksheets("Summary_Report")

' Putting things together on one sheet

    For Each ws In Worksheets

        LastRow_combined = combined.Cells(Rows.Count, "A").End(xlUp).Row + 1
    
        LastRow_ws = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
    
        combined.Range("A" & LastRow_combined & ":G" & ((LastRow_ws - 1) + LastRow_combined)).Value = ws.Range("A2:G" & (LastRow_ws + 1)).Value
    
    Next ws

    combined.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value

' Doing the easy and moderate part

    LastRow_report = combined.Cells(Rows.Count, "A").End(xlUp).Row
    volumn = Cells(2, 7).Value
    Count = 1
    Startp = 41.81
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

    LastRow_results = combined.Cells(Rows.Count, "J").End(xlUp).Row
    
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
    
End Sub

