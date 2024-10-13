Sub NJ_Bill_Payroll_Diff()
    'This macro finds the visits that their visit hours exceed schedule hours more than 7 minutees
    
    'Set the variables
    Dim lastRow, i As Integer
    
    With ActiveWorkbook.Sheets(1)
        'Clean the data to make it easier to read
        .Range("1:2").EntireRow.Delete
        .Range("a:a, b:b, d:d, f:l, o:w, y:z, ab:ah").EntireColumn.Delete
        
        'Insert a column for the new data
        .Range("A:A").Insert
        Range("a1").Value = "Billing"
        Range("a1").Font.Bold = True
            
        'Find the last row
        lastRow = .Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        
        'Fill the column according to the condition
        For i = 2 To lastRow
            If Left(Range("D" & i).Value, 2) * 60 + Right(Range("D" & i).Value, 2) - Left(Range("F" & i).Value, 2) * 60 - Right(Range("F" & i).Value, 2) > 7 Then
                Range("A" & i).Value = "F"
            Else
                Range("A" & i).Value = "T"
            End If
        Next i
        
        'Reorder the columns
        .Columns("H:I").Cut
        .Columns("C").Insert shift:=xlToRight
        .Columns("G").Cut
        .Columns("I").Insert shift:=xlToRight
        
        'Format the headers
        .Range("A1:I1").Interior.Color = RGB(224, 224, 224)
        .Range("A1:I1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        'Format the appearance and apply a filter and show only ones need to be adjusted
        .Cells.WrapText = False
        .UsedRange.AutoFilter field:=1, Criteria1:="F"
        .Columns("A:I").AutoFit
    
    End With
    
    MsgBox ("Done!")
    
End Sub
