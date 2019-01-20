Attribute VB_Name = "Module8"
Sub test()
    Dim p As Long
    Dim o As Long
    Dim SheetName As String
    
    'ThisWorkbook.Activate
    Range(Cells(6, 16), Cells(36, 26)).ClearContents

    p = 0

    If ActiveSheet.ChartObjects.Count > 0 Then
        ActiveSheet.ChartObjects(1).Delete
    End If
    
    For i = 3 To Worksheets.Count - 3
        SheetName = Sheets(i).Name
        Sheets("団体的中表").Cells(3 + i - p, 16) = SheetName
    Next i


    For o = 0 To 9 '10人分
    For i = 3 To Worksheets.Count - 3
        For k = 0 To 46
            If Sheets("団体的中表").Cells(5, 17 + o) = Worksheets(i).Cells(4 + k, 1) Then
                Sheets("団体的中表").Cells(3 + i - p, 17 + o).Value = Worksheets(i).Cells(4 + k, 47)
                Exit For
            End If
        Next k
                If Sheets("団体的中表").Cells(3 + i - p, 17 + o) = "" Then
            'Sheets("団体的中表").Cells(3 + i - p, 16).Delete Shift:=xlShiftUp
            p = p + 1
        End If
    Next i
    Next o
    
    'Set Target = Union(Range(Range("Q5"), Cells(Rows.Count, 26).End(xlUp)), _
                       Range(Range("P5"), Cells(Rows.Count, 16).End(xlUp)))
    
    Set Target = Union(Range("P6", Range("P6").End(xlDown)), Range("AA6", Range("AA6").End(xlDown)))
        
    With ActiveSheet.Shapes.AddChart.Chart
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormatLocal = "0%"
        .Axes(xlValue).MajorUnit = 0.2
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).MinimumScale = 0
        .ChartType = xlLineMarkers
        .SetSourceData Source:=Target
    End With

    ActiveSheet.ChartObjects(1).Activate
    With ActiveChart.ChartArea
        .Top = Range("B6").Top
        .Left = Range("B6").Left
        .Width = 800
        .Height = 250
    End With

End Sub
