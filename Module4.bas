Attribute VB_Name = "Module4"
Sub individual()

    Dim s As Long
    Dim p As Long
    Dim f1 As Long
    Dim f2 As Long
    Dim f3 As Long
    Dim f4 As Long

    Dim SheetName As String
    'ThisWorkbook.Activate
    Range(Cells(6, 16), Cells(36, 17)).ClearContents

    p = 0

    If ActiveSheet.ChartObjects.Count > 0 Then
        ActiveSheet.ChartObjects(1).Delete
    End If

    For i = 3 To Worksheets.Count - 2
        SheetName = Sheets(i).Name
        Sheets("個人的中表").Cells(3 + i - p, 16) = SheetName
        For k = 0 To 46
            If Sheets("個人的中表").Cells(2, 7).MergeArea(1, 1).Value = Worksheets(i).Cells(4 + k, 1) Then
                Sheets("個人的中表").Cells(3 + i - p, 17).Value = Worksheets(i).Cells(4 + k, 47)
                s = s + Worksheets(i).Cells(4 + k, 48)
                f1 = f1 + Worksheets(i).Cells(4 + k, 42)
                f2 = f2 + Worksheets(i).Cells(4 + k, 43)
                f3 = f3 + Worksheets(i).Cells(4 + k, 44)
                f4 = f4 + Worksheets(i).Cells(4 + k, 45)
                Exit For
            End If
        Next k
        If Sheets("個人的中表").Cells(3 + i - p, 17) = "" Then
            Sheets("個人的中表").Cells(3 + i - p, 16).Delete Shift:=xlShiftUp
            p = p + 1
        End If
    Next i

    Sheets("個人的中表").Cells(4, 9) = s '立ち数
    Sheets("個人的中表").Cells(20, 3) = f1 / s
    Sheets("個人的中表").Cells(20, 5) = f2 / s
    Sheets("個人的中表").Cells(20, 7) = f3 / s
    Sheets("個人的中表").Cells(20, 9) = f4 / s

    With ActiveSheet.Shapes.AddChart.Chart
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormatLocal = "0%"
        .Axes(xlValue).MajorUnit = 0.2
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).MinimumScale = 0
        .ChartType = xlLineMarkers
        .SetSourceData Range("P6", Range("Q6").End(xlDown))
    End With

    ActiveSheet.ChartObjects(1).Activate
    With ActiveChart.ChartArea
        .Top = Range("B6").Top
        .Left = Range("B6").Left
        .Width = 800
        .Height = 250
    End With

End Sub
