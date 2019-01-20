Attribute VB_Name = "Module6"
Sub graqh()

    Dim SheetName As String

    If ActiveSheet.ChartObjects.Count > 0 Then
        ActiveSheet.ChartObjects(1).Delete
    End If

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
