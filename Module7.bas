Attribute VB_Name = "Module7"
Sub pmtable()

  Workbooks.Open ("/Users/ono/Desktop/個人的中表.xlsx")

  For i = 0 To 55
     If ThisWorkbook.Worksheets("データ").Cells(4 + i, 2) = "" Then
        GoTo Continue
     End If
     ThisWorkbook.Worksheets("個人的中表").Range("F2") = ThisWorkbook.Worksheets("データ").Cells(4 + i, 2)
     individual
     Workbooks("個人的中表.xlsx").Activate
     ThisWorkbook.Worksheets("個人的中表").Copy After:=Worksheets(Worksheets.Count)
     ActiveSheet.Name = ThisWorkbook.Worksheets("データ").Cells(4 + i, 2)
     ActiveSheet.Range("F2") = ThisWorkbook.Worksheets("データ").Cells(4 + i, 2)

Continue:
   Next i

End Sub
