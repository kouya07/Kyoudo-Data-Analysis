Attribute VB_Name = "Module7"
Sub pmtable()

  Workbooks.Open ("/Users/ono/Desktop/�l�I���\.xlsx")

  For i = 0 To 55
     If ThisWorkbook.Worksheets("�f�[�^").Cells(4 + i, 2) = "" Then
        GoTo Continue
     End If
     ThisWorkbook.Worksheets("�l�I���\").Range("F2") = ThisWorkbook.Worksheets("�f�[�^").Cells(4 + i, 2)
     individual
     Workbooks("�l�I���\.xlsx").Activate
     ThisWorkbook.Worksheets("�l�I���\").Copy After:=Worksheets(Worksheets.Count)
     ActiveSheet.Name = ThisWorkbook.Worksheets("�f�[�^").Cells(4 + i, 2)
     ActiveSheet.Range("F2") = ThisWorkbook.Worksheets("�f�[�^").Cells(4 + i, 2)

Continue:
   Next i

End Sub
