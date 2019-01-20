Attribute VB_Name = "Module1"
Sub decision()
    
    Range(Cells(4, 42), Cells(50, 48)).ClearContents

    Dim t As Integer
    Dim s As Integer
    Dim f1 As Integer
    Dim f2 As Integer
    Dim f3 As Integer
    Dim f4 As Integer

    For h = 3 To Worksheets.Count - 2
        For i = 0 To 50 '人数
            For k = 0 To 39 '10立
                If IsEmpty(Worksheets(h).Cells(4 + i, 1)) = True Then
                    Exit For '名前が空白であれば次へ
                End If

                If Worksheets(h).Cells(4 + i, 2 + k) = "○" Then
                    t = t + 1 '○の数を集計
                End If

                If Worksheets(h).Cells(4 + i, 2 + k) = "○" Or Worksheets(h).Cells(4 + i, 2 + k) = "×" Then
                    s = s + 1 '引いた本数を集計
                End If
            Next k
            
            For p = 0 To 10 '各的中数
                If Worksheets(h).Cells(4 + i, 2 + 4 * p) = "○" Then
                    f1 = f1 + 1
                End If
                If Worksheets(h).Cells(4 + i, 3 + 4 * p) = "○" Then
                    f2 = f2 + 1
                End If
                If Worksheets(h).Cells(4 + i, 4 + 4 * p) = "○" Then
                    f3 = f3 + 1
                End If
                If Worksheets(h).Cells(4 + i, 5 + 4 * p) = "○" Then
                    f4 = f4 + 1
                End If
            Next p

            If Worksheets(h).Cells(4 + i, 1) <> "" And s <> 0 Then
                Worksheets(h).Cells(4 + i, 46) = t '的中数(立ち)
                Worksheets(h).Cells(4 + i, 47) = t / s '的中率
                Worksheets(h).Cells(4 + i, 48) = Application.RoundDown(s / 4, 0) '立ち回数
                
                '各的中数
                Worksheets(h).Cells(4 + i, 42) = f1
                Worksheets(h).Cells(4 + i, 43) = f2
                Worksheets(h).Cells(4 + i, 44) = f3
                Worksheets(h).Cells(4 + i, 45) = f4
           End If
        t = 0: s = 0: f1 = 0: f2 = 0: f3 = 0: f4 = 0 '初期化
        Next i
    Next h

End Sub
