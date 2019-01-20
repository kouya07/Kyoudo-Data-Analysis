Attribute VB_Name = "Module2"
Sub aggregate()

    Dim t As Integer
    Dim q As Integer
    Dim w As Integer
    Dim y As Integer
    Dim day As Integer
    Dim p1 As Integer
    Dim p2 As Integer

    'èâä˙âª
    Range(Cells(4, 3), Cells(55, 10)).ClearContents
    Range(Cells(4, 24), Cells(55, 26)).ClearContents
    Range(Cells(4, 12), Cells(55, 12)).Interior.ColorIndex = 0 ' îwåiêFÇæÇØÇÉNÉäÉA
    Range(Cells(4, 13), Cells(55, 13)).Interior.ColorIndex = 0 ' îwåiêFÇæÇØÇÉNÉäÉA

    For h = 3 To Worksheets.Count - 2
        For k = 0 To 50
            For j = 0 To 39
                If IsEmpty(Worksheets(h).Cells(4 + k, 1)) = True Then
                    Exit For
                'ElseIf Worksheets(h).Cells(4 + k, 2 + j) = "" Then
                 '   Exit For
                ElseIf Worksheets(h).Cells(4 + k, 2 + j) = "Åõ" Then
                    t = t + 1
                End If

                If (j + 1) Mod 4 = 0 And Worksheets(h).Cells(4 + k, 2 + j) <> "" Then
                    For o = 0 To 55 'èWåvÇÃêlêî
                    If Worksheets(h).Cells(4 + k, 1) = Cells(4 + o, 1) And Cells(4 + o, 1) <> "" Then
                        Select Case t
                            Case 0
                                Cells(4 + o, 10) = Cells(4 + o, 10) + 1
                                Exit For
                            Case 1
                                Cells(4 + o, 9) = Cells(4 + o, 9) + 1
                                Exit For
                            Case 2
                                Cells(4 + o, 8) = Cells(4 + o, 8) + 1
                                Exit For
                            Case 3
                                Cells(4 + o, 7) = Cells(4 + o, 7) + 1
                                If Worksheets(h).Cells(4 + k, 2 + j) = "Å~" Then
                                    q = q + 1 'ÉXÉPÉxêîÇâ¡éZ
                                End If
                                Exit For
                            Case 4
                                Cells(4 + o, 6) = Cells(4 + o, 6) + 1
                                w = w + 1 'äFíÜêîÇâ¡éZ
                                Exit For
                        End Select
                    End If
                    Next o
                t = 0
                End If
            Next j
            
            For l = 0 To 55
                If Worksheets(h).Cells(4 + k, 1) = Cells(4 + l, 1) And Cells(4 + l, 1) <> "" Then
                    If Worksheets(h).Name Like "*âŒ*" Or Worksheets(h).Name Like "*ìy*" Or Worksheets(h).Name Like "*ì˙*" Then
                        y = y + 1 'èoê»êîÇâ¡éZ
                    End If
                    Cells(4 + l, 3) = Cells(4 + l, 3) + Worksheets(h).Cells(4 + k, 48) * 4
                    Cells(4 + l, 4) = Cells(4 + l, 4) + Worksheets(h).Cells(4 + k, 46)
                    Cells(4 + l, 24) = Cells(4 + l, 24) + q: q = 0
                    Cells(4 + l, 25) = Cells(4 + l, 25) + w: w = 0
                    Cells(4 + l, 26) = Cells(4 + l, 26) + y: y = 0
                    Exit For
                End If
            Next l
        Next k
        day = day + 1 'ì˙êî
        Next h

    Cells(3, 28) = day: day = 0
    
        For r = 0 To 55
        If IsError(Cells(4 + r, 12)) = True Or IsError(Cells(5 + r, 12)) = True Then
            GoTo Continue
        End If
        
        If Cells(4 + r, 12) = Cells(5 + r, 12) And Cells(4 + r, 12) <> "" Then
            If Cells(3 + r, 12) = Cells(4 + r, 12) Then
                Cells(4 + r, 12).Interior.Color = Cells(3 + r, 12).Interior.Color ' îwåiêF
                Cells(4 + r, 13).Interior.Color = Cells(3 + r, 12).Interior.Color ' îwåiêF
                Cells(5 + r, 12).Interior.Color = Cells(3 + r, 12).Interior.Color ' îwåiêF
                Cells(5 + r, 13).Interior.Color = Cells(3 + r, 12).Interior.Color ' îwåiêF
                
            ElseIf Cells(3 + r, 12).Interior.Color <> RGB(231, 230, 230) Then
                Cells(4 + r, 12).Interior.Color = RGB(231, 230, 230) ' îwåiêF
                Cells(4 + r, 13).Interior.Color = RGB(231, 230, 230) ' îwåiêF
                Cells(5 + r, 12).Interior.Color = RGB(231, 230, 230) ' îwåiêF
                Cells(5 + r, 13).Interior.Color = RGB(231, 230, 230) ' îwåiêF
                
            ElseIf Cells(3 + r, 12).Interior.Color <> RGB(221, 235, 247) Then
                Cells(4 + r, 12).Interior.Color = RGB(221, 235, 247) ' îwåiêF
                Cells(4 + r, 13).Interior.Color = RGB(221, 235, 247) ' îwåiêF
                Cells(5 + r, 12).Interior.Color = RGB(221, 235, 247) ' îwåiêF
                Cells(5 + r, 13).Interior.Color = RGB(221, 235, 247) ' îwåiêF
            End If
        End If
        
Continue:
    Next r

End Sub
