Attribute VB_Name = "Module5"
Sub Confirmation()

    If MsgBox("確認メッセージ シート内のデータが削除されます． 元に戻すことはできません．", vbOKCancel) = vbOK Then
        clear
    End If

End Sub
