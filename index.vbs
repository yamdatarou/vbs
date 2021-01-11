Sub Hello()
    MsgBox "Hello!!!" 'モーダル表示
End Sub

Sub SettingValue()
    X = 1
    Y = 1
    MsgBox "X:" & X & "Y:" & Y
End Sub

Sub doArray()
    Dim list(4) 'サイズ5の配列を宣言
    list(0) = "a"
    MsgBox list(0)
End Sub

Sub doForEach()
    Dim list(5)
    Dim msg
    Dim i
    i = 0
    For Each tmp In list
        msg = msg & i
        i = i + 1
    Next
    MsgBox msg
End Sub

Call doForEach()
