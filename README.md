Private Sub TextBox1_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Initialize()
    LoadQuestion
End Sub

Private Sub LoadQuestion()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("問題")

    If ws.Cells(questionIndex, 1).Value = "" Then
        MsgBox "問題が見つかりません。"
        Unload Me
        Exit Sub
    End If

    LabelQuestion.Caption = ws.Cells(questionIndex, 1).Value

    ' 各選択肢の取得
    Dim choices(1 To 3) As String
    Dim i As Integer
    For i = 1 To 3
        choices(i) = ws.Cells(questionIndex, i + 1).Value
    Next i

    ' オプションボタンに反映
    OptionButton1.Caption = choices(1)
    OptionButton1.Value = False
    OptionButton1.Visible = (choices(1) <> "")

    OptionButton2.Caption = choices(2)
    OptionButton2.Value = False
    OptionButton2.Visible = (choices(2) <> "")

    OptionButton3.Caption = choices(3)
    OptionButton3.Value = False
    OptionButton3.Visible = (choices(3) <> "")
End Sub

Private Sub btnAnswer_Click()
    Dim selectedAnswer As String
    If OptionButton1.Value Then
        selectedAnswer = OptionButton1.Caption
    ElseIf OptionButton2.Value Then
        selectedAnswer = OptionButton2.Caption
    ElseIf OptionButton3.Value Then
        selectedAnswer = OptionButton3.Caption
    Else
        MsgBox "回答を選んでください"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("問題")
    Dim correctAnswer As String
    correctAnswer = ws.Cells(questionIndex, 5).Value

    Dim resultWs As Worksheet
    Set resultWs = ThisWorkbook.Sheets("結果ログ")
    Dim logRow As Long
    logRow = resultWs.Cells(resultWs.Rows.Count, 1).End(xlUp).Row + 1
    If logRow < 2 Then logRow = 2

    ' 時間はあとからまとめて入力
    resultWs.Cells(logRow, 1).Value = userName
    resultWs.Cells(logRow, 2).Value = "" ' ← 後で終了時間を入れる
    resultWs.Cells(logRow, 3).Value = ws.Cells(questionIndex, 1).Value ' 問題文
    resultWs.Cells(logRow, 4).Value = selectedAnswer
    resultWs.Cells(logRow, 5).Value = correctAnswer

    If selectedAnswer = correctAnswer Then
        correctCount = correctCount + 1
    Else
        resultWs.Range(resultWs.Cells(logRow, 1), resultWs.Cells(logRow, 5)).Interior.Color = RGB(255, 255, 0)
    End If

    questionIndex = questionIndex + 1

    If ws.Cells(questionIndex, 1).Value = "" Then
        ' 最終問題が終わったタイミングで終了時間取得
        Dim endTime As String
        endTime = Format(Now, "yyyy/m/d h:mm")

        ' 同じユーザーで時間未記入の行に終了時間を一括入力
        Dim i As Long
        For i = resultWs.Cells(resultWs.Rows.Count, 1).End(xlUp).Row To 2 Step -1
            If resultWs.Cells(i, 1).Value = userName And resultWs.Cells(i, 2).Value = "" Then
                resultWs.Cells(i, 2).Value = endTime
            ElseIf resultWs.Cells(i, 1).Value <> userName Then
                Exit For
            End If
        Next i

        ' 正答率シートに記録
        Dim summaryWs As Worksheet
        Set summaryWs = ThisWorkbook.Sheets("正答率")
        Dim summaryRow As Long
        summaryRow = summaryWs.Cells(summaryWs.Rows.Count, 1).End(xlUp).Row + 1
        If summaryRow < 2 Then summaryRow = 2

        summaryWs.Cells(summaryRow, 1).Value = userName
        summaryWs.Cells(summaryRow, 2).Value = endTime ' ← 終了時間で記録
        summaryWs.Cells(summaryRow, 3).Value = correctCount
        summaryWs.Cells(summaryRow, 4).Value = questionIndex - 2
        summaryWs.Cells(summaryRow, 5).Value = Format(correctCount / (questionIndex - 2), "0.00%")

        MsgBox "回答お疲れ様でした。結果画面を表示します", vbOKOnly
        Call ShowResultToSheet(userName)
        Unload Me
        Exit Sub
    End If

    LoadQuestion
End Sub
