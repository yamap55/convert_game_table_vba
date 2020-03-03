Sub main()
    Dim original As Worksheet
    Set original = Worksheets("original")
   
    Dim meibo As Worksheet
    Set meibo = Worksheets("meibo")

    '初期化
    Call init(meibo)

    '勝敗表を読み込む
    Dim result As Object
    Set result = getResult(original)
    
    '使用するデータを選別する
    Set result = filter(result)

    '出力
    Dim actual As Worksheet
    Set actual = Worksheets("actual")
    Call output(actual, result)
End Sub

Function filter(ByRef result As Object) As Object
    'データから使用しないデータを取り除く
    For Each user In result
        For Each opponent In result(user)
            Dim tmp As Variant
            tmp = Split(result(user)(opponent), ",")
            If 0 >= UBound(tmp) - LBound(tmp) Then
                '2戦以上の対戦がない場合は記録を消す
                result(user).Remove opponent
            End If
        Next opponent
    Next user
    Set filter = result
End Function

'シートの存在チェック
Function existsSheet(ByVal sheetName As String) As Boolean
    existsSheet = False
    
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = sheetName Then
            'シートが存在していた場合にはTrueを返す
            existsSheet = True
            Exit Function
        End If
    Next ws
End Function

'初期化処理
Sub init(ByRef meibo As Worksheet)
    If existsSheet("actual") Then
        '既にactualシートが存在していたら削除
        Application.DisplayAlerts = False
        Sheets("actual").Delete
    End If
    Set actual = Worksheets.Add(After:=Worksheets("original"))
    actual.Name = "actual"
    
    For i = 2 To meibo.UsedRange.Rows.Count
        actual.Cells(i, 1) = meibo.Cells(i, 2)
        actual.Cells(1, i) = meibo.Cells(i, 2)
        actual.Cells(i, i) = "*"
    Next
End Sub

Function getResult(ByRef original As Worksheet) As Object

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    X = original.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To X
        If original.Cells(i, 3) = "" Then
            Exit For
        End If

        Dim user1 As String: user1 = original.Cells(i, 3)
        Dim victory1 As String: victory1 = original.Cells(i, 4)
        Dim user2 As String: user2 = original.Cells(i, 7)
        Dim victory2 As String: victory2 = original.Cells(i, 6)

'        Debug.Print (user1 & " " & victory1 & " " & user2 & " " & victory2)
        
        Call setGame(result, user1, victory1, user2, victory2)
        Call setGame(result, user2, victory2, user1, victory1)
    Next i

    Set getResult = result

End Function

'1ゲームの勝敗を設定する
Function setGame(ByRef result As Object, ByVal user1 As String, ByVal result1 As String, ByVal user2 As String, ByVal result2 As String)
    If Not result.Exists(user1) Then
        'ユーザが存在しない場合には設定
        result.Add user1, CreateObject("Scripting.Dictionary")
    End If

    Dim user As Object
    Set user = result(user1)

    If Not user.Exists(user2) Then
        '対戦相手が存在しない場合に設定
        user.Add user2, result1
    Else
        user(user2) = user(user2) & "," & result1
    End If
End Function

'シートに書き込む
Sub output(ByRef actual As Worksheet, ByRef result As Object)
    
    '最終列取得
    Dim Y As Integer
    Y = actual.Range("B1").End(xlToRight).Column

    Dim head As Range
    Set head = actual.Range(actual.Cells(1, 2), actual.Cells(1, Y))

    '結果をシートに記載しちゃうよ
    For Each user In result
        'ユーザを取得
        Dim user_index As Integer
        user_index = head.Find(What:=user, LookAt:=xlWhole).Column

        For Each opponent In result(user)
            'ユーザの対戦相手と勝敗を取得するよ
            Dim opponent_index As Integer
            opponent_index = head.Find(What:=opponent, LookAt:=xlWhole).Column
            Dim victory As String
            victory = result(user)(opponent)

            '勝敗を記載しちゃうよ
            actual.Cells(user_index, opponent_index).Value = victory
        Next opponent
    Next user
End Sub

