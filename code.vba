Sub main()
    Dim original As Worksheet
    Set original = Worksheets("original")
    
    '初期化
    Call init(original)
    
    '勝敗表を読み込み
    Dim result As Object
    Set result = getResult(original)
    
    '出力
    Dim actual As Worksheet
    Set actual = Worksheets("actual")
    Call output(actual, result)
End Sub


'初期化処理
Sub init(ByRef original As Worksheet)
    'シート削除時の警告抑制
    Application.DisplayAlerts = False
    
    '新規シート作成
    Sheets("actual").Delete
    Set actual = Worksheets.Add(After:=Worksheets("original"))
    actual.Name = "actual"
    
    '全ユーザ取得
    'ユニークにするためにDictionaryで作成
    Set users = CreateObject("Scripting.Dictionary")
    For i = 2 To original.Cells(Rows.Count, 1).End(xlUp).Row
        If original.Cells(i, 2) = "" Then
            '空のユーザ名の場合はループ終了
            Exit For
        End If
        Dim user1 As String: user1 = original.Cells(i, 2)
        Dim user2 As String: user2 = original.Cells(i, 5)
        
        If Not users.Exists(user1) Then
            users.Add user1, user1
        End If
        If Not users.Exists(user2) Then
            users.Add user2, user2
        End If
    Next i
    
    '取得したユーザを元にX軸とY軸にユーザ名を設定
    For i = 0 To users.Count - 1
        actual.Cells(i + 2, 1) = users.Keys()(i)
        actual.Cells(1, i + 2) = users.Keys()(i)
        actual.Cells(i + 2, i + 2) = "*"
    Next i
End Sub


'結果を取得する（シートを読み込んでDictionaryに変換）
Function getResult(ByRef original As Worksheet) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    For i = 2 To original.Cells(Rows.Count, 1).End(xlUp).Row
        If original.Cells(i, 2) = "" Then
            Exit For
        End If
        
        Dim user1 As String: user1 = original.Cells(i, 2)
        Dim victory1 As String: victory1 = original.Cells(i, 3)
        Dim user2 As String: user2 = original.Cells(i, 5)
        Dim victory2 As String: victory2 = original.Cells(i, 4)
        
        Call setGame(result, user1, victory1, user2, victory2)
        Call setGame(result, user2, victory2, user1, victory1)
    
    Next i
    Set getResult = result
End Function


'1ゲームの勝敗を設定する
Function setGame(ByRef result As Object, ByVal user1 As String, ByVal result1 As String, ByVal user2 As String, ByVal result2 As String)
    If Not result.Exists(user1) Then
        'ユーザが存在しない場合には作成
        result.Add user1, CreateObject("Scripting.Dictionary")
    End If
    
    Dim user As Object
    Set user = result(user1)
    
    If Not user.Exists(user2) Then
        '対戦相手が存在しない場合には設定
        user.Add user2, result1
    Else
        user(user2) = user(user2) & "," & result1
    End If


End Function


'シートに書き込む
Sub output(ByRef actual As Worksheet, ByRef result As Object)
    
    '最終列取得
    '何故かA1だと正しく最終列が取得できないためB1で取得
    Dim last_column As Integer
    last_column = actual.Range("B1").End(xlToRight).Column
    
    'ヘッダの範囲を取得
    Dim head As Range
    Set head = actual.Range(actual.Cells(1, 2), Cells(1, last_column))
    
    '結果をシートに記載する
    For Each user In result
        'ユーザを取得
        Dim user_index As Integer
        user_index = head.Find(What:=user, LookAt:=xlWhole).Column
        
        For Each opponent In result(user)
            'ユーザの対戦相手と勝敗を取得
            Dim opponent_index As Integer
            opponent_index = head.Find(What:=opponent, LookAt:=xlWhole).Column
            Dim victory As String
            victory = result(user)(opponent)
            
            '勝敗を記載
            actual.Cells(user_index, opponent_index).Value = victory
        Next opponent
    Next user
End Sub
