Sub main()

   Dim original As Worksheet
   Set original = Worksheets(1)

   '初期化
   Call init(original)

   '勝敗表を読み込む
   Dim result As Object
   Set result = getResult(original)

   '出力
   Dim actual As Worksheet
   Set actual = Worksheets(3)

   Call getResult(original)
   Call SetGame(result, user1, result1, user2, result2)

End Sub

'初期化処理
Sub init(ByRef original As Worksheet)

   Application.DisplayAlerts = False

  Set users = CreateObject("Scripting.Dictionary")

  For i = 2 To original.UsedRange.Rows.Count
     If original.Cells(i, 3) = "" Then
        Exit For
     End If

    Dim user1 As String: user1 = original.Cells(i, 3)
    Dim ressult1 As String: result1 = original.Cells(i, 4)
    Dim user2 As String: user2 = original.Cells(i, 7)
    Dim ressult2 As String: result2 = original.Cells(i, 6)

    Debug.Print (user1 & " " & result1 & " " & user2 & " " & result2)

     If Not users.Exists(user1) Then
          users.Add user1, user1
     End If
     If Not users.Exists(user2) Then
          users.Add user2, user2
     End If

 Next i

 Application.DisplayAlerts = False
    Sheets(3).Delete
    Set actual = Worksheets.Add(After:=Worksheets(2))
    actual.Name = "対戦表"

    Dim meibo As Worksheet

    Set meibo = Worksheets(2)

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

      Debug.Print (user1 & " " & result1 & " " & user2 & " " & result2)

      Call SetGame(result, user1, victory1, user2, victory2)
      Call SetGame(result, user2, victory2, user1, victory1)

    Next i

    Set getResult = result

End Function

'1ゲームの勝敗を設定する
 Function SetGame(ByRef result As Object, ByVal user1 As String, ByVal result1 As String, ByVal user2 As String, ByVal result2 As String)

       If Not result.Exists(user1) Then
          'ユーザが存在しない場合には設定
            result.Add user1, CreateObject("Scripting.Dictionary")
       End If

       Dim user As Object
       Set user = result(user1)

       If Not user.Exists(user2) Then
         '対戦相手が存在しない場合に設定
         user.Add user2, result1
       End If

 End Function

'シートに書き込む
Sub output(ByRef actual As Worksheet, ByRef result As Object, ByVal user_index As Integer)

 '最終列取得
 Dim Y As Integer
 Y = actual.Range("B1").End(xlToRight).Column

 Dim head As Range
 Set head = actual.Range(Cells(2, 1), Cells(1, Y))

 '結果をシートに記載しちゃうよ
 For Each user In result
        'ユーザを取得
     Dim user_index As Integer
     Set user_index = head.Find(What:=user, LookAt:=xlWhole).Column

     For Each opponent In result(user)
      'ユーザの対戦相手と勝敗を取得するよ
      Dim opponent_index As Integer
      opponent_index = original.Find(What:=opponent, LookAt:=xlWhole).Column
      Dim victory As String
      victory = result(user)(opponent)

      '勝敗を記載しちゃうよ
      actual.Cells(user_index, opponent_index).Value = victory
  Next opponent
Next user

End Sub
