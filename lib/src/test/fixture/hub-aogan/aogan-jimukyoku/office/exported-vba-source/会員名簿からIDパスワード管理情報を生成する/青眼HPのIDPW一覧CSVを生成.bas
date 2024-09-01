Attribute VB_Name = "青眼HPのIDPW一覧CSVを生成"
Option Explicit


' 青眼HPの会員のページ https://www.aomori-gankaikai.jp/member がBasic認証によるアクセス制限を施している。
' その根拠となるIDとパスワードの組を収録したCSVファイルを会員名簿Excelから出力する。

Public Sub CSVを生成()

    ' イミディエイトウインドウをけす
    Call BbLog.Clear
    
    Debug.Print "青眼HPのIDとパスワードの組の一覧をCSVファイルに出力する"
    
    ' 会員名簿Excelファイルのパス
    Dim memberFile As String: memberFile = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B2")
    Debug.Print ("会員名簿: " & memberFile)
    
    
    ' 出力先フォルダのパス
    Dim outDir As String: outDir = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B5")
    ' CSVファイルの出力先パス
    Dim CSV As String: CSV = outDir & "\web_account.csv"
    Debug.Print "出力先: " + CSV
    
    ' 出力先フォルダがもしもまだ存在していなかったら作る
    Call BbFile.EnsureFolders(outDir)
    
    ' CSVテキストを出力するためにストリームを開く
    Dim fs As New FileSystemObject
    Dim ts As TextStream
    Set ts = fs.CreateTextFile(CSV, True, False)
    
    ' Webサイト管理者の認証情報を1行目と2行目に書く。
    ' なぜならWebサイト管理者は眼科医会の正会員ではないから会員名簿に含まれていないから、特例として。
    Dim plus1 As String: plus1 = "kazuaki001,a9ft5t72,""浦山和昭　事務局"""
    Call ts.Write(plus1)
    Call ts.Write(vbCrLf)
    
    Dim plus2 As String: plus2 = "aomoriken,gankaikai,""aomoriken / gankaikai"""
    Call ts.Write(plus2)
    Call ts.Write(vbCrLf)

    ' 外部にある会員名簿Excelファイルからシートをコピーして取り込み、
    ' その中にある会員名簿をListObjectとしてとりだす
    Dim memberTable As ListObject
    Set memberTable = MbMemberTableUtil.FetchMemberTable(memberFile, "R6年度", ThisWorkbook)
    Debug.Print "memberTable.ListRows.Count=" & memberTable.ListRows.count
    
    ' 会員名簿の行を処理する
    Dim max As Long: max = 300   ' テストするときにはmaxに小さい数字(3とか)をセットして実行時間を短縮する
                                    ' 本番には総会員数より大きい数字(300とか)をセットする
    Dim count As Long: count = 0
    
    Dim i As Long
    For i = 1 To memberTable.ListRows.count
        If i <= max Then
            ' 会員の氏名とIDとPWのデータを取り出す
            Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "氏名", Trim(memberTable.ListColumns("氏名").DataBodyRange(i))
            dict.Add "氏名カナ", Trim(memberTable.ListColumns("氏名カナ").DataBodyRange(i))
            dict.Add "HPのID", Trim(memberTable.ListColumns("HPのID").DataBodyRange(i))
            dict.Add "HPのパスワード", Trim(memberTable.ListColumns("HPのパスワード").DataBodyRange(i))
            '氏名漢字と氏名カナの２セルに字が書いてある行つまり名簿として有効な行を選ぶ
            If Not dict("氏名") = "" And Not dict("氏名カナ") = "" Then
                Dim line As String: line = dict("HPのID") & "," & dict("HPのパスワード") & ",""" & dict("氏名") & """"
                Debug.Print line
                'ファイルにWrite
                Call ts.Write(line)
                Call ts.Write(vbCrLf)
            End If
        End If
    Next i
    
    '出力先ファイルをクローズ
    Call ts.Close
    
    ' 処理完了を通知する、出力先を表示して
    Call MsgBox("出力先: " + CSV)
    
End Sub

