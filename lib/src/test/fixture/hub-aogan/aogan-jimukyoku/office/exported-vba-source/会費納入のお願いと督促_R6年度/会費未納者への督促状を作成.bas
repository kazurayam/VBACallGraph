Attribute VB_Name = "会費未納者への督促状を作成"
Option Explicit

'督促状を作成する
'work会員名簿ワークシートのなかのTableの会費納入状況が×と印づけられている会員を選択し、督促状を作成する。
'MakeLetterプロシジャは外部ファイルから会員名簿をロードするが、それと異なり、MakeReminderクラスは外部ファイルをロードしない。
'MakeReminderプロシジャは「会費納入状況チェック」モジュールが作成し更新した「work会員名簿」ワークシート」を参照する。
'MakeReminderプロシジャは「work会員名簿」ワークシートの「会費納入状況」列を調べ、そこに◎や×などの有効な文字が
'記入されていることをチェックする。もしも有効な文字が記入されていなければエラーメッセージを表示して終了する。


Public Sub MakeReminder()
    
    Call BbLog.Clear
    Call BbLog.Info("会費未納者への督促状を作成", "MakeReminder", "始めます")
    
    '督促状のテンプレートとしてのWordファイルのパス
    Dim templateFile As String: templateFile = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B6")
    Call BbLog.Info("会費未納者への督促状を作成", "MakeReminder", "テンプレート: " & templateFile)
    
    '出力先フォルダのパス
    Dim outDir As String: outDir = BbFile.AbsolutifyPath( _
        ThisWorkbook.Path, _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B7"))
    Call BbLog.Info("会費未納者への督促状を作成", "MakeReminder", "出力先フォルダ: " & outDir)
    
    ' 出力先フォルダがすでにあったら削除する
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(outDir) Then
        FSO.DeleteFolder outDir
    End If
    ' 出力先フォルダをもういちど作る
    Call BbFile.EnsureFolders(outDir)
    
    ' BbDocTransformerインスタンスを準備する
    Dim DT As BbDocTransformer: Set DT = BbDocTransformerFactory.CreateDocTransformer()
    ' Wordアプリケーションのインスタンスを与えて
    Dim WordApp As Word.Application: Set WordApp = CreateObject("Word.application")
    ' DocTrasnsformerを初期化する
    Call DT.Initialize(WordApp)
    
    '=================================================================================
    ' このワークブックに「work会員名簿」シートがあるはず。その中身をListObjectとして取り出す
    '
    Dim memberTable As ListObject
    Set memberTable = ThisWorkbook.Worksheets("work会員名簿").ListObjects("MembersTable13")
    
    Call BbLog.Info("会費未納者への督促状を作成", "MakeReminder", _
                            "memberTable.ListRows.Count=" & memberTable.ListRows.count)
    
    ' ================================================================================
    ' 会員名簿の各行を処理する
    Dim max As Long: max = 300     'テスト時には小さい数字(3とか)にして早く終了させる
                                 '本番には総会員数より大きい数字(300とか)にする
    Dim count As Long: count = 0

    Dim i As Long
    For i = 1 To memberTable.ListRows.count
        If i < max Then
            ' 会員の行から氏名、氏名カナ、資格、会費納入状況を取り出す
            Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "氏名", Trim(memberTable.ListColumns("氏名").DataBodyRange(i))
            dict.Add "氏名カナ", Trim(memberTable.ListColumns("氏名カナ").DataBodyRange(i))
            dict.Add "資格", Trim(memberTable.ListColumns("資格").DataBodyRange(i))
            dict.Add "会費納入状況", Trim(memberTable.ListColumns("会費納入状況").DataBodyRange(i))
            Call BbLog.Info("会費未納者への督促状を作成", "MakeReminder", dict("氏名カナ") & dict("会費納入状況"))
    
            
            ' 会費納入状況が未記入ならば（会費納入状況チェックがまだ実行されていないことを意味するので）
            ' Errを投げて終了する
            If dict("会費納入状況") = "" Then
                Err.Raise 5000, "会費未納者への督促状を作成.MakeReminder", _
                        dict("氏名カナ") & "の会費納入状況が未記入。" & _
                        "事前に会費納入状況チェックを実行する必要があります"
            End If
            
            ' 会費納入状況が×である会員を対象として督促状のWordファイルを作成する。
            ' 会費納入状況が×ではない会員については生成しない。
            If dict("会費納入状況") Like "×" Then
                Dim msg As String: Let msg = dict("氏名カナ") & " " & dict("氏名") & " " & dict("資格")
                Call BbLog.Info("会費未納者への督促状を作成", "MakeReminder", msg)

                ' 出力Wordファイルのパスを決定して
                Dim r As String: r = outDir & "\" & dict("氏名カナ") & "_" & dict("氏名") & "_" & dict("資格") & ".docx"
                Call BbLog.Info("会費未納者への督促状を作成", "MakeReminder", "出力ファイル " & r)
            
                ' Wordドキュメントを変換する処理を実行する
                Call DT.Transform(templateFile, dict, r)
    
            End If
        End If
    Next i
    
    Call BbLog.Info("会費未納者への督促状を作成", "MakeReminder", "終わりました")
    
End Sub
