Attribute VB_Name = "会費納入のお願いを作成"
Option Explicit


' テンプレートであるWord文書を下敷きとして会員別にパーソナライズしたWord文書を生成する。

' ワークシート（R4年度）に書かれた会員名簿から会員個人の氏名と資格を読み取り、
' 一定条件で選別したうえで、テンプレート内のプレースホルダー（たとえば ${氏名}）を具体的な
' 文字に置換して、適切なファイル名を決定して、出力する。

Public Sub Main()

    ' イミディエイト・ウインドウを消す。
    ' 今回の実行でDebug.Printが出力するメッセージを見やすくするため。
    Call KzUtil.KzCls
    
    Debug.Print ("会費納入のお願いletterを作成します")
    
    ' 会員名簿Excelファイルのパス
    Dim memberFile As String: memberFile = _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B2")
    Debug.Print ("会員名簿: " & memberFile)
    
    ' テンプレートとしてのWordファイルのパス
    Dim templateFile As String: templateFile = _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B3")
    Debug.Print ("テンプレート: " & templateFile)
    
    ' 出力先フォルダのパス
    Dim outDir As String: outDir = KzFile.KzAbsolutifyPath( _
        ThisWorkbook.Path, _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B4"))
    Debug.Print ("出力フォルダ: " & outDir)

    ' 出力先フォルダがもしもまだ存在していなかったら作る
    Call KzFile.KzEnsureFolders(outDir)
    
    ' DocTransformerインスタンスを生成して
    Dim DT As DocTransformer: Set DT = DocTransformerUtil.Create
    ' Wordアプリケーションのインスタンスを与えて
    Dim WordApp As Word.Application: Set WordApp = CreateObject("Word.application")
    ' DocTrasnsformerを初期化する
    Call DT.Initialize(WordApp)
    
    '=================================================================================
    ' 外部にある会員名簿Excelファイルの[R6年度]シートをカレントのワークブックに
    ' コピーする。"work会員名簿"シートが作られる。その中身をListObjectとして取り出す
    '
    Dim memberTable As ListObject
    Set memberTable = AoMemberUtils.FetchMemberTable(memberFile, "R6年度", ThisWorkbook)
    Debug.Print "memberTable.ListRows.Count=" & memberTable.ListRows.count
    
    ' ================================================================================
    ' 会員名簿の各行を処理する
    Dim max As Long: max = 300     'テスト時には小さい数字(3とか)にして早く終了させる
                                 '本番には総会員数より大きい数字(300とか)にする
    Dim count As Long: count = 0
        
    Dim i As Long
    For i = 1 To memberTable.ListRows.count
        If i < max Then
            ' 会員の氏名、氏名カナ、資格を取り出す
            Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "氏名", Trim(memberTable.ListColumns("氏名").DataBodyRange(i))
            dict.Add "氏名カナ", Trim(memberTable.ListColumns("氏名カナ").DataBodyRange(i))
            dict.Add "資格", Trim(memberTable.ListColumns("資格").DataBodyRange(i))
            Debug.Print dict("氏名カナ"), dict("氏名"), dict("資格")
            
            ' A会員とB会員とC会員とD会員を対象とする。
            ' 免除会員はWord文書を生成しない。
            ' ”B弘大”は”B"と同じ、”C弘大”は”C”と同じとみなす
            If dict("資格") = "A" Or _
                StartsWith(dict("資格"), "B") Or _
                StartsWith(dict("資格"), "C") Or _
                dict("資格") = "D" Then
        
                Call dict.Add("資格short", Left(dict("資格"), 1))
                
                If dict("資格") Like "*弘大" Then
                    Call dict.Add("なお弘大", "なお弘前大学所属の先生方につきましては､教室の田澤さんに支払いを取りまとめていただきます｡どうぞご協力くださいますよう宜しくお願い申し上げます｡")
                Else
                    Call dict.Add("なお弘大", "")
                End If

            
                ' 出力Wordファイルのパスを決定して
                Dim r As String: r = outDir & "\" & dict("氏名カナ") & "_" & dict("氏名") & "_" & dict("資格") & ".docx"
                Debug.Print r
            
                ' Wordドキュメントを変換する処理を実行する
                Call DT.Transform(templateFile, dict, r)
    
            End If
        End If
    Next i
    
    Debug.Print "終了しました"
End Sub



Public Function StartsWith(target_str As String, search_str As String) As Boolean
'###################################################################################
'target_str文字列がsearch_str文字列で始まっているか確認する
'search_strで始まっている場合はTrue
'search_strで始まっていない、もしくはsearch_strがtarget_strの文字数を超える場合はFalseを返す
'
'例
'    StartsWith('C弘大', 'C') はTrueを返す
'    StartsWith('C弘大', 'E') はFalseを返す
'
'###################################################################################
  
  If Len(search_str) > Len(target_str) Then
    Exit Function
  End If
  
  If Left(target_str, Len(search_str)) = search_str Then
    StartsWith = True
  End If

End Function

