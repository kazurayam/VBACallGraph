Attribute VB_Name = "IDPW通知するWord文書を生成"
Option Explicit

' 本Subは、青森県眼科医会の会員一人一人宛に郵送するWord文書を生成します。
' 「あなたがＨＰの会員ページにサインインするのに必要なＩＤとパスワードはこれです」という文書。
' 本Subは外部にある青森県眼科医会会員名簿のExcelファイルを読み込む。
' 名簿には各会員の氏名とＩＤとパスワードが書いてある。
' テンプレートとしてのWordファイルを読み込み、プレースホルダーとしての記述（${氏名} など）を
' Excelから拾ったデータで置換する。これを会員人数分繰り返して、人数分のWordファイルを出力する。

Public Sub IDPW通知文書を生成()

    ' イミディエイト・ウインドウを消す
    Call KzCls
    
    Debug.Print ("ID/PWを通知するWord文書を生成します")
    
    ' 会員名簿Excelファイルのパス
    Dim memberFile As String: memberFile = _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B2")
    Debug.Print ("会員名簿: " & memberFile)
    
    ' テンプレートとしてのWordファイルのパス
    Dim templateFile As String: templateFile = _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B3")
    Debug.Print ("テンプレート: " & templateFile)
    
    ' 出力フォルダのパス
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
    ' DocTransformerを初期化する
    Call DT.Initialize(WordApp)

    ' 外部にある会員名簿Excelファイルからシートをコピーして取り込み、その中にある会員名簿をListObjectとしてとりだす
    Dim memberTable As ListObject
    Set memberTable = AoMemberUtils.FetchMemberTable(memberFile, "R6年度", ThisWorkbook)
    Debug.Print "memberTable.ListRows.Count=" & memberTable.ListRows.count
    
    
    ' 会員名簿の行を処理する
    Dim max As Long: max = 300  ' テストするときにはmaxを小さい数字(3とか)にし、本番には総会員数より大きい数字にする
    Dim count As Long: count = 0
    
    Dim i As Long
    For i = 1 To memberTable.ListRows.count
        If i < max Then
            ' 会員の氏名とIDとPWのデータを取り出す
            Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "氏名", Trim(memberTable.ListColumns("氏名").DataBodyRange(i))
            dict.Add "氏名カナ", Trim(memberTable.ListColumns("氏名カナ").DataBodyRange(i))
            dict.Add "ID", Trim(memberTable.ListColumns("HPのID").DataBodyRange(i))
            dict.Add "PW", Trim(memberTable.ListColumns("HPのパスワード").DataBodyRange(i))
            
            '氏名漢字と氏名カナの２セルに字が書いてある行つまり名簿として有効な行を選ぶ
            If Not dict("氏名") = "" And Not dict("氏名カナ") = "" Then
                
                Debug.Print (dict("氏名") & " " & dict("氏名カナ") & " " & dict("ID") & " " & dict("PW"))
                
                ' 新しく作るWordファイルの名前を決める
                Dim r As String: r = outDir & "\" & "IDPW_" & dict("氏名カナ") & ".docx"
                Debug.Print r
                
                ' Wordドキュメントを変換する処理を実行する
                Call DT.Transform(templateFile, dict, r)
            
            End If
        End If
    Next i
    
    ' Wordアプリケーションを閉じる
    WordApp.Quit
    Set WordApp = Nothing
    
    Debug.Print "終了しました。"
    MsgBox "出力先: " & outDir
    
End Sub

