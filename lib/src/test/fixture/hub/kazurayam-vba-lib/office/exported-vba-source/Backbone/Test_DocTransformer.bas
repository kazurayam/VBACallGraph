Attribute VB_Name = "Test_DocTransformer"
Option Explicit


' DocTransformerクラスを実行して動作を確認します、Rubberduckによる自動化テストではなく手動で実行する
' 入力としてのKeyValuePairを３つ与える、その３つはコードのなかに固定値として書いてある

Public Sub Test()
    ' イミディエイト・ウインドウを消す。
    ' 今回の実行でDebug.Printが出力するメッセージを見やすくするため。
    Call KzCls
    
    ' 入力となるテンプレートとしてのWordファイルが格納されたフォルダのパス
    Dim TemplateFolder As String: TemplateFolder = KzFile.KzToLocalFilePath(ThisWorkbook.path & "\" & "data")
    ' テンプレートとしてのWordファイルの名前
    Dim TemplateName As String: TemplateName = "テンプレート.docx"
    
    ' 出力されるWordファイルを格納する先としてのフォルダのパス
    Dim ResultFolder As String
    ' OneDrive固有の形式のファイルパスならば普通のファイルパスに変換する
    ResultFolder = KzFile.KzToLocalFilePath(ThisWorkbook.path) & "\" & "build\Test_DocTransformer"
    ' 出力先フォルダがもしもまだ存在していなかったら作る
    Call KzFile.KzEnsureFolders(ResultFolder)
    
    ' プレースホルダーの名前と値の組
    Dim dict As Dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "氏名", "織田信長"
    dict.Add "勤務先", "長良川クリニック"
    dict.Add "資格", "A"
    
    ' 「氏名」に基づいて出力Wordファイルの名前を決める
    Dim ResultName As String: ResultName = dict("氏名") & ".docx"
        
    ' 入力Wordファイルと出力Wordファイルのパスを決定して
    Dim t As String: t = TemplateFolder & "\" & TemplateName
    Dim r As String: r = ResultFolder & "\" & ResultName
        
    Debug.Print "テンプレート=" & t
    Debug.Print "出力ファイル=" & r
    
    ' Wordドキュメントを変換するクラスを初期化する
    ' DocTransformerインスタンスを生成する
    Dim DT As DocTransformer
    Set DT = DocTransformerUtil.Create
    
    
    ' Wordアプリケーションのインスタンスを生成
    Dim WordApp As Word.Application: Set WordApp = CreateObject("Word.application")
    Call DT.Initialize(WordApp)
    ' 入力ファイルと置換データと出力ファイルを指定してWordドキュメントを変換し出力する
    Call DT.Transform(t, dict, r)
    
    Debug.Print "変換処理を実行しました"
    
    ' =========================================================================
    ' 後始末
    Set DT = Nothing
    Set dict = Nothing
    Set WordApp = Nothing

End Sub


