Attribute VB_Name = "BbDocTransformerTest"
Option Explicit
Option Private Module

' DocTransformerクラスを実行して動作を確認します、Rubberduckによる自動化テストではなく手動で実行する
' 入力としてのKeyValuePairを３つ与える、その３つはコードのなかに固定値として書いてある

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("DocTransformerクラスを味見する")
Private Sub TestDocTransformer()
    On Error GoTo TestFail
    
    Call BbLog.Clear
    
    ' 入力となるテンプレートとしてのWordファイルのパス
    Dim templateFile As String
    templateFile = BbUtil.ResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B2")
    
    ' 出力されるWordファイルを格納する先としてのフォルダのパス
    Dim resultFolder As String
    
    ' OneDrive固有の形式のファイルパスならば普通のファイルパスに変換する
    resultFolder = BbFile.ToLocalFilePath(ThisWorkbook.path) & "\" & "build\TestDocTransformer"
    
    ' 出力先フォルダがもしもまだ存在していなかったら作る
    Call BbFile.EnsureFolders(resultFolder)
    
    ' プレースホルダーの名前と値の組
    Dim dict As Dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "氏名", "ジョン・ドーエ"
    dict.Add "ID", "John Doe"
    dict.Add "PW", "ThisIsNotAPassword"
    
    ' 「氏名」に基づいて出力Wordファイルの名前を決める
    Dim resultName As String: resultName = dict("氏名") & ".docx"
        
    ' 入力Wordファイルと出力Wordファイルのパスを決定して
    Dim t As String: t = templateFile
    Dim r As String: r = resultFolder & "\" & resultName
        
    Debug.Print "テンプレート=" & t
    Debug.Print "出力ファイル=" & r
    
    ' Wordドキュメントを変換するクラスを初期化する
    ' DocTransformerインスタンスを生成する
    Dim dt As BbDocTransformer
    Set dt = BbDocTransformerFactory.CreateDocTransformer
    
    
    ' Wordアプリケーションのインスタンスを生成
    Dim wordApp As Word.Application: Set wordApp = CreateObject("Word.application")
    Call dt.Initialize(wordApp)
    ' 入力ファイルと置換データと出力ファイルを指定してWordドキュメントを変換し出力する
    Call dt.Transform(t, dict, r)
    
    Debug.Print "変換処理を実行しました"
    
    ' =========================================================================
    ' 後始末
    Set dt = Nothing
    Set dict = Nothing
    Set wordApp = Nothing

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


