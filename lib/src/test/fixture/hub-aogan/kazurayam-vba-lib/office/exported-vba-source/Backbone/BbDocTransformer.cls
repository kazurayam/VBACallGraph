VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BbDocTransformer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' BbDocTransformer Word文書変換機
'
' テンプレートとしてのWordファイルを読み、新しいWordファイルを生成して書き込む。
' Excelワークシートに書かれた名簿（人の名前とその属性）の１行ぶんのデータを入力として受け取る。
' テンプレートの中にあるplaceholders（たとえば ${名前}）をワークシートから取得した値で置換する。
'
' 親モジュールがExcelワークシートからたとえば2００人分の名簿を読み、
' DocTransformerを2００回呼び出せば、Wordファイルを2００個生成する処理を自動化することができる。
'
' WordファイルをIOするためにWord.Applicationオブジェクトを必要とするが、パフォーマンスを考慮して、
' 親モジュールがWord.Applicationオブジェクトを生成して、Initializeメソッドを呼び出して渡すことを求める。
'

Dim wordApp As Word.Application

Public Sub Initialize(wa As Word.Application)
    Set wordApp = wa
End Sub

Public Sub Transform( _
    ByVal TemplateWordFile As String, _
    ByVal KeyValuePairs As Dictionary, _
    ByVal ResultWordFile As String)
    
    Dim key As Variant
    
    ' テンプレートとしてのWordファイルを結果としてのWordファイルへコピー
    FileCopy TemplateWordFile, ResultWordFile
    
    ' 結果としてのWordファイルをopenする
    Dim doc As Word.Document: Set doc = wordApp.Documents.Open(ResultWordFile)
    
    ' Word文書のなかのプレースホルダーをExcelから取り出したデータで置き換える
    Dim KeysArray: KeysArray = KeyValuePairs.keys
    For Each key In KeysArray
        ' Debug.Print Key & "=" & KeyValuePairs.Item(Key)
        Dim placeholder As String: placeholder = "${" & key & "}"
        
        doc.Content.Find.Execute findText:=placeholder, _
                                ReplaceWith:=KeyValuePairs(key), _
                                MatchCase:=True, Replace:=wdReplaceAll
    Next
    
    ' 置換結果をファイルに保存する
    doc.Save
    
    ' 後始末
    doc.Close   ' Closeするのを忘れるな
                ' Word.DocumentオブジェクトをCloseを忘れると次回実行時にFieCopyでエラーコード71が発生して失敗する
    Set doc = Nothing
    
End Sub





