Attribute VB_Name = "CbFactories"
Option Explicit

' CbFactories --- Cashbookプロジェクトのクラスモジュールのうち外部プロジェクトがインスタンスを生成するために利用するFunctionたち

' ここで宣言されたFunctionをユニットテストするコードがTestFactoriesモジュールにある

' このプロジェクトをAddinつまり*.xlamファイルにして、外部プロジェクトが ツール > 参照設定 > 参照 で*.xlamファイルを選択したとき
' 外部プロジェクトは　New Cashbookすることができない。
' というのもVBAのクラスはPublicNonCreatableな属性を持っているので、外部プロジェクトのコードはNew Cashbookすることができない。
' そこで*.xlamのなかにPublicなファクトリ関数を準備しておく。外部プロジェクトはPublicなファクトリ関数ならば呼び出すことが
' 許される。


Public Function CreateCashbook(ByVal wb As Workbook, _
                                ByVal sheetName As String, _
                                ByVal tableId As String) As Cashbook
    '引数として渡されたWorkbookを入力としてCashbookオブジェクトを生成して返す
    '外部プロジェクトがCashbookオブジェクトを生成するためにPublicなこの関数が必要だ
    Call BbLog.Info("Factories", "CreateCashbook", "wbFullName=" & wb.FullName)
    Call BbLog.Info("Factories", "CreateCashbook", "sheetName=" & sheetName)
    Call BbLog.Info("Factories", "CreateCashbook", "tableId=" & tableId)
    
    Dim ws As Worksheet: Set ws = wb.Worksheets(sheetName)
    Dim tbl As ListObject: Set tbl = ws.ListObjects(tableId)
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Set CreateCashbook = cb
End Function


Public Function CreateCashbookTransformer(ByVal cb As Cashbook) As CashbookTransformer
    '引数として渡されたCashbookを入力としてCashbookTransformerオブジェクトを生成して返す
    '外部プロジェクトがCashbookTransformerオブジェクトを生成するためにPublicなこの関数が必要だ
    Dim cbTransformer As CashbookTransformer: Set cbTransformer = New CashbookTransformer
    Call cbTransformer.Initialize(cb)
    Set CreateCashbookTransformer = cbTransformer
End Function



Public Function CreateCashSelector(ByVal cb As Cashbook, _
        Optional ByVal periodStart As Date = #4/1/2022#, _
        Optional ByVal periodEnd As Date = #3/31/2023#) As CashSelector
    '引数として渡されたCashbookを入力としてCashSelectorオブジェクトを生成して返す
    '外部プロジェクトがCashSelectorオブジェクトを生成するためにPublicなこの関数が必要だ
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb, periodStart, periodEnd)
    Set CreateCashSelector = cs
End Function


Public Function CreateEmptyCashList() As CashList
    '長さゼロのCashListオブジェクトを返す
    'CashListオブジェクトをPrivateと宣言したから外部モジュールがNew CashList()できなくなった。
    'それを補完するため。
    Dim cl As CashList
    Set cl = New CashList
    Set CreateEmptyCashList = cl
End Function
