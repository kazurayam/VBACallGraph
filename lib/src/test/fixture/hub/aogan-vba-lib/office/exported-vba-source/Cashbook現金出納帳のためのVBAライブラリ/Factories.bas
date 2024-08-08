Attribute VB_Name = "Factories"
Option Explicit

' Factories --- Cashbookプロジェクトのクラスモジュールのうち外部プロジェクトがインスタンスを生成するために利用するFunctionたち

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
    Debug.Print "wbFullName=" & wb.FullName
    Debug.Print "sheetName=" & sheetName
    Debug.Print "tableId=" & tableId
    
    Dim ws As Worksheet: Set ws = wb.Worksheets(sheetName)
    Dim tbl As ListObject: Set tbl = ws.ListObjects(tableId)
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Set CreateCashbook = cb
End Function


Public Function CreateAccountsFinder(ByVal cb As Cashbook) As AccountsFinder
    '引数として渡されたCashbookを入力としてAccountsFinderオブジェクトを生成して返す
    '外部プロジェクトがAccountsFinderオブジェクトを生成するためにPublicなこの関数が必要だ
    Dim accFinder As AccountsFinder: Set accFinder = New AccountsFinder
    Call accFinder.Initialize(cb)
    Set CreateAccountsFinder = accFinder
End Function



Public Function CreateCashSelector(ByVal cb As Cashbook) As CashSelector
    '引数として渡されたCashbookを入力としてCashSelectorオブジェクトを生成して返す
    '外部プロジェクトがCashSelectorオブジェクトを生成するためにPublicなこの関数が必要だ
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
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
