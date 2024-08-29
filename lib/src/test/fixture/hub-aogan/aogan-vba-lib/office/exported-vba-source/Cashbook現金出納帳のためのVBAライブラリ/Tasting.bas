Attribute VB_Name = "Tasting"
                                                                                                                                                                                                                                                Option Explicit

' Tasting : 味見する、気軽に触ってみる、ちょっと試す、いろいろと

Sub TasteCashbook()
    ' Cashbookクラスを味見する
    ' Cashbookオブジェクトを生成して、Countプロパティを読みだしてPrintする
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    'Act:
    Dim cb As Cashbook: Set cb = CreateCashbook(wb)
    'Assert:
    Debug.Print "cb.Count=" & cb.Count
    'TearDown
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb.Close
End Sub


Sub TasteCash_ColumnHeader()
    ' CashオブジェクトのColumnHeader()メソッドを味見する
    Dim cs As Cash
    Set cs = New Cash
    Debug.Print cs.ColumnHeader()
End Sub


Sub TasteCash_ToString()
    ' CashオブジェクトのToString()メソッドを味見する
    Call KzCls
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("現金出納帳")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    '
    Dim data As Range: Set data = tbl.ListRows(12).Range '収入の一例
    Dim cs As Cash: Set cs = New Cash
    Call cs.Initialize(data)
    Debug.Print cs.ToString()
    '
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb.Close
End Sub


Sub TasteCash_ToDate()
    ' CashオブジェクトのToDate()メソッドを味見する
    Call KzCls
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("現金出納帳")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    '
    Dim data As Range: Set data = tbl.ListRows(12).Range '収入の一例
    Dim cs As Cash: Set cs = New Cash
    Call cs.Initialize(data)
    Debug.Print cs.ToDate()
    'TearDown
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb.Close
End Sub

Sub TasteCashSelector_SelectCashList_ofIncome()
    ' CashSelectorオブジェクトのSelectCashListメソッドを味見する
    ' Incomeの場合
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("現金出納帳")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    '
    Call KzCls
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "雑収入", "セミナー参加料", "眼科フォーラム")
    Debug.Print selected.Count
    Debug.Print selected.ToString()
    'TearDown
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb.Close
End Sub


Sub TasteCashSelector_SelectCashList_ofExpense()
    ' CashSelectorオブジェクトのSelectCashListOfIncomeメソッドを味見する
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("現金出納帳")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    '
    Call KzCls
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Expense, "事業費", "広報費", "会報")
    
    Debug.Print selected.Count
    Debug.Print selected.ToString()
    'TearDown
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb.Close
End Sub


Sub TasteAccount()
    Call KzCls
    ' Expense account
    Dim accExpense As Account: Set accExpense = New Account
    Call accExpense.Initialize(AccountType.Expense, "事業費", "広報費")
    Debug.Print "accExpense.accType: " & accExpense.AccType
    Debug.Print "accExpense.AccountName: " & accExpense.AccountName
    Debug.Print "accExpense.SubAccountName: " & accExpense.SubAccountName
    Debug.Print "accExpense.ToString(): " & accExpense.ToString()
    ' Income account
    Dim accIncome As Account: Set accIncome = New Account
    Call accIncome.Initialize(AccountType.Income, "雑収入", "セミナー参加料")
    Debug.Print "accIncome.accType: " & accIncome.AccType
    Debug.Print "accIncome.AccountName: " & accIncome.AccountName
    Debug.Print "accIncome.SubAccountName: " & accIncome.SubAccountName
    Debug.Print "accIncome.ToString(): " & accIncome.ToString()
End Sub


Sub TasteAccountsFinder_allUnit()
    ' AccountsFinderクラスを味見する
    'Arrange
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("現金出納帳")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act
    Dim accFinder As AccountsFinder: Set accFinder = New AccountsFinder
    Call accFinder.Initialize(cb)
    Dim dic As Dictionary
    Set dic = accFinder.FindAccounts()
    'Assert
    Call KzCls
    Dim key As Variant
    For Each key In dic
        Debug.Print key & ":" & dic(key)
    Next
'0//:1
'収入/会費/A会員:44
'収入/会費/B会員:32
'収入/会費/C会員:1
'収入/補助金/補助金:1
'収入/雑収入/セミナー参加料:4
'収入/雑収入/広告料収入:12
'収入/雑収入/雑収入:1
'支出/事務費/消耗品費:27
'支出/事務費/通信費:58
'支出/事業費/公衆衛生費:1
'支出/事業費/学術費:17
'支出/事業費/広報費:3
'支出/事業費/庶務費:1
'支出/事業費/眼科コ・メディカル関係費:1
'支出/事業費/通信費:1
'支出/会議費/役員会費:1
'支出/出張費/出張旅費補助:1
'支出/慶弔費/慶弔費:1

    Debug.Print vbNewLine & "----------- by FindKeysAsString"
    Debug.Print accFinder.FindKeysAsString()
'0//
'収入/会費/A会員
'収入/会費/B会員
'収入/会費/C会員
'収入/補助金/補助金
'収入/雑収入/セミナー参加料
'収入/雑収入/広告料収入
'収入/雑収入/雑収入
'支出/事務費/消耗品費
'支出/事務費/通信費
'支出/事業費/公衆衛生費
'支出/事業費/学術費
'支出/事業費/広報費
'支出/事業費/庶務費
'支出/事業費/眼科コ・メディカル関係費
'支出/事業費/通信費
'支出/会議費/役員会費
'支出/出張費/出張旅費補助
'支出/慶弔費/慶弔費
'
    'TearDown
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb.Close
End Sub



Sub TasteAccountsFinder_specifyReportingUnit()
    ' AccountsFinderクラスを味見する ただし収支報告単位を特定して
    'Arrange
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("現金出納帳")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act
    Dim accFinder As AccountsFinder: Set accFinder = New AccountsFinder
    Call accFinder.Initialize(cb)
    Dim dic As Dictionary
    Set dic = accFinder.FindAccounts(ofReportingUnit:="東北ブロック講習会")
    'Assert
    Call KzCls
    Dim key As Variant
    For Each key In dic
        Debug.Print key & ":" & dic(key)
    Next
    
'支出/事務費/通信費:17
'支出/事業費/学術費:9
'支出/事業費/通信費:1

    'TearDown
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb.Close
End Sub





