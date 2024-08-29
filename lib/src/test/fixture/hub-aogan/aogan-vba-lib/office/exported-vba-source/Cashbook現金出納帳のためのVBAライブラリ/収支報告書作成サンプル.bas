Attribute VB_Name = "収支報告書作成サンプル"
Option Explicit

' ワークシート「work」のなかの「第４５回Tブロック講習会会計報告」のなかに数字を書き込む。
' 本プロジェクトのsrcフォルダの下に「現金出納帳青森県眼科医会令和4年度.xlsm」がある。
' このワークブックの「現金出納帳」ワークシートに書かれた入出金データを入力として参照する。

Sub Main_TohokuBlockLectures45th()

    Call KzCls
    Debug.Print "Main_TohokuBlockLectuers45th STARTED"
    
    ThisWorkbook.Activate
    
    '出力先となるワークシートを特定する
    Dim outWs As Worksheet
    Set outWs = ActiveWorkbook.Worksheets("work")
    
    '明細の出力先となるテーブルを宣言する
    Dim fullTbl As ListObject: Set fullTbl = outWs.ListObjects("テーブル1")
    '明細のテーブルを初期化する
    'テーブルが空でないことを確認してからDeleteする
    If Not fullTbl.DataBodyRange Is Nothing Then
        fullTbl.DataBodyRange.Delete
    End If
    
    '収支報告単位を宣言する
    Dim rpUnit As String: rpUnit = "東北ブロック講習会"

    '===================================================================================
    '東北眼科医会連合会青森県代表の現金出納帳であるExcelファイルを開く
    ThisWorkbook.Activate
    Debug.Print ">>" & KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B3")
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B3"))
    Dim cb1 As Cashbook: Set cb1 = CreateCashbook(wb1, "現金出納帳", "CashbookTable2")
    Call PrintAccounts(cb1, rpUnit)
    '
    Dim cs1 As CashSelector: Set cs1 = New CashSelector
    Call cs1.Initialize(cb1, periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#)
    
    '集計
    Call TranscribeSum(cs1, rpUnit, AccountType.Income, "雑収入", "セミナー参加料", outWs.Range("$G$7"))
    Call TranscribeSum(cs1, rpUnit, AccountType.Income, "雑収入", "広告料収入", outWs.Range("$G$8"))
    
    '明細
    Call TranscribeDetail(cs1, rpUnit, AccountType.Income, "雑収入", "セミナー参加料", fullTbl)
    Call TranscribeDetail(cs1, rpUnit, AccountType.Income, "雑収入", "広告料収入", fullTbl)
    
    '後始末。外部のExcelファイルを閉じる
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb1.Close
    Set cs1 = Nothing
    Set cb1 = Nothing
    Set wb1 = Nothing
    
    '===================================================================================
    '青森県眼科医会の現金出納帳であるExcelファイルを開く
    ThisWorkbook.Activate
    Debug.Print ">>" & KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Dim cb2 As Cashbook: Set cb2 = CreateCashbook(wb2, "現金出納帳", "CashbookTable1")
    '青森県眼科医会の現金出納帳に書かれた入出金データで「東北ブロック講習会」に関連する勘定科目を列挙する
    Call PrintAccounts(cb2, rpUnit)
    
    'CashSelectorオブジェクトを生成し、Cashbookオブジェクトを参照するよう設定する
    Dim cs2 As CashSelector: Set cs2 = New CashSelector
    Call cs2.Initialize(cb2, periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#)
    
    '集計
    '勘定科目ごとの明細を問い合わせ、金額の小計を求める。
    Call TranscribeSum(cs2, rpUnit, AccountType.Expense, "事業費", "学術費", outWs.Range("$F$10"))
    Call TranscribeSum(cs2, rpUnit, AccountType.Expense, "事業費", "通信費", outWs.Range("$F$11"))
    Call TranscribeSum(cs2, rpUnit, AccountType.Expense, "事務費", "通信費", outWs.Range("$F$12"))
    
    '明細
    Call TranscribeDetail(cs2, rpUnit, AccountType.Expense, "事業費", "学術費", fullTbl)
    Call TranscribeDetail(cs2, rpUnit, AccountType.Expense, "事業費", "通信費", fullTbl)
    Call TranscribeDetail(cs2, rpUnit, AccountType.Expense, "事務費", "通信費", fullTbl)
    
    '後始末。外部のExcelファイルを閉じる
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb2.Close
    Set cs2 = Nothing
    Set cb2 = Nothing
    Set wb2 = Nothing
    
    
    Debug.Print "Main_TohokuBlockLectuers45th FINISHED"
End Sub


Private Sub TranscribeSum(ByRef cs As CashSelector, _
                                ByVal theReportingUnit As String, _
                                ByVal AccType As AccountType, _
                                ByVal accName As String, _
                                ByVal subAccName As String, _
                                ByRef targetCell As Range)
    '指定された勘定科目の入出金の金額を合算し、指定されたセルに転記する
    
    'Debug.Print theReportingUnit & "," & accType & "," & accName & "," & subAccName & "," & targetCell
    
    'パラメータで指定された
    Dim selected As CashList
    If AccType = AccountType.Expense Then
        Set selected = cs.SelectCashList(AccType, accName, subAccName, theReportingUnit)
        targetCell.value = selected.SumOfExpenseAmount()
    Else
        Set selected = cs.SelectCashList(AccType, accName, subAccName, theReportingUnit)
        targetCell.value = selected.SumOfIncomeAmount()
    End If
    'tearDown
    Set selected = Nothing
End Sub

Private Sub PrintAccounts(ByVal cb As Cashbook, ByVal rpUnit As String)
    Dim accFinder As AccountsFinder: Set accFinder = New AccountsFinder
    Call accFinder.Initialize(cb, periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#)
    Dim dic As Dictionary: Set dic = accFinder.FindAccounts(rpUnit)
    Dim key As Variant
    For Each key In dic
        Debug.Print key & ":" & dic(key)
    Next
End Sub


Private Sub TranscribeDetail(ByVal cs As CashSelector, _
                            ByVal rpUnit As String, _
                            ByVal AccType As AccountType, _
                            ByVal ofAccountName As String, _
                            ByVal ofSubAccountName As String, _
                            ByRef targetTable As ListObject)
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccType, ofAccountName, ofSubAccountName, rpUnit)
    Dim i As Long
    Dim ch As Cash
    For i = 1 To selected.Count
        Set ch = selected.Items(i)
        With targetTable.ListRows.Add
            .Range(1).value = ch.ItsAccount.AccountTypeAsString
            .Range(2).value = ch.ItsAccount.AccountName
            .Range(3).value = ch.ItsAccount.SubAccountName
            .Range(4).value = ch.Description
            If ch.ExpenseAmount > 0 Then
                .Range(5).value = ch.ExpenseAmount
            End If
            If ch.IncomeAmount > 0 Then
                .Range(6).value = ch.IncomeAmount
            End If
            .Range(7).value = ch.ToDate()
        End With
    Next i
End Sub
