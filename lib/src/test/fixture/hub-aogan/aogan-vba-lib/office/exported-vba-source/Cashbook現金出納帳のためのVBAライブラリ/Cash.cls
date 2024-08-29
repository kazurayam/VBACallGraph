VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Cashクラス: お金の出し入れ1件1件を表すクラス。
' 現金出納帳に記載された1件1件の入出金記録を表すクラス

Public ReceiptNo As Long    '領収書No.
Public YY As Long           '年度（令和）
Public MM As Long           '月
Public DD As Long           '日
Public incomeAccount As String      '収入科目
Public IncomeSubAccount As String   '収入補助科目
Public expenseAccount As String     '支出科目
Public ExpenseSubAccount As String  '支出補助科目
Public ReportingUnit As String      '収支報告単位
Public Description As String        '適用
Public IncomeAmount As Long         '借方金額
Public ExpenseAmount As Long        '貸方金額


Public Sub Initialize(ByVal rng As Range)
    ' ワークシートの1行であるRangeオブジェクトを受け取って
    ' Cashオブジェクトのプロパティに代入する
    ReceiptNo = rng(1).value
    YY = rng(2).value
    MM = rng(3).value
    DD = rng(4).value
    incomeAccount = rng(5).value
    IncomeSubAccount = rng(6).value
    expenseAccount = rng(7).value
    ExpenseSubAccount = rng(8).value
    ReportingUnit = rng(9).value
    Description = rng(10).value
    IncomeAmount = rng(11).value
    ExpenseAmount = rng(12).value
End Sub


Private Sub Class_Initialize()
    ' 必要におうじてインスタンス変数を初期化する
End Sub

Public Function ColumnHeader() As String
    Dim ch As String
    ch = "|"
    ch = ch & "|" & "領収書No."
    ch = ch & "|" & "年"
    ch = ch & "|" & "月"
    ch = ch & "|" & "日"
    ch = ch & "|" & "収入科目"
    ch = ch & "|" & "収入補助科目"
    ch = ch & "|" & "支出科目"
    ch = ch & "|" & "支出補助科目"
    ch = ch & "|" & "収支報告単位"
    ch = ch & "|" & "適用"
    ch = ch & "|" & "借方金額"
    ch = ch & "|" & "貸方金額"
    ch = ch & "|"
    ColumnHeader = ch
End Function

Public Function ToString() As String
    Dim s As String
    s = "|"
    s = s & "|" & str(ReceiptNo)
    s = s & "|" & str(YY)
    s = s & "|" & str(MM)
    s = s & "|" & str(DD)
    s = s & "|" & incomeAccount
    s = s & "|" & IncomeSubAccount
    s = s & "|" & expenseAccount
    s = s & "|" & ExpenseSubAccount
    s = s & "|" & ReportingUnit
    s = s & "|" & Description
    s = s & "|" & str(IncomeAmount)
    s = s & "|" & str(ExpenseAmount)
    s = s & "|"
    ToString = s
End Function

Public Function ToDate() As Date
    ' 年YYと月MMと日DDから日付を生成し西暦のDateとして返す
    ' 現金出納帳の年YYは和暦であるはずだから、適切に変換する。
    ' 令和を西暦年に変換するには（手抜きだとしりつつ）YYに整数2018を+することにした
    Dim d As Date
    d = DateSerial(YY + 2018, MM, DD)
    ToDate = d
End Function

Public Property Get ItsAccount() As Account
    Dim acc As Account: Set acc = New Account
    If Len(incomeAccount) > 0 And Len(IncomeSubAccount) > 0 Then
        Call acc.Initialize(AccountType.Income, incomeAccount, IncomeSubAccount)
    ElseIf Len(expenseAccount) > 0 And Len(ExpenseSubAccount) > 0 Then
        Call acc.Initialize(AccountType.Expense, expenseAccount, ExpenseSubAccount)
    Else
        Call acc.Initialize(AccountType.Undefined, "", "")
    End If
    Set ItsAccount = acc
End Property
