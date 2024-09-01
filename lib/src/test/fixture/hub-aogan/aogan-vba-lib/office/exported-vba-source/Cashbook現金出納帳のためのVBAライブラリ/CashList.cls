VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CashList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' CashListクラス

' Cashオブジェクトの集合体（リスト）を表すクラス
' CashbookオブジェクトのSelectメソッドが結果を返すのにCashListオブジェクトを使う
' CashListオブジェクトはCashオブジェクトのCollectionを内包し、金額の合計（Sum）を計算するメソッドなどを提供する

Private csList As Collection

Private Sub Class_Initialize()
    'インスタンス変数を初期化する
    Set csList = New Collection
End Sub


Public Sub Add(ByVal cs As cash)
    Call csList.Add(cs)
End Sub


Public Function Count() As Long
    Count = csList.Count
End Function


Public Function Items(ByVal index As Long) As cash
    Set Items = csList(index)
End Function


Public Function SumOfIncomeAmount() As Long
    ' 収入金額の合計を計算して返す
    Dim sum As Long: sum = 0
    Dim i As Long
    For i = 1 To csList.Count
        Dim cs As cash: Set cs = csList(i)
        sum = sum + cs.IncomeAmount
    Next i
    SumOfIncomeAmount = sum
End Function


Public Function SumOfExpenseAmount() As Long
    ' 支出金額の合計を計算して返す
    Dim sum As Long: sum = 0
    Dim i As Long
    For i = 1 To csList.Count
        Dim cs As cash: Set cs = csList(i)
        sum = sum + cs.ExpenseAmount
    Next i
    SumOfExpenseAmount = sum
End Function


Public Function ToString() As String
    Dim s As String: s = ""
    Dim i As Long
    For i = 1 To csList.Count
        Dim cs As cash: Set cs = csList(i)
        s = s & cs.ToString() & vbCrLf
    Next i
    ToString = s
End Function
