VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cashbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Cashbookクラス

' 現金出納帳_青森県眼科医会_令和4年度.xlsmワークブックのなかの
' "現金出納帳"ワークシートのなかの
' "テーブル１"に記録された入出金データ（Cashオブジェクト）を集めたオブジェクト。
' 件数を問い合わせるメソッドや
' 条件を指定して入出金データのリストを応答するメソッド
' などを実装している。
'

Private cashRecords As Collection


Private Sub Class_Initialize()
    'インスタンス変数を初期化する
    Set cashRecords = New Collection
End Sub



Public Sub Initialize(ByVal tbl As ListObject)
    ' 現金出納帳ワークブックのなかの現金出納帳ワークシートのなかのテーブルを引数として受け取る
    ' テーブルのすべての行について次のことを行う。すなわち
    ' 1行をCashingオブジェクト1個に変換し
    ' entries_変数に格納する
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim cs As Cash
        Set cs = New Cash
        Call cs.Initialize(tbl.ListRows(i).Range)
        Call cashRecords.Add(cs, str(i))
    Next i
End Sub



Public Function Count() As Long
    Count = cashRecords.Count
End Function


Public Function Items(ByVal index As Long) As Cash
    Set Items = cashRecords.item(index)
End Function


Public Function GetCashByKey(ByVal key As String) As Cash
    Set GetCashByKey = cashRecords.item(key)
End Function

