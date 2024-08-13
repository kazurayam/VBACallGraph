Attribute VB_Name = "KzCollection"
Option Explicit

'KzCollection

' https://y-moride.com/vba/collection-key-exists.html#toc1
'*********************************************************
'* ExistsKey（Collention内のキー検索関数）
'*********************************************************
'* 第１引数 | Collection | 検索対象となるオブジェクト
'* 第２引数 |   String   | 検索するキー
'*  戻り値　|   Boolan   | True Or False ※False＠初期値
'*********************************************************
'*   説明   | 第２引数をキーとしてItemメソッドを実行し、
'*   　　   | 結果をもとにキーの存在を確認する。
'*********************************************************
'*   備考   | オブジェクト未設定の場合 ⇒ 戻り値「False」
'*   　　   | メンバー数「0」の場合 ⇒ 戻り値「False」
'*********************************************************

Public Function KzExistsKey(objCol As Collection, strKey As String) As Boolean
     
    '戻り値の初期値：False
    KzExistsKey = False
     
    '変数にCollection未設定の場合は処理終了
    If objCol Is Nothing Then Exit Function
     
    'Collectionのメンバー数が「0」の場合は処理終了
    If objCol.Count = 0 Then Exit Function
     
    On Error Resume Next
     
    'Itemメソッドを実行
    Call objCol.item(strKey)
         
    'エラー値がない場合：キー検索はヒット（戻り値：True）
    If Err.Number = 0 Then KzExistsKey = True
 
End Function


