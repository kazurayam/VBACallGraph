VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Personクラス

Public ID As Long
Public Name As String
Public Gender As String
Public Birthday As Date
Public Active As Boolean

'**
'* Personクラスのインスタンスに初期値を設定する
'*
'* @param myRange {Range} 設定するレコードのセル範囲
'*
Public Sub Initialize(ByVal aRange As Range)
    ID = aRange(eFields.eId).value
    Name = aRange(eFields.eName).value
    Gender = aRange(eFields.eGender).value
    Birthday = aRange(eFields.eBirthday).value
    Active = aRange(eFields.eActive).value
End Sub

Public Sub Construct(ByVal anID As Long, _
                        ByVal aName As String, _
                        ByVal aGender As String, _
                        ByVal aBirthday As Date, _
                        ByVal isActive As Boolean)
    ID = anID
    Name = aName
    Gender = aGender
    Birthday = aBirthday
    Active = isActive
End Sub

'**
'* Personオブジェクトの年齢を取得する
'*
'* @return {Long} 年齢
'*
Public Property Get Age() As Long
    Dim myAge As Long
    '現時点の日付とBirthdayプロパティの「年」の差をとる
    myAge = DateDiff("yyyy", Me.Birthday, Date)
    '今日が今年の誕生日を過ぎている場合と未だ過ぎていない場合とがある。
    '今日が今年の誕生日を過ぎていたらmyAgeから1だけ減らす
    If Date < DateSerial(Year(Now), Month(Me.Birthday), Day(Me.Birthday)) Then
        myAge = myAge - 1
    End If
    Age = myAge
End Property
