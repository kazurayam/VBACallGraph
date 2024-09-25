Attribute VB_Name = "Relay"
Option Explicit

Private Const SHEETPASSWORD As String = "umeume0416"

'**
'* 名簿シートに保護を適用するさいに指定するパスワード
'*
Public Property Get sheetPW()
    sheetPW = SHEETPASSWORD
End Property

'**
'* 「名簿」シートの「名簿管理」ボタンがクリックされたらこのShowUserFormを呼び出す。
'* 「名簿テスト用」シートの「名簿管理」ボタンも同じくこのShowUserFormを呼び出す。
'*
Public Sub ShowUserForm()
    With UserForm1
        'フォームを表示する位置を指定する
        .StartUpPosition = 0
        .Left = 320
        .Top = 220
        
        '名簿管理フォームを開く
        UserForm1.Show vbModeless
    
        '名簿シートのデータをUserForm1のなかにロードする
        UserForm1.LoadData ActiveSheet
    End With
End Sub
