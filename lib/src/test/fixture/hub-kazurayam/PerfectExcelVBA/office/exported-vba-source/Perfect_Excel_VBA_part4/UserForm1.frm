VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "名簿管理"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cp As ComposedPersons

Public Sub LoadData(ByRef baseSheet As Worksheet)
    Set cp = New ComposedPersons
    Call cp.Initialize(baseSheet)
    Call cp.LoadData
    Call LoadIdList
    
    '名簿シートのRange("B1")に "$B$14" のようなアドレス文字列が書いてある。
    'このアドレスにがどのPersonオブジェクトに対応するかを判定してい
    'そのPersonのデータをフォームに展開する
    Dim aPerson As Person
    Set aPerson = cp.InterpreteSelectedRangeToPerson(baseSheet.Range("B1"))
    If Not aPerson Is Nothing Then
        LoadFields aPerson.ID
    End If
End Sub


'**
'* cpからIDのリストを読み込んでコンボボックスComboBoxIdにセットする
'*
Private Sub LoadIdList()
    ComboBoxId.List = cp.GetIdList()
    ComboBoxId.AddItem "New"
End Sub

'**
'* イベントプロシージャ： ComboBoxId_Change
'*
Private Sub ComboBoxId_Change()
    With ComboBoxId
        If IsValidId Then
            If Information.IsNumeric(.value) Then
                Call LoadFields(.value)
            Else
                Call ClearFields
            End If
        End If
    End With
End Sub

Private Property Get IsValidId() As Boolean
    IsValidId = False
    With ComboBoxId
        If (.value > 0 And .value <= cp.MaxId) Or _
            (.value = "New") Then
            IsValidId = True
        End If
    End With
End Property

'**
'* anIdで指定されたIDのPersonオブジェクトをPersonsコレクションから
'* 取り出して、UserFormの各コントロールの値としてセットする
'*
Private Sub LoadFields(ByVal anID As Long)
    With cp.persons.Item(anID)
        ComboBoxId.value = anID
        TextBoxName.value = .Name
        Call SetGender(.Gender)
        TextBoxBirthday.value = .Birthday
        LabelAge.Caption = .Age
        CheckBoxActive.value = .Active
    End With
End Sub

'**
'* 性別を表す文字列（「男」「女」）にもとづいて
'* オプションボタンの値を設定する
'*
Private Sub SetGender(ByVal aGender As String)
    OptionButtonFemale.value = True
    If aGender = "男" Then
        OptionButtonMale.value = True
    End If
End Sub

'**
'* UserFormの各コントロールの値を初期化する
'*
Private Sub ClearFields()
    TextBoxName.value = ""
    OptionButtonMale.value = True
    TextBoxBirthday.value = ""
    LabelAge.Caption = ""
    CheckBoxActive.value = True
End Sub

'**
'* Event Procedure:
'* フォームの「更新」ボタンが押下されたときに起動される
'*
Private Sub CommandButtonUpdate_Click()
    '
    If CheckFields Then
        Dim p As Person: Set p = New Person
        If ComboBoxId.value = "New" Then
            p.ID = cp.MaxId + 1
        Else
            p.ID = CLng(ComboBoxId.value)
        End If
        p.Name = TextBoxName.text
        p.Birthday = TextBoxBirthday.value
        p.Gender = "女"
        If OptionButtonMale.value = True Then p.Gender = "男"
        p.Active = CheckBoxActive.value
        
        'このPersonオブジェクトの状態をComposedPersonsに記録する
        cp.UpdatePerson p
        ' 更新されたComposePersonsをワークシートのテーブルに反映する
        cp.ApplyData
        
        '更新されたComposedPersonsに同期するようにフォームの表示を更新する
        Call LoadFields(p.ID)
        Call LoadIdList
        
    End If
End Sub

'**
'* フォーム上にある各コントロールの値が正しく入力されているかどうかを
'* 判定する
'*
Private Function CheckFields() As Boolean
    
    CheckFields = True
    
    If Not IsValidId Then
        MsgBox "「ID」としては1以上で" & cp.MaxId & "以下の数値または""New""を入力してください", vbInformation
        CheckFields = False
    End If
    
    If Len(TextBoxName.text) = 0 Then
        MsgBox "「名前」を入力してください", vbInformation
        CheckFields = False
    End If
    
    If Not IsDate(TextBoxBirthday.value) Then
        MsgBox "「誕生日」に日付を入力してください", vbInformation
        CheckFields = False
    End If
    
End Function


'**
'* Event Procedure:
'* フォームの「閉じる」ボタンが押されたときに起動される
'*
Private Sub CommandButtonClose_Click()
    Unload Me
End Sub
