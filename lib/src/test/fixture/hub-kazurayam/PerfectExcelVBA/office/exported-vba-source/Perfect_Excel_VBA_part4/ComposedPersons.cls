VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComposedPersons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**
'* The container of a Collection of Person objects equipped with various
'* accessor methods: LoadData, ApplyData, AddPerson, UpdatePerson
'*

Public Enum eFields
    eId = 1
    eName
    eGender
    eBirthday
    eActive
End Enum

'名簿のデータが収録されたワークシート
Private baseSheet As Worksheet

'Worksheetの中のテーブル、すなわち手入力されたデータの源
Private table As ListObject

'IDをkeyとしPersonオブジェクトをvalueとする
'keyはLong型である。
'例 1,2,3,4,...
Public persons As Dictionary

'personsに格納済みのIDのうち最も大きな値
'次にPersonオブジェクトを追加するときそのIDはMaxId+1となる
Public MaxId As Long

Private logger As BbLogger

'**
'* このMemberSheetImplオブジェクトを初期化する。すなわち
'* 1. originTableへの参照を記憶し
'* 2. personsコレクションをNewして
'* 3. originTableの行をPersonオブジェクトに変換してpersonsに格納し
'* 4. personsコレクションのサイズをMaxIdに記憶しておく
'*
Public Sub Initialize(ByRef ws As Worksheet)
    Set logger = BbLoggerFactory.CreateLogger("ComposedPersons")
    logger.procedureName = "Initialize"
    '
    Set baseSheet = ws
    Set table = baseSheet.ListObjects(1)
    Set persons = New Dictionary
    MaxId = 0
    'logger.Info "Initialized with Sheet:", ws.Name
End Sub

'**
'* 与えられたListObjectsのデータをPersonsコレクションに格納する
'*
Public Function LoadData() As Long
    logger.procedureName = "LoadData"
    '
    If table.ListRows.Count = 0 Then
        With err
            .Clear
            .Raise vbObjectError + 513, "ComposedPersons.Initialize", _
                "与えられたシートのテーブルが空っぽでした。与えたシートが正しい状態かどうか確かめてください。"
        End With
    End If
    '
    With table
        'logger.Info "table.ListRows.Count", .ListRows.Count
        Dim aRow As ListRow
        For Each aRow In .ListRows
            Dim aPerson As Person: Set aPerson = New Person
            aPerson.Initialize aRow.Range
            'PersonオブジェクトのIDプロパティ（Long型）をkeyとして
            'PersonオブジェクトをDictionaryに格納する
            persons.Add aPerson.ID, aPerson
            '
            If MaxId < aPerson.ID Then
                'personsに格納されたPersonオブジェクトのIDの最大値を覚えておく
                MaxId = aPerson.ID
            End If
            'logger.Info "MaxId", MaxId, "aPerson.ID", aPerson.ID
        Next aRow
    End With
    LoadData = MaxId
    Exit Function
    
ErrorHandler:
    With err
        Dim title As String: title = "エラーが発生しました"
        Dim m As String: m = ""
        m = m & "エラー番号:" & .number & vbNewLine
        m = m & "内容:" & .Description & vbNewLine
        MsgBox m, vbExclamation, title
        LoadData = 0
        Exit Function
    End With
End Function

'**
'* 現在のPersonsコレクションの内容を指定されたListObjectのなかに保存する
'*
Public Sub ApplyData()

    'シートの保護を一時的に解除する
    If baseSheet.ProtectContents Then
        baseSheet.Unprotect Relay.sheetPW
    End If
    
    'テーブルのAutoFilterを解除する
    table.ShowAutoFilter = False
    
    With table
        '出力先のワークシートのなかのテーブルをいったん空っぽにする
        If .ListRows.Count > 0 Then
            .DataBodyRange.EntireRow.Delete
        End If
        
        'personsディクショナリの内容をシートのなかのテーブルに書き込みたいが
        'パフォーマンスを良くするため準備をする。personsを2次元配列に直す。
        Dim values() As Variant: ReDim values(1 To MaxId, 1 To 5)
        
        Dim aKey As Variant
        For Each aKey In persons
            Dim p As Person: Set p = persons.Item(aKey)
            values(p.ID, eId) = p.ID
            values(p.ID, eName) = p.Name
            values(p.ID, eGender) = p.Gender
            values(p.ID, eBirthday) = p.Birthday
            values(p.ID, eActive) = p.Active
        Next aKey
        
        Dim baseRange As Range: Set baseRange = .Range(1, 1).Offset(1)
        baseRange.Resize(MaxId, 5).value = values
        
    End With
    
    'シートにAutoFilterを適用する
    table.ShowAutoFilter = True
    
    'シートを保護する
    baseSheet.Protect Relay.sheetPW, AllowFiltering:=True
End Sub

'**
'* Personsディクショナリに含まれたPersonオブジェクトのIDの配列を返す。
'* PersonオブジェクトのIDはLong型。それをString型に変換したVariant型配列を返す。
'* たとえば Array("1", "2", "3", ...) と同じ。
'* UserForm1が利用することを想定している。
'*
Public Function GetIdList() As Variant
    GetIdList = persons.Keys()
End Function

'**
'* 引数aPersonとして与えられたPersonオブジェクトをpersonsディクショナリのなかに
'* 書きこみ、かつワークシートのテーブルにも反映する。
'* 引数aPersonとして与えられたPersonオブジェクトがpersonsディクショナリのなかに
'* すでに存在している場合と、まだ存在していない場合と二通りありうる。
'* どちらの場合でもOK。
'* すでに存在していた場合にはaPersonの内容でディクショナリのValueを置き換える。
'* まだ存在していなかった場合には結果的にaPersonをディクショナリに追加する。
'*
'* @param aPerson {Person} Person object to update
'*
Public Sub UpdatePerson(ByVal aPerson As Person)
    If persons.Exists(aPerson.ID) Then
        With persons.Item(aPerson.ID)
            .ID = aPerson.ID
            .Name = aPerson.Name
            .Gender = aPerson.Gender
            .Birthday = aPerson.Birthday
            .Active = aPerson.Active
        End With
    Else
        Dim newPerson As Person: Set newPerson = New Person
        With newPerson
            .ID = aPerson.ID
            .Name = aPerson.Name
            .Gender = aPerson.Gender
            .Birthday = aPerson.Birthday
            .Active = aPerson.Active
        End With
        persons.Add aPerson.ID, newPerson
        MaxId = MaxId + 1
    End If
    Call ApplyData
End Sub

Public Function Exists(ByVal anID As String) As Boolean
    Exists = persons.Exists(CLng(anID))
End Function

'**
'* 「名簿」シートのB1セルがpinterCellパラメータとして与えられたとする。
'* B1セルのValueが "$D$14" という文字列であったとする。
'* $D$14セルがpersonsディクショナリのどのレコードに対応するかを演算で求める。
'* ここで「対応する」とは$D$14セルのデータがPersonオブジェクトXのデータであるということ。
'* なおセルの列番号$Dを無視して行番号$14だけでPersonオブジェクトとの対応関係を判定する。
'* 発見されたPersonオブジェクトを関数の値として返す。
'* もしも与えられたパラメータに対応するPersonが特定できなかったらNothingを返す。
'*
Public Function InterpreteSelectedRangeToPerson(ByRef pointerCell As Range) As Person
    logger.procedureName = "InterpreteSelectedRangeToPerson"
    'logger.Info "pointerCell.Value", pointerCell.value
    'logger.Info "persons.Count", persons.Count
    '
    Dim pointedCell As Range
    Set pointedCell = baseSheet.Range(pointerCell.value)
    'logger.Info "pointedCell.Row", pointedCell.Row
    'logger.Info "table.DataBodyRange.Address", table.DataBodyRange.Address
    'logger.Info "table.DataBodyRange.Row", table.DataBodyRange.Row
    '
    Dim personId As Long
    personId = pointedCell.Row - table.DataBodyRange.Row + 1
    'logger.Info "personId", personId
    '
    If 0 < personId And personId <= MaxId Then
        Set InterpreteSelectedRangeToPerson = persons.Item(personId)
    Else
        Set InterpreteSelectedRangeToPerson = Nothing
    End If
End Function


