Attribute VB_Name = "ComposedPersonsTest"
Option Explicit


' 書籍「パーフェクトExcel VBA」高橋宣成
' §14 アプリケーション開発
' kazurayam modified the design drastically

'**
'* 名簿管理prj.ComposedPersonsTest
'*

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Private logger As BbLogger

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set logger = BbLoggerFactory.CreateLogger("ComposedPersonsTest")
    logger.Clear
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    logger.procedureName = "TestInitialize"
    'This method runs before every test in the module..
    ' TestMemberSheetシートの中のテーブルを初期化したうえで
    ' MemberSheetシートの中のテーブルの行をコピーして
    ' TestMemberSheetのテーブルに追加する。
    ' TestInitializeメソッドはRubberduckによってCallされる。
    ' 各テストメソッドが実行される直前にCallされる。
    ' 結果的にTestMemberSheetシートのテーブルは各テストメソッドが
    ' 実行される直前にMemberSheetのテーブルの内容で上書きされる。
    
    With ThisWorkbook
        Dim sourceSheet As Worksheet:  Set sourceSheet = .Sheets("名簿")
        Dim sourceTable As ListObject: Set sourceTable = sourceSheet.ListObjects(1)
        Dim targetSheet As Worksheet:  Set targetSheet = .Sheets("名簿テスト用")
        Dim targetTable As ListObject: Set targetTable = targetSheet.ListObjects(1)
        
        '名簿テスト用シートの保護を一時的に解除する
        targetSheet.Unprotect Relay.sheetPW
            
        ' 名簿テスト用シートのテーブルをいったん空っぽにする
        If targetTable.ListRows.Count > 0 Then
            targetTable.DataBodyRange.EntireRow.Delete
        End If
        ' 名簿シートから名簿テスト用シートへテーブルを行単位にコピーする
        Dim record As ListRow
        Dim aSourceRow As ListRow
        For Each aSourceRow In sourceTable.ListRows
            With aSourceRow
                Set record = targetTable.ListRows.Add
                record.Range.value = Array(.Range(1).value, _
                                            .Range(2).value, _
                                            .Range(3).value, _
                                            .Range(4).value, _
                                            .Range(5).value)
            End With
        Next aSourceRow
            
        '名簿テスト用シートをもう一度保護する
        targetSheet.Protect Relay.sheetPW, AllowFiltering:=True
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("ComposedPersonsのInitializeメソッドをテストする")
Private Sub TestPersonsInitialize()
    logger.procedureName = "TestPersonsInitialize"
    Dim cp As ComposedPersons: Set cp = New ComposedPersons
    Call cp.Initialize(TempMemberSheet)
    Call cp.LoadData
    
    Assert.IsTrue CLng(0) < cp.MaxId
End Sub


'@TestMethod("Existsメソッドをテストする")
Private Sub TestExists()
    logger.procedureName = "TestExists"
    '
    Dim cp As ComposedPersons: Set cp = New ComposedPersons
    Call cp.Initialize(TempMemberSheet)
    Call cp.LoadData
    '
    Assert.IsTrue cp.Exists("1")
End Sub

'@TestMethod("名簿シートのLoadDataプロシジャをテストする")
Private Sub TestLoadData()
    logger.procedureName = "TestLoadData"
    '
    Dim cp As ComposedPersons: Set cp = New ComposedPersons
    Call cp.Initialize(TempMemberSheet)
    Call cp.LoadData
    '
    logger.Info "cp.MaxId", cp.MaxId
    Assert.IsTrue 0 < cp.MaxId
    '
    With cp.persons.Item(1)
        logger.Info ".ID", .ID
        logger.Info ".Name", .Name
        logger.Info ".Gender", .Gender
        logger.Info ".Birthday", .Birthday
        logger.Info ".Active", .Active
        
        Assert.AreEqual CLng(1), .ID
        Assert.AreEqual "横尾 勉", .Name
        Assert.AreEqual "男", .Gender
        Assert.AreEqual #3/30/1988#, .Birthday
        Assert.AreEqual True, .Active
    End With
End Sub

'@TestMethod("名簿シートのLoadDataプロシジャをテストする")
Private Sub TestApplyData()
    logger.procedureName = "TestApplyData"
    '
    Dim cp As ComposedPersons: Set cp = New ComposedPersons
    Call cp.Initialize(TempMemberSheet)
    
    With cp
        .LoadData
        Assert.IsTrue .persons.Count > 0
        With .persons.Item(1)
            .Birthday = #3/30/1989#
        End With
        .ApplyData
    End With
End Sub

'@TestMethod("GetIDListメソッドをテストする")
Private Sub TestGetIDList()
    logger.procedureName = "TestGetIDList"
    '
    Dim cp As ComposedPersons: Set cp = New ComposedPersons
    Call cp.Initialize(TempMemberSheet)
    Call cp.LoadData
    '
    Dim idList As Variant: idList = cp.GetIdList
    '
    logger.Info idList
    Assert.AreEqual CLng(16), UBound(idList) + 1
End Sub

'@TestMethod("UpdatePerson(aPerson)メソッドをテストする---既存のキー")
Private Sub TestUpdatePersonExisting()
    logger.procedureName = "TestUpdatePersonExisting"
    '
    Dim cp As ComposedPersons: Set cp = New ComposedPersons
    Call cp.Initialize(TempMemberSheet)
    With cp
        .LoadData
        Assert.IsTrue .persons.Count = 16
        'personsディクショナリにすでにあるPersonオブジェクトを引数としてUpdatePersonする
        Dim aPerson As Person
        Set aPerson = .persons.Item(1)
        aPerson.Birthday = #3/30/1989#
        Call cp.UpdatePerson(aPerson)
        Assert.IsTrue .persons.Count = 16
    End With
End Sub

'@TestMethod("UpdatePerson(aPerson)メソッドをテストする---Newキー")
Private Sub TestUpdatePersonNew()
    logger.procedureName = "TestUpdatePersonNew"
    '
    Dim cp As ComposedPersons: Set cp = New ComposedPersons
    Call cp.Initialize(TempMemberSheet)
    With cp
        .LoadData
        Assert.IsTrue .persons.Count = 16
        '新しいPersonオブジェクトを作ってUpdatePersonする
        Dim newPerson As Person
        Set newPerson = New Person
        Call newPerson.Construct(.persons.Count + 1, _
                        "高橋宣成", "男", #9/21/1981#, False)
        cp.UpdatePerson newPerson
        Assert.IsTrue .persons.Count = 17
    End With
End Sub


