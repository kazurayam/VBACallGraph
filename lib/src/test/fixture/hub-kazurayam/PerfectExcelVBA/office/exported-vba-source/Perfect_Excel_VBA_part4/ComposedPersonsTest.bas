Attribute VB_Name = "ComposedPersonsTest"
Option Explicit


' ���Ёu�p�[�t�F�N�gExcel VBA�v�����鐬
' ��14 �A�v���P�[�V�����J��
' kazurayam modified the design drastically

'**
'* ����Ǘ�prj.ComposedPersonsTest
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
    ' TestMemberSheet�V�[�g�̒��̃e�[�u��������������������
    ' MemberSheet�V�[�g�̒��̃e�[�u���̍s���R�s�[����
    ' TestMemberSheet�̃e�[�u���ɒǉ�����B
    ' TestInitialize���\�b�h��Rubberduck�ɂ����Call�����B
    ' �e�e�X�g���\�b�h�����s����钼�O��Call�����B
    ' ���ʓI��TestMemberSheet�V�[�g�̃e�[�u���͊e�e�X�g���\�b�h��
    ' ���s����钼�O��MemberSheet�̃e�[�u���̓��e�ŏ㏑�������B
    
    With ThisWorkbook
        Dim sourceSheet As Worksheet:  Set sourceSheet = .Sheets("����")
        Dim sourceTable As ListObject: Set sourceTable = sourceSheet.ListObjects(1)
        Dim targetSheet As Worksheet:  Set targetSheet = .Sheets("����e�X�g�p")
        Dim targetTable As ListObject: Set targetTable = targetSheet.ListObjects(1)
        
        '����e�X�g�p�V�[�g�̕ی���ꎞ�I�ɉ�������
        targetSheet.Unprotect Relay.sheetPW
            
        ' ����e�X�g�p�V�[�g�̃e�[�u���������������ۂɂ���
        If targetTable.ListRows.Count > 0 Then
            targetTable.DataBodyRange.EntireRow.Delete
        End If
        ' ����V�[�g���疼��e�X�g�p�V�[�g�փe�[�u�����s�P�ʂɃR�s�[����
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
            
        '����e�X�g�p�V�[�g��������x�ی삷��
        targetSheet.Protect Relay.sheetPW, AllowFiltering:=True
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("ComposedPersons��Initialize���\�b�h���e�X�g����")
Private Sub TestPersonsInitialize()
    logger.procedureName = "TestPersonsInitialize"
    Dim cp As ComposedPersons: Set cp = New ComposedPersons
    Call cp.Initialize(TempMemberSheet)
    Call cp.LoadData
    
    Assert.IsTrue CLng(0) < cp.MaxId
End Sub


'@TestMethod("Exists���\�b�h���e�X�g����")
Private Sub TestExists()
    logger.procedureName = "TestExists"
    '
    Dim cp As ComposedPersons: Set cp = New ComposedPersons
    Call cp.Initialize(TempMemberSheet)
    Call cp.LoadData
    '
    Assert.IsTrue cp.Exists("1")
End Sub

'@TestMethod("����V�[�g��LoadData�v���V�W�����e�X�g����")
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
        Assert.AreEqual "���� ��", .Name
        Assert.AreEqual "�j", .Gender
        Assert.AreEqual #3/30/1988#, .Birthday
        Assert.AreEqual True, .Active
    End With
End Sub

'@TestMethod("����V�[�g��LoadData�v���V�W�����e�X�g����")
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

'@TestMethod("GetIDList���\�b�h���e�X�g����")
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

'@TestMethod("UpdatePerson(aPerson)���\�b�h���e�X�g����---�����̃L�[")
Private Sub TestUpdatePersonExisting()
    logger.procedureName = "TestUpdatePersonExisting"
    '
    Dim cp As ComposedPersons: Set cp = New ComposedPersons
    Call cp.Initialize(TempMemberSheet)
    With cp
        .LoadData
        Assert.IsTrue .persons.Count = 16
        'persons�f�B�N�V���i���ɂ��łɂ���Person�I�u�W�F�N�g�������Ƃ���UpdatePerson����
        Dim aPerson As Person
        Set aPerson = .persons.Item(1)
        aPerson.Birthday = #3/30/1989#
        Call cp.UpdatePerson(aPerson)
        Assert.IsTrue .persons.Count = 16
    End With
End Sub

'@TestMethod("UpdatePerson(aPerson)���\�b�h���e�X�g����---New�L�[")
Private Sub TestUpdatePersonNew()
    logger.procedureName = "TestUpdatePersonNew"
    '
    Dim cp As ComposedPersons: Set cp = New ComposedPersons
    Call cp.Initialize(TempMemberSheet)
    With cp
        .LoadData
        Assert.IsTrue .persons.Count = 16
        '�V����Person�I�u�W�F�N�g�������UpdatePerson����
        Dim newPerson As Person
        Set newPerson = New Person
        Call newPerson.Construct(.persons.Count + 1, _
                        "�����鐬", "�j", #9/21/1981#, False)
        cp.UpdatePerson newPerson
        Assert.IsTrue .persons.Count = 17
    End With
End Sub


