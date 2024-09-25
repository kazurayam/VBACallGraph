Attribute VB_Name = "PersonTest"
Option Explicit


Option Private Module

' ���Ёu�p�[�t�F�N�gExcel VBA�v�����鐬
' ��14 �A�v���P�[�V�����J��

'**
'* PerfectBook_part4.PersonTest
'*

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Private logger As BbLogger

Private aPerson As Person

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set logger = BbLoggerFactory.CreateLogger("PersonTest")
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
    'This method runs before every test in the module..
    Dim datasheet As Worksheet: Set datasheet = Sheets("����")
    With datasheet
        Set aPerson = New Person
        aPerson.Initialize datasheet.ListObjects(1).ListRows(1).Range
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("����V�[�g��1�s����͂Ƃ���Person�I�u�W�F�N�g�𐶐�����")
Private Sub TestInstantiatePerson()
    logger.procedureName = "TestInstantiatePerson"
    '
    logger.Info "aPerson.Id", aPerson.ID
    logger.Info "aPerson.Name", aPerson.Name
    logger.Info "aPerson.Gender", aPerson.Gender
    logger.Info "aPerson.Birthday", aPerson.Birthday
    logger.Info "aPerson.Active", aPerson.Active
        
    Assert.AreEqual CLng(1), aPerson.ID
    Assert.AreEqual "���� ��", aPerson.Name
    Assert.AreEqual "�j", aPerson.Gender
    Assert.AreEqual #3/30/1988#, aPerson.Birthday
    Assert.AreEqual True, aPerson.Active
End Sub

'@TestMethod("Age property")
Private Sub TestAgeProperty()
    '�����ׂ���̒a������3/30/1988�ł���
    '��������̍��N�̒a�������������Ƃ�����
    Dim birthdayThisYear As Date
    birthdayThisYear = DateSerial(Year(Now), Month(aPerson.Birthday), Day(aPerson.Birthday))
    
    '������2024/2/1�܂荡�N�̒a�������O�Ȃ�2024 - 1988=36���1���Ȃ�35�΂��������N��
    '������2024/4/1�܂荡�N�̒a��������Ȃ��2024 - 1988=36�΂��������N��
    Dim expected As Long: expected = DateDiff("yyyy", aPerson.Birthday, Date)
    If Date < birthdayThisYear Then
        expected = expected - 1
    End If
    
    logger.Info "aPerson.Age", aPerson.Age
    Assert.AreEqual expected, aPerson.Age
End Sub
