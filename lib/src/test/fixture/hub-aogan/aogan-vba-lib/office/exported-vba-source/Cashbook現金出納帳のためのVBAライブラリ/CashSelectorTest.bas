Attribute VB_Name = "CashSelectorTest"
Option Explicit
Option Private Module

' TestCashSelector : CashSelector�N���X�����j�b�g�e�X�g����

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object


Dim wb As Workbook
Dim ws As Worksheet
Dim tbl As ListObject

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    
    Set wb = Workbooks.Open(BbUtil.ResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Set ws = wb.Worksheets("�����o�[��")
    Set tbl = ws.ListObjects("CashbookTable1")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    'TearDown
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb.Close
    Set wb = Nothing
    Set ws = Nothing
    Set tbl = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("CashSelector�C���X�^���X�𐶐���SelectCashList���\�b�h���e�X�g���� - Income�̏ꍇ")
Private Sub Test_SelectCashList_ofIncome()
    On Error GoTo TestFail
    'Arrange:
    Call BbLog.Clear
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act:
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "�G����", "�Z�~�i�[�Q����", "��ȃt�H�[����")
    'Assert:
    Assert.AreEqual CLng(4), selected.Count
    Debug.Print selected.ToString()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CashSelector�C���X�^���X�𐶐���SelectCashList���\�b�h���e�X�g���� - Not Like�̏ꍇ")
Private Sub Test_SelectCachList_NotLike()
    ' positiveLike�p�����[�^��false���w�肵�āu���k�u���b�N�u�K��v�����O����������^�ʐM������X�g����
    On Error GoTo TestFail:
    'Arrange
    Call BbLog.Clear
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Expense, "������", "�ʐM��", "���k�u���b�N�u�K��", False)
    'Assert
    Debug.Print selected.ToString()
    Assert.IsTrue (CLng(0) < selected.Count)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CashSelector�C���X�^���X�𐶐���SelectCashListByMatchingDescription���\�b�h���e�X�g����")
Private Sub Test_SelectCashListByMatchingDescription()
    On Error GoTo TestFail
    'Arrange:
    Call BbLog.Clear
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act:
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    Set selected = cs.SelectCashListByMatchingDescription(AccountType.Income, "���", "A���", "��� ϻ�")
    'Assert:
    Assert.AreEqual CLng(1), selected.Count
    Debug.Print selected.ToString()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("CashSelector�I�u�W�F�N�g��SelectCashList���\�b�h�𖡌�����:Income�̏ꍇ")
Private Sub TasteCashSelector_SelectCashList_ofIncome()
    'Arrange
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    'Act
    Set selected = cs.SelectCashList(AccountType.Income, "�G����", "�Z�~�i�[�Q����", "��ȃt�H�[����")
    Debug.Print selected.Count
    Debug.Print selected.ToString()
End Sub


'@TestMethod("CashSelector�I�u�W�F�N�g��SelectCashListOfIncome���\�b�h�𖡌�����")
Private Sub TasteCashSelector_SelectCashList_ofExpense()
    'Arrange
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    'Act
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Expense, "���Ɣ�", "�L���", "���")
    'Assert
    Debug.Print selected.Count
    Debug.Print selected.ToString()
End Sub



