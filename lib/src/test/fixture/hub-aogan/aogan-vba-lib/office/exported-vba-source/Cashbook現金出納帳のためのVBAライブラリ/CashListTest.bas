Attribute VB_Name = "CashListTest"
Option Explicit
Option Private Module

' TestCashList: CashList�N���X�����j�b�g�e�X�g����

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Dim wb As Workbook
Dim ws As Worksheet
Dim tbl As ListObject
Dim cb As Cashbook
Dim cs As CashSelector
    
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
    Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Set cs = New CashSelector
    Call cs.Initialize(cb)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set cb = Nothing
    Set cs = Nothing
End Sub


'@TestMethod("CashList�I�u�W�F�N�g��Count�v���p�e�B���e�X�g����")
Private Sub TestCount()
    On Error GoTo TestFail
    'Arrange:
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "�G����", "�Z�~�i�[�Q����", "��ȃt�H�[����")
    'Act:
    'Assert:
    Assert.AreEqual CLng(4), selected.Count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CashList�I�u�W�F�N�g��Items�֐����e�X�g����")
Private Sub TestItems()
    On Error GoTo TestFail
    'Arrange:
    Call BbLog.Clear
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "�G����", "�Z�~�i�[�Q����", "��ȃt�H�[����")
    'Act:
    'Assert:
    Assert.AreEqual CLng(4), selected.Count
    Dim i As Long
    For i = 1 To selected.Count
        Debug.Print selected.Items(i).ToString()
    Next i
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CashList�I�u�W�F�N�g��SumOfIncomeAmount���\�b�h���e�X�g����")
Private Sub TestSumOfIncomeAmount()
    On Error GoTo TestFail
    'Arrange:
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "�G����", "�Z�~�i�[�Q����", "��ȃt�H�[����")
    Call BbLog.Clear
    Debug.Print selected.ToString()
    'Act:
    Dim sum As Long
    sum = selected.SumOfIncomeAmount()
    'Assert:
    Assert.AreEqual CLng(56000), sum
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CashList�I�u�W�F�N�g��SumOfExpenseAmount���\�b�h���e�X�g����")
Private Sub TestSumOfExpenseAmount()
    On Error GoTo TestFail
    'Arrange:
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Expense, "���Ɣ�", "���O�q����")
    Call BbLog.Clear
    Debug.Print selected.ToString()
    'Act:
    Dim sum As Long
    sum = selected.SumOfExpenseAmount()
    'Assert:
    Assert.AreEqual CLng(2), selected.Count
    Debug.Print sum
    Assert.AreEqual CLng(540000), sum
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

