Attribute VB_Name = "TestCash"
Option Explicit
Option Private Module

' TestCash : Cash�N���X�����j�b�g�e�X�g����

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
    
    Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Set ws = wb.Worksheets("�����o�[��")
    Set tbl = ws.ListObjects("CashbookTable1")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
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


'@TestMethod("�����o�[�����[�N�V�[�g�̃e�[�u��1�ɃA�N�Z�X����e�X�g")
Private Sub TestAccessTable1()
    On Error GoTo TestFail
    'Arrange:
    'Act:
    'Assert:
    Assert.isTrue tbl.ListRows.Count > 0, "�e�[�u��������ۂ�������Ȃ�"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub




'@TestMethod("Cash�I�u�W�F�N�g�����v���p�e�B��ǂ݂����e�X�g�@�x�o�̏ꍇ")
Private Sub TestExpense()
    On Error GoTo TestFail
    'Arrange:
    Dim data As Range: Set data = tbl.ListRows(10).Range '�x�o�̈��
    Dim cs As Cash: Set cs = New Cash
    'Act:
    Call cs.Initialize(data)
    'Assert:
        'Debug.Print "cs.ReceiptNo is " + KzVarTypeAsString(cs.ReceiptNo)
        'Debug.Print "5 is " + KzVarTypeAsString(5)
        'Debug.Print "CLng(5) is " + KzVarTypeAsString(CLng(5))
    Assert.AreEqual CLng(5), cs.ReceiptNo     '�̎���No.
    Assert.AreEqual CLng(4), cs.YY            '�N�x�i�ߘa�j
    Assert.AreEqual CLng(5), cs.MM            '��
    Assert.AreEqual CLng(24), cs.DD           '��
    Assert.AreEqual "", cs.incomeAccount            '�����Ȗ�
    Assert.AreEqual "", cs.IncomeSubAccount         '�����⏕�Ȗ�
    Assert.AreEqual "���Ɣ�", cs.expenseAccount     '�x�o�Ȗ�
    Assert.AreEqual "�L���", cs.ExpenseSubAccount  '�x�o�⏕�Ȗ�
    Assert.AreEqual "���", cs.ReportingUnit        '���x�񍐒P��
    Assert.AreEqual "����73�������@�ʔŃ��f�B�A", cs.Description   '�K�p
    Assert.AreEqual CLng(0), cs.IncomeAmount              '�ؕ����z
    Assert.AreEqual CLng(495000), cs.ExpenseAmount        '�ݕ����z
    Assert.AreEqual "�x�o/���Ɣ�/�L���", cs.itsAccount.ToString
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Cash�I�u�W�F�N�g�����v���p�e�B��ǂ݂����e�X�g�@�����̏ꍇ")
Private Sub TestIncome()
    On Error GoTo TestFail
    'Arrange:
    Dim data As Range: Set data = tbl.ListRows(12).Range '�����̈��
    '    4   5   28  �G����  �Z�~�i�[�Q����          ��ȃt�H�[����  �϶޲��ݲ��@�Z�~�i�[��u��   2,000
    Dim cs As Cash: Set cs = New Cash
    'Act:
    Call cs.Initialize(data)
    'Assert:
    Assert.AreEqual CLng(0), cs.ReceiptNo     '�̎���No.
    Assert.AreEqual CLng(4), cs.YY            '�N�x�i�ߘa�j
    Assert.AreEqual CLng(5), cs.MM            '��
    Assert.AreEqual CLng(28), cs.DD           '��
    Assert.AreEqual "�G����", cs.incomeAccount            '�����Ȗ�
    Assert.AreEqual "�Z�~�i�[�Q����", cs.IncomeSubAccount         '�����⏕�Ȗ�
    Assert.AreEqual "", cs.expenseAccount     '�x�o�Ȗ�
    Assert.AreEqual "", cs.ExpenseSubAccount  '�x�o�⏕�Ȗ�
    Assert.AreEqual "��ȃt�H�[����", cs.ReportingUnit        '���x�񍐒P��
    Assert.AreEqual "�϶޲��ݲ��@�Z�~�i�[��u��", cs.Description   '�K�p
    Assert.AreEqual CLng(2000), cs.IncomeAmount            '�ؕ����z
    Assert.AreEqual CLng(0), cs.ExpenseAmount        '�ݕ����z
    
    Assert.AreEqual "����/�G����/�Z�~�i�[�Q����", cs.itsAccount.ToString()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Cash�N���X��ColumnHeader���\�b�h��ToString���\�b�h���e�X�g����")
Private Sub TestToString()
    On Error GoTo TestFail
    'Arrange:
    Call KzCls
    Dim data As Range: Set data = tbl.ListRows(12).Range '�����̈��
    Dim cs As Cash: Set cs = New Cash
    Call cs.Initialize(data)
    'Act:
    Dim ch As String: ch = cs.ColumnHeader()
    Dim s As String: s = cs.ToString()
    Debug.Print ch
    Debug.Print s
    'Assert:
    Assert.isTrue InStr(ch, "�̎���") <> 0, "ColumnHeader"
    Assert.isTrue InStr(s, " ") <> 0, "�̎���No."
    Assert.isTrue InStr(s, "4") <> 0, "�N"
    Assert.isTrue InStr(s, "5") <> 0, "��"
    Assert.isTrue InStr(s, "28") <> 0, "��"
    Assert.isTrue InStr(s, "�G����") <> 0, "�����Ȗ�"
    Assert.isTrue InStr(s, "�Z�~�i�[�Q����") <> 0, "�����⏕�Ȗ�"
    Assert.isTrue InStr(s, " ") <> 0, "�x�o�Ȗ�"
    Assert.isTrue InStr(s, " ") <> 0, "�x�o�⏕�Ȗ�"
    Assert.isTrue InStr(s, "��ȃt�H�[����") <> 0, "���x�񍐒P��"
    Assert.isTrue InStr(s, "�϶޲��ݲ��@�Z�~�i�[��u��") <> 0, "�K�p"
    Assert.isTrue InStr(s, "2000") <> 0, "�ؕ����z"
    Assert.isTrue InStr(s, " ") <> 0, "�ݕ����z"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Cash�N���X��ToDate���\�b�h���e�X�g����")
Private Sub TestToDate()
    On Error GoTo TestFail
    'Arrange:
    Dim data As Range: Set data = tbl.ListRows(12).Range '�����̈��
    Dim cs As Cash: Set cs = New Cash
    Call cs.Initialize(data)
    'Act:
    Debug.Print cs.ToDate()
    'Assert:
    Assert.AreEqual #5/28/2022#, cs.ToDate()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
