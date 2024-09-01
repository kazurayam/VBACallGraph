Attribute VB_Name = "FactoriesTest"
Option Explicit
Option Private Module

'TestFactories: Factories���W���[�������j�b�g�e�X�g����

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Private wb As Workbook
Private ws As Worksheet
Private tbl As ListObject
Private cb As Cashbook

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    '
    Set wb = Workbooks.Open(BbUtil.ResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Set ws = wb.Worksheets("�����o�[��")
    Set tbl = ws.ListObjects("CashbookTable1")
    Set cb = New Cashbook
    Call cb.Initialize(tbl)
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    'Teardown
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb.Close
    Set wb = Nothing
    Set ws = Nothing
    Set tbl = Nothing
    Set cb = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("CreateCashbook�֐����e�X�g����")
Private Sub TestCreateCashbook()
    'Act:
    Dim cbx As Cashbook: Set cbx = CreateCashbook(wb, "�����o�[��", "CashbookTable1")
    'Assert:
    Assert.IsTrue (CLng(0) < cbx.Count)
End Sub


'@TestMethod("CreateCashbookTransformer�֐����e�X�g����")
Private Sub TestCreateCashbookTransformer()
    'Act
    Dim cbTransformer As CashbookTransformer
    Set cbTransformer = CreateCashbookTransformer(cb)
    Call BbLog.Clear
    Debug.Print cbTransformer.FindKeysAsString
    '
End Sub


'@TestMethod("CreateCashSelector�֐����e�X�g����")
Private Sub TestCreateCashSelector()
    Dim cs As CashSelector: Set cs = CreateCashSelector(cb)
    Call BbLog.Clear
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "�G����", "�Z�~�i�[�Q����", "��ȃt�H�[����")
    'Assert:
    Debug.Print selected.Count
    Assert.AreEqual CLng(4), selected.Count
End Sub

'@TestMethod("CreateEmptyCashList�֐����e�X�g����")
Private Sub TestCreateEmptyCashList()
    Call BbLog.Clear
    Dim cl As CashList
    Set cl = CreateEmptyCashList()
    Assert.AreEqual CLng(0), cl.Count
End Sub
