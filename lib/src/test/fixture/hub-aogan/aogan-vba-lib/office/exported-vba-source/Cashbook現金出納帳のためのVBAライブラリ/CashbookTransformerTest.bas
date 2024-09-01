Attribute VB_Name = "CashbookTransformerTest"
Option Explicit

Option Private Module

' CashbookTransformerTest : CashbookTransformer�N���X�����j�b�g�e�X�g����

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
    
    Set wb = Workbooks.Open(BbUtil.ResolveExternalFilePath(ThisWorkbook, _
                                            "�����o�[���t�@�C���̃p�X", "B2"))
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

'@TestMethod("�����Ȃ��FrpUnit=*, positiveLike=True�̏ꍇ")
Private Sub Test_ByAccounts()
    On Error GoTo TestFail
    'Arrange:
    Call BbLog.Clear
    Dim cashbook_ As Cashbook
    Set cashbook_ = New Cashbook
    Call cashbook_.Initialize(tbl)
    'Act:
    Dim cbTransformer As CashbookTransformer
    Set cbTransformer = New CashbookTransformer
    Call cbTransformer.Initialize(cashbook_)
    
    Dim cashListDic As Dictionary
    Set cashListDic = cbTransformer.ByAccounts()
    
    'Assert:
    Assert.IsTrue (CLng(0) < cashListDic.Count)
    Debug.Print cashListDic.Count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("��������FrpUnit=���k�u���b�N�u�K��, positiveLike=False�̏ꍇ")
Private Sub Test_byAccounts_positiveLike_False()
    On Error GoTo TestFail
    'Arrange:
    Call BbLog.Clear
    Dim cashbook_ As Cashbook
    Set cashbook_ = New Cashbook
    Call cashbook_.Initialize(tbl)
    'Act:
    Dim cbTransformer As CashbookTransformer
    Set cbTransformer = New CashbookTransformer
    Call cbTransformer.Initialize(cashbook_)
    
    Dim cashListDic As Dictionary
    Set cashListDic = cbTransformer.ByAccounts("���k�u���b�N�u�K��", False)
    'Assert:
    Assert.IsTrue (CLng(0) < cashListDic.Count)
    Debug.Print cashListDic.Count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CashbookTransformer�N���X�𖡌�����")
Private Sub Taste_byAccounts_showCount()
    On Error GoTo TestFail
    'Arrange
    Dim cashbook_ As Cashbook: Set cashbook_ = New Cashbook
    Call cashbook_.Initialize(tbl)
    'Act
    Dim cbTransformer As CashbookTransformer: Set cbTransformer = New CashbookTransformer
    Call cbTransformer.Initialize(cashbook_)
    
    Dim cashListDic As Dictionary: Set cashListDic = cbTransformer.ByAccounts("*", True)
    
    Dim accounts As Variant: Let accounts = cashListDic.Keys
    '�L�[�������Ƀ\�[�g����
    Call BbArraySort.InsertionSort(accounts, LBound(accounts), UBound(accounts))
    
    Call BbLog.Clear
    Dim i As Long
    For i = LBound(accounts) To UBound(accounts)
        Dim account_ As String: account_ = accounts(i)
        Dim msg As String
        msg = account_ & ":" & CStr(cashListDic(account_).Count())
        Call BbLog.Info("TestCashbookTransformer", "Taste_byAccounts_showCount", msg)
    Next i
'0//:1
'����/���/A���:44
'����/���/B���:32
'����/���/C���:1
'����/�⏕��/�⏕��:1
'����/�G����/�Z�~�i�[�Q����:4
'����/�G����/�L��������:12
'����/�G����/�G����:1
'�x�o/������/���Օi��:27
'�x�o/������/�ʐM��:58
'�x�o/���Ɣ�/���O�q����:1
'�x�o/���Ɣ�/�w�p��:17
'�x�o/���Ɣ�/�L���:3
'�x�o/���Ɣ�/������:1
'�x�o/���Ɣ�/��ȃR�E���f�B�J���֌W��:1
'�x�o/���Ɣ�/�ʐM��:1
'�x�o/��c��/�������:1
'�x�o/�o����/�o������⏕:1
'�x�o/�c����/�c����:1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("FindKeysAs String���e�X�g����")
Private Sub TestFindKeysAsString()
    On Error GoTo TestFail
    'Arrange
    Dim cashbook_ As Cashbook: Set cashbook_ = New Cashbook
    Call cashbook_.Initialize(tbl)
    'Act
    Dim cbTransformer As CashbookTransformer: Set cbTransformer = New CashbookTransformer
    Call cbTransformer.Initialize(cashbook_)

    Call BbLog.Info("TestCashbookTransformer", "Test_FindKeysAsString", cbTransformer.FindKeysAsString())
'0//
'����/���/A���
'����/���/B���
'����/���/C���
'����/�⏕��/�⏕��
'����/�G����/�Z�~�i�[�Q����
'����/�G����/�L��������
'����/�G����/�G����
'�x�o/������/���Օi��
'�x�o/������/�ʐM��
'�x�o/���Ɣ�/���O�q����
'�x�o/���Ɣ�/�w�p��
'�x�o/���Ɣ�/�L���
'�x�o/���Ɣ�/������
'�x�o/���Ɣ�/��ȃR�E���f�B�J���֌W��
'�x�o/���Ɣ�/�ʐM��
'�x�o/��c��/�������
'�x�o/�o����/�o������⏕
'�x�o/�c����/�c����
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AccountsFinder�N���X�𖡌����� ���������x�񍐒P�ʂ���肵��")
Private Sub Test_byAccounts_ReportingUnit()
    'Arrange
    Dim cashbook_ As Cashbook: Set cashbook_ = New Cashbook
    Call cashbook_.Initialize(tbl)
    'Act
    Dim cbTransformer As CashbookTransformer: Set cbTransformer = New CashbookTransformer
    Call cbTransformer.Initialize(cashbook_)
    
    Dim cashListDic As Dictionary
    Set cashListDic = cbTransformer.ByAccounts(ofReportingUnit:="���k�u���b�N�u�K��")
    'Assert
    Call BbLog.Clear
    Dim accounts As Variant
    accounts = cashListDic.Keys
    Dim i As Long
    For i = LBound(accounts) To UBound(accounts)
        Dim account As String: account = accounts(i)
        Dim msg As String: msg = i & " " & account & " " & cashListDic(account).Count()
        Call BbLog.Info("TestCashbookTransformer", "Test_byAccounts_ReportingUnit", msg)
    Next i
    
'�x�o/������/�ʐM��:17
'�x�o/���Ɣ�/�w�p��:9
'�x�o/���Ɣ�/�ʐM��:1

End Sub

