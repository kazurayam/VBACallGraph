Attribute VB_Name = "TestAccount"
Option Explicit
Option Private Module

' TestAccount : Account�I�u�W�F�N�g�����j�b�g�e�X�g����

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
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
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Account�I�u�W�F�N�g��AccountTypeAsString�v���p�e�B���e�X�g����")
Private Sub TestAccoutTypeAsString()
    On Error GoTo TestFail
    'Arrange:
    'Act:
    Dim expenseAccount As Account: Set expenseAccount = New Account
    Call expenseAccount.Initialize(AccountType.Expense, "���Ɣ�", "�L���")
    'Assert:
    Assert.AreEqual "�x�o", expenseAccount.AccountTypeAsString
    'Act:
    Dim incomeAccount As Account: Set incomeAccount = New Account
    Call incomeAccount.Initialize(AccountType.Income, "�G����", "�Z�~�i�[�Q����")
    'Assert:
    Assert.AreEqual "����", incomeAccount.AccountTypeAsString
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("�x�o�̊���Ȗڂ̈��ł���Account�I�u�W�F�N�g���e�X�g����")
Private Sub TestExpenseAccount()
    On Error GoTo TestFail
    'Arrange:
    'Act:
    Dim expenseAccount As Account: Set expenseAccount = New Account
    Call expenseAccount.Initialize(AccountType.Expense, "���Ɣ�", "�L���")
    'Assert:
    Assert.AreEqual AccountType.Expense, expenseAccount.accType
    Assert.AreEqual "���Ɣ�", expenseAccount.AccountName
    Assert.AreEqual "�L���", expenseAccount.SubAccountName()
    Assert.AreEqual "�x�o/���Ɣ�/�L���", expenseAccount.ToString()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("�����̊���Ȗڂ̈��ł���Account�I�u�W�F�N�g���e�X�g����")
Private Sub TestIncomeAccount()
    On Error GoTo TestFail
    'Arrange:
    'Act:
    Dim incomeAccount As Account: Set incomeAccount = New Account
    Call incomeAccount.Initialize(AccountType.Income, "�G����", "�Z�~�i�[�Q����")
    'Assert:
    Assert.AreEqual AccountType.Income, incomeAccount.accType
    Assert.AreEqual "�G����", incomeAccount.AccountName
    Assert.AreEqual "�Z�~�i�[�Q����", incomeAccount.SubAccountName
    Assert.AreEqual "����/�G����/�Z�~�i�[�Q����", incomeAccount.ToString()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("String�������Ƃ��Ď󂯎��C�j�V�����C�U of(str)���e�X�g����")
Private Sub TestOfString()
    On Error GoTo TestFail
    'Arrange:
    Dim incomeAccount As Account: Set incomeAccount = New Account
    Dim expenseAccount As Account: Set expenseAccount = New Account
    'Act:
    incomeAccount.of ("����/�G����/�Z�~�i�[�Q����")
    expenseAccount.of ("�x�o/������/�ʐM��")
    'Assert:
    Assert.AreEqual AccountType.Income, incomeAccount.accType
    Assert.AreEqual "�G����", incomeAccount.AccountName
    Assert.AreEqual "�Z�~�i�[�Q����", incomeAccount.SubAccountName
    Assert.AreEqual "����/�G����/�Z�~�i�[�Q����", incomeAccount.ToString()
    Assert.AreEqual AccountType.Expense, expenseAccount.accType
    Assert.AreEqual "������", expenseAccount.AccountName
    Assert.AreEqual "�ʐM��", expenseAccount.SubAccountName
    Assert.AreEqual "�x�o/������/�ʐM��", expenseAccount.ToString()
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

