Attribute VB_Name = "AccountTest"
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
    Dim expenseAccount As account: Set expenseAccount = New account
    Call expenseAccount.Initialize(AccountType.Expense, "���Ɣ�", "�L���")
    'Assert:
    Assert.AreEqual "�x�o", expenseAccount.AccountTypeAsString
    'Act:
    Dim incomeAccount As account: Set incomeAccount = New account
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
    Dim expenseAccount As account: Set expenseAccount = New account
    Call expenseAccount.Initialize(AccountType.Expense, "���Ɣ�", "�L���")
    'Assert:
    Assert.AreEqual AccountType.Expense, expenseAccount.AccType
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
    Dim incomeAccount As account: Set incomeAccount = New account
    Call incomeAccount.Initialize(AccountType.Income, "�G����", "�Z�~�i�[�Q����")
    'Assert:
    Assert.AreEqual AccountType.Income, incomeAccount.AccType
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
    Dim incomeAccount As account: Set incomeAccount = New account
    Dim expenseAccount As account: Set expenseAccount = New account
    'Act:
    incomeAccount.Of ("����/�G����/�Z�~�i�[�Q����")
    expenseAccount.Of ("�x�o/������/�ʐM��")
    'Assert:
    Assert.AreEqual AccountType.Income, incomeAccount.AccType
    Assert.AreEqual "�G����", incomeAccount.AccountName
    Assert.AreEqual "�Z�~�i�[�Q����", incomeAccount.SubAccountName
    Assert.AreEqual "����/�G����/�Z�~�i�[�Q����", incomeAccount.ToString()
    Assert.AreEqual AccountType.Expense, expenseAccount.AccType
    Assert.AreEqual "������", expenseAccount.AccountName
    Assert.AreEqual "�ʐM��", expenseAccount.SubAccountName
    Assert.AreEqual "�x�o/������/�ʐM��", expenseAccount.ToString()
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("��������: �x�o����")
Private Sub TasteAccount_Expense()
    ' Expense account
    Dim accExpense As account: Set accExpense = New account
    Call accExpense.Initialize(AccountType.Expense, "���Ɣ�", "�L���")
    Debug.Print "accExpense.accType: " & accExpense.AccType
    Debug.Print "accExpense.AccountName: " & accExpense.AccountName
    Debug.Print "accExpense.SubAccountName: " & accExpense.SubAccountName
    Debug.Print "accExpense.ToString(): " & accExpense.ToString()
End Sub


'@TestMethod("��������: ��������")
Private Sub TasteAccount_Income()
    ' Income account
    Dim accIncome As account: Set accIncome = New account
    Call accIncome.Initialize(AccountType.Income, "�G����", "�Z�~�i�[�Q����")
    Debug.Print "accIncome.accType: " & accIncome.AccType
    Debug.Print "accIncome.AccountName: " & accIncome.AccountName
    Debug.Print "accIncome.SubAccountName: " & accIncome.SubAccountName
    Debug.Print "accIncome.ToString(): " & accIncome.ToString()
End Sub


