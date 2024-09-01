Attribute VB_Name = "AccountTest"
Option Explicit
Option Private Module

' TestAccount : Accountオブジェクトをユニットテストする

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

'@TestMethod("AccountオブジェクトのAccountTypeAsStringプロパティをテストする")
Private Sub TestAccoutTypeAsString()
    On Error GoTo TestFail
    'Arrange:
    'Act:
    Dim expenseAccount As account: Set expenseAccount = New account
    Call expenseAccount.Initialize(AccountType.Expense, "事業費", "広報費")
    'Assert:
    Assert.AreEqual "支出", expenseAccount.AccountTypeAsString
    'Act:
    Dim incomeAccount As account: Set incomeAccount = New account
    Call incomeAccount.Initialize(AccountType.Income, "雑収入", "セミナー参加料")
    'Assert:
    Assert.AreEqual "収入", incomeAccount.AccountTypeAsString
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("支出の勘定科目の一例であるAccountオブジェクトをテストする")
Private Sub TestExpenseAccount()
    On Error GoTo TestFail
    'Arrange:
    'Act:
    Dim expenseAccount As account: Set expenseAccount = New account
    Call expenseAccount.Initialize(AccountType.Expense, "事業費", "広報費")
    'Assert:
    Assert.AreEqual AccountType.Expense, expenseAccount.AccType
    Assert.AreEqual "事業費", expenseAccount.AccountName
    Assert.AreEqual "広報費", expenseAccount.SubAccountName()
    Assert.AreEqual "支出/事業費/広報費", expenseAccount.ToString()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("収入の勘定科目の一例であるAccountオブジェクトをテストする")
Private Sub TestIncomeAccount()
    On Error GoTo TestFail
    'Arrange:
    'Act:
    Dim incomeAccount As account: Set incomeAccount = New account
    Call incomeAccount.Initialize(AccountType.Income, "雑収入", "セミナー参加料")
    'Assert:
    Assert.AreEqual AccountType.Income, incomeAccount.AccType
    Assert.AreEqual "雑収入", incomeAccount.AccountName
    Assert.AreEqual "セミナー参加料", incomeAccount.SubAccountName
    Assert.AreEqual "収入/雑収入/セミナー参加料", incomeAccount.ToString()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Stringを引数として受け取るイニシャライザ of(str)をテストする")
Private Sub TestOfString()
    On Error GoTo TestFail
    'Arrange:
    Dim incomeAccount As account: Set incomeAccount = New account
    Dim expenseAccount As account: Set expenseAccount = New account
    'Act:
    incomeAccount.Of ("収入/雑収入/セミナー参加料")
    expenseAccount.Of ("支出/事務費/通信費")
    'Assert:
    Assert.AreEqual AccountType.Income, incomeAccount.AccType
    Assert.AreEqual "雑収入", incomeAccount.AccountName
    Assert.AreEqual "セミナー参加料", incomeAccount.SubAccountName
    Assert.AreEqual "収入/雑収入/セミナー参加料", incomeAccount.ToString()
    Assert.AreEqual AccountType.Expense, expenseAccount.AccType
    Assert.AreEqual "事務費", expenseAccount.AccountName
    Assert.AreEqual "通信費", expenseAccount.SubAccountName
    Assert.AreEqual "支出/事務費/通信費", expenseAccount.ToString()
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("味見する: 支出勘定")
Private Sub TasteAccount_Expense()
    ' Expense account
    Dim accExpense As account: Set accExpense = New account
    Call accExpense.Initialize(AccountType.Expense, "事業費", "広報費")
    Debug.Print "accExpense.accType: " & accExpense.AccType
    Debug.Print "accExpense.AccountName: " & accExpense.AccountName
    Debug.Print "accExpense.SubAccountName: " & accExpense.SubAccountName
    Debug.Print "accExpense.ToString(): " & accExpense.ToString()
End Sub


'@TestMethod("味見する: 収入勘定")
Private Sub TasteAccount_Income()
    ' Income account
    Dim accIncome As account: Set accIncome = New account
    Call accIncome.Initialize(AccountType.Income, "雑収入", "セミナー参加料")
    Debug.Print "accIncome.accType: " & accIncome.AccType
    Debug.Print "accIncome.AccountName: " & accIncome.AccountName
    Debug.Print "accIncome.SubAccountName: " & accIncome.SubAccountName
    Debug.Print "accIncome.ToString(): " & accIncome.ToString()
End Sub


