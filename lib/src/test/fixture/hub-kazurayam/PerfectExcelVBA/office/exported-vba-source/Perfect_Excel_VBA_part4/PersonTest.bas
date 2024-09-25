Attribute VB_Name = "PersonTest"
Option Explicit


Option Private Module

' 書籍「パーフェクトExcel VBA」高橋宣成
' §14 アプリケーション開発

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
    Dim datasheet As Worksheet: Set datasheet = Sheets("名簿")
    With datasheet
        Set aPerson = New Person
        aPerson.Initialize datasheet.ListObjects(1).ListRows(1).Range
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("名簿シートの1行を入力としてPersonオブジェクトを生成する")
Private Sub TestInstantiatePerson()
    logger.procedureName = "TestInstantiatePerson"
    '
    logger.Info "aPerson.Id", aPerson.ID
    logger.Info "aPerson.Name", aPerson.Name
    logger.Info "aPerson.Gender", aPerson.Gender
    logger.Info "aPerson.Birthday", aPerson.Birthday
    logger.Info "aPerson.Active", aPerson.Active
        
    Assert.AreEqual CLng(1), aPerson.ID
    Assert.AreEqual "横尾 勉", aPerson.Name
    Assert.AreEqual "男", aPerson.Gender
    Assert.AreEqual #3/30/1988#, aPerson.Birthday
    Assert.AreEqual True, aPerson.Active
End Sub

'@TestMethod("Age property")
Private Sub TestAgeProperty()
    '横尾勉さんの誕生日が3/30/1988である
    '横尾さんの今年の誕生日が何日かというと
    Dim birthdayThisYear As Date
    birthdayThisYear = DateSerial(Year(Now), Month(aPerson.Birthday), Day(aPerson.Birthday))
    
    '今日が2024/2/1つまり今年の誕生日より前なら2024 - 1988=36より1少ない35歳が正しい年齢
    '今日が2024/4/1つまり今年の誕生日より後ならば2024 - 1988=36歳が正しい年齢
    Dim expected As Long: expected = DateDiff("yyyy", aPerson.Birthday, Date)
    If Date < birthdayThisYear Then
        expected = expected - 1
    End If
    
    logger.Info "aPerson.Age", aPerson.Age
    Assert.AreEqual expected, aPerson.Age
End Sub
