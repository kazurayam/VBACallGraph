Attribute VB_Name = "TestKzWorksheet"
Option Explicit
Option Private Module

'Kzモジュールに書かれたPublicなSubやFunctionをRubberduckを使ってユニットテストする

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


'@TestMethod("KzIsWorksheetPresentInWorkbookをテストする")
Private Sub Test_KzIsWorksheetPresentInWorkbook()
    'Assert:
    Assert.IsTrue KzIsWorksheetPresentInWorkbook(ThisWorkbook, "テストデータ")
    Assert.IsFalse KzIsWorksheetPresentInWorkbook(ThisWorkbook, "No Such Worksheet")
End Sub

'@TestMethod("KzCreateWorksheetInWorkbookをテストする")
Private Sub Test_KzCreateWorksheetInWorkbook()
    'Arrange
    ' カレントのWorkbookにtempという名前のワークシートがもしもあったら削除する
    Dim wsName As String: wsName = "temp"
    Dim r As Boolean
    If KzIsWorksheetPresentInWorkbook(ThisWorkbook, wsName) Then
        r = KzDeleteWorksheetInWorkbook(ThisWorkbook, wsName)
    End If
    'Act:
    ' tempワークシートを追加する
    r = KzCreateWorksheetInWorkbook(ThisWorkbook, wsName)
    'Assert
    ' tempワークシートができていることを確認する
    Assert.IsTrue KzIsWorksheetPresentInWorkbook(ThisWorkbook, wsName)
End Sub

'@TestMethod("KzDeleteWorksheetInWorkbookをテストする")
Private Sub Test_KzDeleteWorksheetInWorkbook()
    'Arrange
    ' カレントのWorkbookに一時的なワークシートを挿入する、
    ' シートの名前を temp とする
    Dim wsName As String: wsName = "temp"
    Worksheets.Add(After:=Worksheets(Worksheets.Count)) _
        .Name = wsName
    'Act:
    ' 挿入したワークシートを削除する
    Dim r As Boolean
    r = KzDeleteWorksheetInWorkbook(ThisWorkbook, wsName)
    ' 一時的に挿入したワークシートがもはや存在しないことを確認するA
    Assert.IsFalse KzIsWorksheetPresentInWorkbook(ThisWorkbook, wsName)
End Sub



'@TestMethod("KzFetchWorksheetFromWorkbookをユニットテストする")
Private Sub Test_KzFetchWorksheetFromWorkbook()
    On Error GoTo TestFail
    'Arrange
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheetName As String: sourceSheetName = "テストデータ"
    Dim targetWorkbook As Workbook: Set targetWorkbook = sourceWorkbook
    Dim targetSheetName As String: targetSheetName = "copy"
    'Act
    Call KzFetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
    'Assert
    Assert.IsTrue KzIsWorksheetPresentInWorkbook(targetWorkbook, "copy")
    'TearDown
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(targetSheetName).Delete
    Application.DisplayAlerts = True
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("KzFetchWorksheetFromWorkbookがErrを投げる場合")
Private Sub Test_KzFetchWorksheetFromWorkbook_shouldThrowErr()
    On Error GoTo TestFail
    KzUtil.KzCls
    'Arrange
    '入力と出力に同じワークブックの同じシートが指定したらエラーになるよ
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheetName As String: sourceSheetName = "テストデータ"
    Dim targetWorkbook As Workbook: Set targetWorkbook = sourceWorkbook
    Dim targetSheetName As String: targetSheetName = "テストデータ"
    'Act
    Call KzFetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
    'Assert
    Assert.IsFalse KzIsWorksheetPresentInWorkbook(targetWorkbook, "copy")
    'TearDown
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(targetSheetName).Delete
    Application.DisplayAlerts = True
TestExit:
    Exit Sub
TestFail:
    'Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    'Errがraiseされるはずとわかっているので、何もしないで、静かに終了するべし
End Sub

