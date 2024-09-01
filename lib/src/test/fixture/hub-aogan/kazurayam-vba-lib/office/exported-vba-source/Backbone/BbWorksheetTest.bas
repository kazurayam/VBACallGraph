Attribute VB_Name = "BbWorksheetTest"
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


'@TestMethod("IsWorksheetPresentInWorkbookをテストする")
Private Sub Test_IsWorksheetPresentInWorkbook()
    'Assert:
    Assert.IsTrue BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, "テストデータ")
    Assert.IsFalse BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, "No Such Worksheet")
End Sub

'@TestMethod("CreateWorksheetInWorkbookをテストする")
Private Sub Test_CreateWorksheetInWorkbook()
    'Arrange
    ' カレントのWorkbookにtempという名前のワークシートがもしもあったら削除する
    Dim wsName As String: wsName = "temp"
    Dim r As Boolean
    If BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, wsName) Then
        r = BbWorksheet.DeleteWorksheetInWorkbook(ThisWorkbook, wsName)
    End If
    'Act:
    ' tempワークシートを追加する
    r = BbWorksheet.CreateWorksheetInWorkbook(ThisWorkbook, wsName)
    'Assert
    ' tempワークシートができていることを確認する
    Assert.IsTrue BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, wsName)
    
End Sub

'@TestMethod("DeleteWorksheetInWorkbookをテストする")
Private Sub Test_DeleteWorksheetInWorkbook()
    'Arrange
    ' カレントのWorkbookにtempという名前のワークシートがもしもあったら削除する
    Dim wsName As String: wsName = "temp"
    Dim r As Boolean
    If BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, wsName) Then
        r = BbWorksheet.DeleteWorksheetInWorkbook(ThisWorkbook, wsName)
    End If
    ' カレントのWorkbookに一時的なワークシートを挿入する、
    ' シートの名前を temp とする
    Worksheets.Add(After:=Worksheets(Worksheets.Count)) _
        .Name = wsName
    'Act:
    ' 挿入したワークシートを削除する
    r = BbWorksheet.DeleteWorksheetInWorkbook(ThisWorkbook, wsName)
    ' 一時的に挿入したワークシートがもはや存在しないことを確認するA
    Assert.IsFalse BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, wsName)
End Sub



'@TestMethod("FetchWorksheetFromWorkbookをユニットテストする")
Private Sub Test_FetchWorksheetFromWorkbook()
    On Error GoTo TestFail
    'Arrange
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheetName As String: sourceSheetName = "テストデータ"
    Dim targetWorkbook As Workbook: Set targetWorkbook = sourceWorkbook
    Dim targetSheetName As String: targetSheetName = "copy"
    'Act
    Call BbWorksheet.FetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
    'Assert
    Assert.IsTrue BbWorksheet.IsWorksheetPresentInWorkbook(targetWorkbook, "copy")
    'TearDown
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(targetSheetName).Delete
    Application.DisplayAlerts = True
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("FetchWorksheetFromWorkbookがErrを投げる場合")
Private Sub Test_FetchWorksheetFromWorkbook_shouldThrowErr()
    On Error GoTo TestFail
    Call BbLog.Clear
    'Arrange
    '入力と出力に同じワークブックの同じシートが指定したらエラーになるよ
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheetName As String: sourceSheetName = "テストデータ"
    Dim targetWorkbook As Workbook: Set targetWorkbook = sourceWorkbook
    Dim targetSheetName As String: targetSheetName = "テストデータ"
    'Act
    Call BbWorksheet.FetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
    'Assert
    Assert.IsFalse BbWorksheet.IsWorksheetPresentInWorkbook(targetWorkbook, "copy")
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

