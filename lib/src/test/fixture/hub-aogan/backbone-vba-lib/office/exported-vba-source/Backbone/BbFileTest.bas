Attribute VB_Name = "BbFileTest"
Option Explicit
Option Private Module

'BbFileTest

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

'@TestMethod("KzIsAbsoluthPath関数")
Private Sub Test_IsAbsolutePath()
    Dim r1 As Boolean: r1 = BbFile.IsAbsolutePath("C:\somepath")
    Debug.Print ("r1: " & r1)
    Assert.IsTrue r1
    
    Dim r2 As Boolean: r2 = BbFile.IsAbsolutePath("\somepath")
    Debug.Print ("r2: " & r2)
    Assert.IsTrue r2
    
    Dim r3 As Boolean: r3 = BbFile.IsAbsolutePath("..\somepath")
    Debug.Print ("r3: " & r3)
    Assert.IsFalse r3
End Sub



'@TestMethod("AbsolutifyPath関数に相対パスを与えるケース")
Private Sub Test_AbsolutifyPath_relative()
    'ファイルの相対パスを与えたら絶対パスに変換されること
    Call BbLog.Clear
    'Arrange:
    Dim base As String: base = ThisWorkbook.path
    Const givenPath = ".\data\Book1.xlsx"
    'Act:
    Dim absPath As String: absPath = BbFile.AbsolutifyPath(base, givenPath)
    'Assert:
    Debug.Print "base : " & base
    Debug.Print "given: " & givenPath
    Debug.Print "abs  : " & absPath
    Assert.IsTrue absPath Like "C:\*"         ' 絶対パスならC:\で始まって
    Assert.IsTrue absPath Like "*\Book1.xlsx" ' \Book1.xlsxで終わるはず
End Sub

'@TestMethod("AbsolutifyPath関数に絶対パスを与えるケース")
Private Sub Test_AbsolutifyPath_absolute()
    'ファイルの絶対パスを与えたら変換せずにそのまま返すこと
    Call BbLog.Clear
    'Arrange
    Dim base As String: base = ThisWorkbook.path
    Const givenPath = "C:\Users\someone\tmp\Book1.xlsx"
    'Act
    Dim absPath As String: absPath = BbFile.AbsolutifyPath(base, givenPath)
    'Asserth
    Debug.Print "base : " & base
    Debug.Print "given: " & givenPath
    Debug.Print "abs  : " & absPath
    Assert.IsTrue absPath Like givenPath   '絶対パスが与えられたらそのまま返されるはず
End Sub


'@TestMethod("ToLocalFilePath()をユニットテストする")
Private Sub Test_ToLocalFilePath()
    'ToLocalFilePathはOneDriveにマッピングされてhttps://で始まるURLに対応するファイルをC:\で始まるローカルファイルのパスの文字列に変換する
    'Arrange:
    Dim Source As String: Source = "https://d.docs.live.net/c5960fe753e170b9/デスクトップ/Excel-Word-VBA"
    Dim expect As String: expect = "C:\Users\aogan\OneDrive\デスクトップ\Excel-Word-VBA"
    Dim actual As String
    'Act:
    actual = BbFile.ToLocalFilePath(Source)
    'Assert
    Debug.Print "source:" & vbTab; Chr(34); Source; Chr(34)
    Debug.Print "expect:" & vbTab; Chr(34); expect; Chr(34)
    Debug.Print "actual:" & vbTab; Chr(34); actual; Chr(34)
    Assert.IsTrue Len(actual) > 0
    Assert.IsTrue StrComp(expect, actual) = 0
End Sub


'@TestMethod("CreateFolder関数をユニットテストする")
Private Sub Test_CreateFolder()
    'ユーザのHomeフォルダの下の "OneDriver\ドキュメント" フォルダの下にtmpフォルダを作る
    Dim wshell As Object: Set wshell = CreateObject("Wscript.Shell")
    Dim docsPath As String: docsPath = wshell.SpecialFolders("MyDocuments")
    'CreateFolder(p)は引数として指定されたpをフォルダのパスとみなしてそのフォルダを作る。
    'pがまだ無ければ新しく作る。pがすでにあったらなにもしない。
    'pの親フォルダがまだ無ければエラー。親フォルダを作るにはEnsureFolder関数を使え。
    Dim p As String: p = docsPath & "\" & "tmp"
    BbFile.CreateFolder (p)
    Assert.IsTrue BbFile.PathExists(p)
    BbFile.DeleteFolder (p)
End Sub



'@TestMethod("EnsureFolders関数をユニットテストする")
Private Sub Test_EnsureFolders()
    'EnsureFolders(p)はフォルダを作る。pの親フォルダが無かったらその祖先にまで遡って作る。
    Dim wshell As Object: Set wshell = CreateObject("Wscript.Shell")
    Dim docsPath As String: docsPath = wshell.SpecialFolders("MyDocuments")
    Dim p As String: p = docsPath & "\build\tmp\testOutput"
    BbFile.EnsureFolders (p)
    Assert.IsTrue BbFile.PathExists(p)
    BbFile.DeleteFolder (p)
End Sub




'@TestMethod("PathExists関数をユニットテストする")
Private Sub Test_PathExists()
    Dim wshell As Object: Set wshell = CreateObject("Wscript.Shell")
    Dim docsPath As String: docsPath = wshell.SpecialFolders("MyDocuments")
    Assert.IsTrue BbFile.PathExists(docsPath)
End Sub


'@TestMethod("WriteTextIntoFile関数とDeleteFile関数をテストする")
Private Sub Test_WriteTextIntoFile_and_DeleteFile()
    'Arrange:
    Dim wshell As Object: Set wshell = CreateObject("Wscript.Shell")
    Dim docsPath As String: docsPath = wshell.SpecialFolders("MyDocuments")
    Dim folder As String: folder = docsPath & "\build"
    Dim file As String: file = folder & "\hello.txt"
    'Act:
    Call BbFile.WriteTextIntoFile("Hello, world", file)
    'Assert:
    Debug.Assert BbFile.PathExists(file)
    'TearDown
    BbFile.DeleteFile (file)
End Sub




