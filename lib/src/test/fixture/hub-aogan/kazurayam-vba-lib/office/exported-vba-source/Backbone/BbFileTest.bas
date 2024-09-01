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

'@TestMethod("KzIsAbsoluthPath�֐�")
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



'@TestMethod("AbsolutifyPath�֐��ɑ��΃p�X��^����P�[�X")
Private Sub Test_AbsolutifyPath_relative()
    '�t�@�C���̑��΃p�X��^�������΃p�X�ɕϊ�����邱��
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
    Assert.IsTrue absPath Like "C:\*"         ' ��΃p�X�Ȃ�C:\�Ŏn�܂���
    Assert.IsTrue absPath Like "*\Book1.xlsx" ' \Book1.xlsx�ŏI���͂�
End Sub

'@TestMethod("AbsolutifyPath�֐��ɐ�΃p�X��^����P�[�X")
Private Sub Test_AbsolutifyPath_absolute()
    '�t�@�C���̐�΃p�X��^������ϊ������ɂ��̂܂ܕԂ�����
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
    Assert.IsTrue absPath Like givenPath   '��΃p�X���^����ꂽ�炻�̂܂ܕԂ����͂�
End Sub


'@TestMethod("ToLocalFilePath()�����j�b�g�e�X�g����")
Private Sub Test_ToLocalFilePath()
    'ToLocalFilePath��OneDrive�Ƀ}�b�s���O�����https://�Ŏn�܂�URL�ɑΉ�����t�@�C����C:\�Ŏn�܂郍�[�J���t�@�C���̃p�X�̕�����ɕϊ�����
    'Arrange:
    Dim Source As String: Source = "https://d.docs.live.net/c5960fe753e170b9/�f�X�N�g�b�v/Excel-Word-VBA"
    Dim expect As String: expect = "C:\Users\aogan\OneDrive\�f�X�N�g�b�v\Excel-Word-VBA"
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


'@TestMethod("CreateFolder�֐������j�b�g�e�X�g����")
Private Sub Test_CreateFolder()
    '���[�U��Home�t�H���_�̉��� "OneDriver\�h�L�������g" �t�H���_�̉���tmp�t�H���_�����
    Dim wshell As Object: Set wshell = CreateObject("Wscript.Shell")
    Dim docsPath As String: docsPath = wshell.SpecialFolders("MyDocuments")
    'CreateFolder(p)�͈����Ƃ��Ďw�肳�ꂽp���t�H���_�̃p�X�Ƃ݂Ȃ��Ă��̃t�H���_�����B
    'p���܂�������ΐV�������Bp�����łɂ�������Ȃɂ����Ȃ��B
    'p�̐e�t�H���_���܂�������΃G���[�B�e�t�H���_�����ɂ�EnsureFolder�֐����g���B
    Dim p As String: p = docsPath & "\" & "tmp"
    BbFile.CreateFolder (p)
    Assert.IsTrue BbFile.PathExists(p)
    BbFile.DeleteFolder (p)
End Sub



'@TestMethod("EnsureFolders�֐������j�b�g�e�X�g����")
Private Sub Test_EnsureFolders()
    'EnsureFolders(p)�̓t�H���_�����Bp�̐e�t�H���_�����������炻�̑c��ɂ܂ők���č��B
    Dim wshell As Object: Set wshell = CreateObject("Wscript.Shell")
    Dim docsPath As String: docsPath = wshell.SpecialFolders("MyDocuments")
    Dim p As String: p = docsPath & "\build\tmp\testOutput"
    BbFile.EnsureFolders (p)
    Assert.IsTrue BbFile.PathExists(p)
    BbFile.DeleteFolder (p)
End Sub




'@TestMethod("PathExists�֐������j�b�g�e�X�g����")
Private Sub Test_PathExists()
    Dim wshell As Object: Set wshell = CreateObject("Wscript.Shell")
    Dim docsPath As String: docsPath = wshell.SpecialFolders("MyDocuments")
    Assert.IsTrue BbFile.PathExists(docsPath)
End Sub


'@TestMethod("WriteTextIntoFile�֐���DeleteFile�֐����e�X�g����")
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




