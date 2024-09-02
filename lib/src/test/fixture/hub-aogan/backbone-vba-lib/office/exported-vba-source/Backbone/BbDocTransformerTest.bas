Attribute VB_Name = "BbDocTransformerTest"
Option Explicit
Option Private Module

' DocTransformer�N���X�����s���ē�����m�F���܂��ARubberduck�ɂ�鎩�����e�X�g�ł͂Ȃ��蓮�Ŏ��s����
' ���͂Ƃ��Ă�KeyValuePair���R�^����A���̂R�̓R�[�h�̂Ȃ��ɌŒ�l�Ƃ��ď����Ă���

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

'@TestMethod("DocTransformer�N���X�𖡌�����")
Private Sub TestDocTransformer()
    On Error GoTo TestFail
    
    Call BbLog.Clear
    
    ' ���͂ƂȂ�e���v���[�g�Ƃ��Ă�Word�t�@�C���̃p�X
    Dim templateFile As String
    templateFile = BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B2")
    
    ' �o�͂����Word�t�@�C�����i�[�����Ƃ��Ẵt�H���_�̃p�X
    Dim resultFolder As String
    
    ' OneDrive�ŗL�̌`���̃t�@�C���p�X�Ȃ�Ε��ʂ̃t�@�C���p�X�ɕϊ�����
    resultFolder = BbFile.ToLocalFilePath(ThisWorkbook.path) & "\" & "build\TestDocTransformer"
    
    ' �o�͐�t�H���_���������܂����݂��Ă��Ȃ���������
    Call BbFile.EnsureFolders(resultFolder)
    
    ' �v���[�X�z���_�[�̖��O�ƒl�̑g
    Dim dict As Dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "����", "�W�����E�h�[�G"
    dict.Add "ID", "John Doe"
    dict.Add "PW", "ThisIsNotAPassword"
    
    ' �u�����v�Ɋ�Â��ďo��Word�t�@�C���̖��O�����߂�
    Dim resultName As String: resultName = dict("����") & ".docx"
        
    ' ����Word�t�@�C���Əo��Word�t�@�C���̃p�X�����肵��
    Dim t As String: t = templateFile
    Dim r As String: r = resultFolder & "\" & resultName
        
    Debug.Print "�e���v���[�g=" & t
    Debug.Print "�o�̓t�@�C��=" & r
    
    ' Word�h�L�������g��ϊ�����N���X������������
    ' DocTransformer�C���X�^���X�𐶐�����
    Dim dt As BbDocTransformer
    Set dt = BbDocTransformerFactory.CreateDocTransformer
    
    
    ' Word�A�v���P�[�V�����̃C���X�^���X�𐶐�
    Dim wordApp As Word.Application: Set wordApp = CreateObject("Word.application")
    Call dt.Initialize(wordApp)
    ' ���̓t�@�C���ƒu���f�[�^�Əo�̓t�@�C�����w�肵��Word�h�L�������g��ϊ����o�͂���
    Call dt.Transform(t, dict, r)
    
    Debug.Print "�ϊ����������s���܂���"
    
    ' =========================================================================
    ' ��n��
    Set dt = Nothing
    Set dict = Nothing
    Set wordApp = Nothing

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


