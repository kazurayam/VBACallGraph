Attribute VB_Name = "���[���̂��肢���쐬"
Option Explicit


' �e���v���[�g�ł���Word���������~���Ƃ��ĉ���ʂɃp�[�\�i���C�Y����Word�����𐶐�����B

' ���[�N�V�[�g�iR4�N�x�j�ɏ����ꂽ������납�����l�̎����Ǝ��i��ǂݎ��A
' �������őI�ʂ��������ŁA�e���v���[�g���̃v���[�X�z���_�[�i���Ƃ��� ${����}�j����̓I��
' �����ɒu�����āA�K�؂ȃt�@�C���������肵�āA�o�͂���B

Public Sub Main()

    ' �C�~�f�B�G�C�g�E�E�C���h�E�������B
    ' ����̎��s��Debug.Print���o�͂��郁�b�Z�[�W�����₷�����邽�߁B
    Call KzUtil.KzCls
    
    Debug.Print ("���[���̂��肢letter���쐬���܂�")
    
    ' �������Excel�t�@�C���̃p�X
    Dim memberFile As String: memberFile = _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B2")
    Debug.Print ("�������: " & memberFile)
    
    ' �e���v���[�g�Ƃ��Ă�Word�t�@�C���̃p�X
    Dim templateFile As String: templateFile = _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B3")
    Debug.Print ("�e���v���[�g: " & templateFile)
    
    ' �o�͐�t�H���_�̃p�X
    Dim outDir As String: outDir = KzFile.KzAbsolutifyPath( _
        ThisWorkbook.Path, _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B4"))
    Debug.Print ("�o�̓t�H���_: " & outDir)

    ' �o�͐�t�H���_���������܂����݂��Ă��Ȃ���������
    Call KzFile.KzEnsureFolders(outDir)
    
    ' DocTransformer�C���X�^���X�𐶐�����
    Dim DT As DocTransformer: Set DT = DocTransformerUtil.Create
    ' Word�A�v���P�[�V�����̃C���X�^���X��^����
    Dim WordApp As Word.Application: Set WordApp = CreateObject("Word.application")
    ' DocTrasnsformer������������
    Call DT.Initialize(WordApp)
    
    '=================================================================================
    ' �O���ɂ���������Excel�t�@�C����[R6�N�x]�V�[�g���J�����g�̃��[�N�u�b�N��
    ' �R�s�[����B"work�������"�V�[�g�������B���̒��g��ListObject�Ƃ��Ď��o��
    '
    Dim memberTable As ListObject
    Set memberTable = AoMemberUtils.FetchMemberTable(memberFile, "R6�N�x", ThisWorkbook)
    Debug.Print "memberTable.ListRows.Count=" & memberTable.ListRows.count
    
    ' ================================================================================
    ' �������̊e�s����������
    Dim max As Long: max = 300     '�e�X�g���ɂ͏���������(3�Ƃ�)�ɂ��đ����I��������
                                 '�{�Ԃɂ͑���������傫������(300�Ƃ�)�ɂ���
    Dim count As Long: count = 0
        
    Dim i As Long
    For i = 1 To memberTable.ListRows.count
        If i < max Then
            ' ����̎����A�����J�i�A���i�����o��
            Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "����", Trim(memberTable.ListColumns("����").DataBodyRange(i))
            dict.Add "�����J�i", Trim(memberTable.ListColumns("�����J�i").DataBodyRange(i))
            dict.Add "���i", Trim(memberTable.ListColumns("���i").DataBodyRange(i))
            Debug.Print dict("�����J�i"), dict("����"), dict("���i")
            
            ' A�����B�����C�����D�����ΏۂƂ���B
            ' �Ə������Word�����𐶐����Ȃ��B
            ' �hB�O��h�́hB"�Ɠ����A�hC�O��h�́hC�h�Ɠ����Ƃ݂Ȃ�
            If dict("���i") = "A" Or _
                StartsWith(dict("���i"), "B") Or _
                StartsWith(dict("���i"), "C") Or _
                dict("���i") = "D" Then
        
                Call dict.Add("���ishort", Left(dict("���i"), 1))
                
                If dict("���i") Like "*�O��" Then
                    Call dict.Add("�Ȃ��O��", "�Ȃ��O�O��w�����̐搶���ɂ��܂��Ăͤ�����̓c�V����Ɏx���������܂Ƃ߂Ă��������܂���ǂ��������͂��������܂��悤�X�������肢�\���グ�܂��")
                Else
                    Call dict.Add("�Ȃ��O��", "")
                End If

            
                ' �o��Word�t�@�C���̃p�X�����肵��
                Dim r As String: r = outDir & "\" & dict("�����J�i") & "_" & dict("����") & "_" & dict("���i") & ".docx"
                Debug.Print r
            
                ' Word�h�L�������g��ϊ����鏈�������s����
                Call DT.Transform(templateFile, dict, r)
    
            End If
        End If
    Next i
    
    Debug.Print "�I�����܂���"
End Sub



Public Function StartsWith(target_str As String, search_str As String) As Boolean
'###################################################################################
'target_str������search_str������Ŏn�܂��Ă��邩�m�F����
'search_str�Ŏn�܂��Ă���ꍇ��True
'search_str�Ŏn�܂��Ă��Ȃ��A��������search_str��target_str�̕������𒴂���ꍇ��False��Ԃ�
'
'��
'    StartsWith('C�O��', 'C') ��True��Ԃ�
'    StartsWith('C�O��', 'E') ��False��Ԃ�
'
'###################################################################################
  
  If Len(search_str) > Len(target_str) Then
    Exit Function
  End If
  
  If Left(target_str, Len(search_str)) = search_str Then
    StartsWith = True
  End If

End Function

