Attribute VB_Name = "MbMemberTableUtil"
Option Explicit

'
' MbMemberTableUtil --- ��������Excel���[�N�V�[�g���������p��Sub��Function��񋟂���
'

' ��������Excel�t�@�C���̃p�X���p�����[�^�Ƃ��Ď󂯂�B
' ���̂Ȃ��̉�����냏�[�N�V�[�g�̖��O��memberSheetName�p�����[�^�Ƃ��Ď󂯂�B
' ���[�N�V�[�g��ǂݎ���āA�w�肳�ꂽ���[�N�u�b�N�̂Ȃ��ɃR�s�[����B
' �o�͐�Ƃ��Ẵ��[�N�V�[�g�̖��O���w�肵�Ȃ���΃f�t�H���g�Ƃ��āuwork�������v�Ƃ���B
' �����ďo�͂��ꂽ���[�N�V�[�g�ɂɊ܂܂�Ă���Excel Table�̓��eListObject�ɕϊ����ĕԂ��B
Public Function FetchMemberTable(memberFilePath As String, _
        memberSheetName As String, _
        targetWorkbook As Workbook, _
        Optional ByVal targetSheetName As String = "work�������", _
        Optional ByVal renew As Boolean = True, _
        Optional ByVal tableId As String = "MembersTable13") As ListObject


    ' target���[�N�u�b�N�̂Ȃ��Ɂ@"work�������" �V�[�g��������΍��B
    ' ���łɃR�s�[�����݂��Ă��Ă���renew��True�Ǝw�肳��Ă�����R�s�[���㏑������B
    ' ���łɃR�s�[�����݂��Ă��Ă���renew��False�Ȃ�Ή������Ȃ��B
    If (Not BbWorksheet.IsWorksheetPresentInWorkbook(targetWorkbook, targetSheetName)) Or _
        (BbWorksheet.IsWorksheetPresentInWorkbook(targetWorkbook, targetSheetName) And renew = True) Then
        
        ' �O���ɂ��������냏�[�N�u�b�N�̃E�C���h�E���J��
        Dim sourceWorkbook As Workbook
        ' memberFilePath��targetWorkbook��Path�����Ƃ��鑊�΃p�X�Ŏw�肳��Ă��Ă����v�Ȃ悤�ɗp�S����
        Set sourceWorkbook = Workbooks.Open(BbFile.AbsolutifyPath( _
                                            targetWorkbook.Path, memberFilePath))
                                            
        '�ʂ̃��[�N�u�b�N��open�����ThisWorkbook�������I�ɐ؂�ւ���Ă��܂����ɒ��ӂ���
        
        ' �O���ɂ���������Excel�t�@�C���̃��[�N�V�[�g���J�����g��Wookbook�ɃR�s�[����
        Call BbWorksheet.FetchWorksheetFromWorkbook( _
                sourceWorkbook, memberSheetName, _
                targetWorkbook, targetSheetName)
        
        ' �O���ɂ��������냏�[�N�u�b�N��window�����B���Ȃ��ŕ����Ă����Ă͂����܂���B
        ' �l���蓮�ŃE�C���h�E����Ȃ���΂Ȃ�Ȃ����A���B���� �Ղ�������B
        Application.DisplayAlerts = False   '�u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ�
        sourceWorkbook.Close
        Application.DisplayAlerts = True
    End If
    
    ' target���[�N�u�b�N�́hwork�������h�V�[�g�̂Ȃ���Excel Table������͂�
    Dim ws As Worksheet: Set ws = targetWorkbook.Worksheets(targetSheetName)
    
    ' �e�[�u����ListObject�ɕϊ�����
    Dim tbl As ListObject: Set tbl = ws.ListObjects(tableId)
    
    ' ListObject�̌`�𐮂��� --------------------------------------------------
    
    '�񕝂𒲐�
    tbl.ListColumns("�����J�i").Range.EntireColumn.AutoFit
    'tbl.ListColumns("�Ζ��於").Range.EntireColumn.AutoFit
    'tbl.ListColumns("�ٓ�").Range.EntireColumn.AutoFit
    
    '�d�v�łȂ�����\���ɂ���B�V�[�g����������Ƃ��ɕ֗��Ȃ悤��
    'tbl.ListColumns("�N��").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("��oNo").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("��o�^��").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("�����o�^").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("�Ζ���").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("�Ζ���Z��").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("�ΐ�TELNo").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("���").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("����Z��").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("����TELNo").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("�g�єԍ�").Range.EntireColumn.Hidden = True
    
    Set FetchMemberTable = tbl
        
End Function


